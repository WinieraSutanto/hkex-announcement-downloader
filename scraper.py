"""Save HKEXnews English and Traditional Chinese Title Search results as JSON.

The site uses an undocumented internal search endpoint, not a documented public API.
Use modest request rates and check HKEX's terms before reusing or sharing data.

Example:
    python scraper.py --stock-code 00005 --from-date 2026-10-01 \
        --to-date 2026-10-29 --headline-category "Announcements and Notices"
    python scraper.py --from-date 2026-10-01 --to-date 2026-10-29

This creates separate English and Chinese JSON files. The English category
code is used for both searches; HKEXnews uses matching category codes across
the two languages. Omit --stock-code to search all stock codes.
"""

from __future__ import annotations

import argparse
import http.cookiejar
import html
import json
import re
import sys
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode, urljoin
from urllib.request import (
    HTTPCookieProcessor,
    Request,
    build_opener,
)


BASE_URL = "https://www.hkexnews.hk"
INDEX_URL = BASE_URL + "/index.htm"
TITLE_SEARCH_URL = BASE_URL + "/search/titlesearch.xhtml?lang=en"
CHINESE_INDEX_URL = BASE_URL + "/index_c.htm"
CHINESE_TITLE_SEARCH_URL = BASE_URL + "/search/titlesearch.xhtml?lang=zh-HK"
USER_AGENT = "Regevio-HKEXnews-Scraper/0.1"
TIMEOUT_SECONDS = 30
MAX_RESULTS_PER_REQUEST = 10000


class StockNotFoundError(ValueError):
    """Raised when HKEXnews does not recognize a stock code."""


def fetch_text(opener: Any, url: str) -> str:
    headers = {
        "User-Agent": USER_AGENT,
        "Accept": "text/html,application/json,application/javascript,*/*",
    }
    request = Request(url, headers=headers)
    with opener.open(request, timeout=TIMEOUT_SECONDS) as response:
        charset = response.headers.get_content_charset() or "utf-8"
        return response.read().decode(charset, errors="replace")


def fetch_bytes(opener: Any, url: str) -> bytes:
    request = Request(
        url,
        headers={
            "User-Agent": USER_AGENT,
            "Accept": "application/pdf,application/octet-stream,*/*",
        },
    )
    with opener.open(request, timeout=TIMEOUT_SECONDS) as response:
        return response.read()


def fetch_json(opener: Any, url: str) -> Any:
    return json.loads(fetch_text(opener, url))


def normalize_stock_code(value: str) -> str:
    code = value.strip()
    if not re.fullmatch(r"\d{1,5}", code):
        raise ValueError("Stock code must contain between 1 and 5 digits.")
    return code.zfill(5)


def resolve_stock(opener: Any, stock_code: str) -> Dict[str, Any]:
    normalized_code = normalize_stock_code(stock_code)
    query = urlencode(
        {
            "lang": "EN",
            "type": "A",
            "name": normalized_code,
            "market": "SEHK",
            "callback": "callback",
        }
    )
    response = fetch_text(opener, BASE_URL + "/search/prefix.do?" + query)
    match = re.fullmatch(r"\s*callback\((.*)\)\s*;?\s*", response, re.DOTALL)
    if not match:
        raise RuntimeError("HKEXnews returned an unexpected stock lookup response.")

    data = json.loads(match.group(1))
    for stock in data.get("stockInfo", []):
        if stock.get("code") == normalized_code:
            return stock
    raise StockNotFoundError(
        "Stock code {} does not exist.".format(normalized_code)
    )


def fetch_categories(
    opener: Any, language: str
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    tier_one = fetch_json(
        opener, BASE_URL + "/ncms/script/eds/tierone_{}.json".format(language)
    )
    tier_two = fetch_json(
        opener, BASE_URL + "/ncms/script/eds/tiertwo_{}.json".format(language)
    )
    return tier_one, tier_two


def fetch_category_groups(opener: Any, language: str) -> List[Dict[str, Any]]:
    return fetch_json(
        opener, BASE_URL + "/ncms/script/eds/tiertwogrp_{}.json".format(language)
    )


def verify_chinese_category(
    category: Dict[str, str],
    tier_one: List[Dict[str, Any]],
    tier_two: List[Dict[str, Any]],
) -> None:
    if category["t1code"] == "-2":
        return

    if category["t2code"] == "-2" and category["t2Gcode"] == "-2":
        matches = any(
            str(record.get("code")) == category["t1code"] for record in tier_one
        )
    elif category["t2code"] == "-2":
        matches = any(
            str(record.get("t1code")) == category["t1code"]
            and str(record.get("t2Gcode") or "-2") == category["t2Gcode"]
            for record in tier_two
        )
    else:
        matches = any(
            str(record.get("t1code")) == category["t1code"]
            and str(record.get("t2Gcode") or "-2") == category["t2Gcode"]
            and str(record.get("code")) == category["t2code"]
            for record in tier_two
        )
    if not matches:
        raise RuntimeError(
            "The selected English headline category has no matching category "
            "in HKEXnews's Chinese category list."
        )


def resolve_category(
    value: Optional[str],
    tier_one: List[Dict[str, Any]],
    tier_two: List[Dict[str, Any]],
) -> Dict[str, str]:
    no_filter = {"t1code": "-2", "t2Gcode": "-2", "t2code": "-2"}
    if not value:
        return no_filter

    selection = value.strip()
    matches: List[Tuple[str, Dict[str, Any]]] = []
    for record in tier_one:
        names_and_codes = (
            str(record.get("name", "")).casefold(),
            str(record.get("code", "")),
        )
        if selection.casefold() in names_and_codes:
            matches.append(("tier_one", record))
    for record in tier_two:
        names_and_codes = (
            str(record.get("name", "")).casefold(),
            str(record.get("code", "")),
        )
        if selection.casefold() in names_and_codes:
            matches.append(("tier_two", record))

    if not matches:
        raise ValueError(
            "Headline category {!r} was not found. Use an exact category name "
            "or category code.".format(selection)
        )
    if len(matches) > 1:
        codes = sorted({str(record.get("code")) for _, record in matches})
        raise ValueError(
            "Headline category {!r} is ambiguous; use one of these category "
            "codes: {}.".format(selection, ", ".join(codes))
        )

    level, record = matches[0]
    if level == "tier_one":
        return {"t1code": str(record["code"]), "t2Gcode": "-2", "t2code": "-2"}
    return {
        "t1code": str(record["t1code"]),
        "t2Gcode": str(record.get("t2Gcode") or "-2"),
        "t2code": str(record["code"]),
    }


def parse_date(value: str, option_name: str) -> date:
    try:
        return date.fromisoformat(value.strip().replace("/", "-"))
    except ValueError as exc:
        raise ValueError(
            "{} must be a date in YYYY-MM-DD or YYYY/MM/DD format.".format(option_name)
        ) from exc


def parse_search_response(response_text: str) -> Tuple[int, List[Dict[str, Any]]]:
    response = json.loads(response_text)
    if not isinstance(response, dict) or "result" not in response:
        raise RuntimeError("HKEXnews returned an unexpected search response.")
    if not isinstance(response.get("result"), str):
        raise RuntimeError("HKEXnews returned results in an unexpected format.")
    total_records = int(response["recordCnt"])
    raw_results = json.loads(response["result"])
    if not isinstance(raw_results, list):
        raise RuntimeError("HKEXnews returned results in an unexpected format.")

    results: List[Dict[str, Any]] = []
    for row in raw_results:
        if not isinstance(row, dict):
            raise RuntimeError("HKEXnews returned a result row in an unexpected format.")
        release_time = str(row.get("DATE_TIME") or "")
        try:
            published_at = datetime.strptime(release_time, "%d/%m/%Y %H:%M")
        except ValueError as exc:
            raise RuntimeError(
                "HKEXnews returned an unrecognized publication date: {!r}".format(
                    release_time
                )
            ) from exc

        category = str(row.get("LONG_TEXT") or row.get("SHORT_TEXT") or "")
        category = html.unescape(re.sub(r"<[^>]*>", " ", category))
        category = " ".join(category.split())
        file_link = row.get("FILE_LINK")

        results.append(
            {
                "publication_date": published_at.date().isoformat(),
                "release_time": release_time,
                "stock_code": row.get("STOCK_CODE"),
                "stock_name": row.get("STOCK_NAME"),
                "headline_category": category,
                "headline": row.get("TITLE"),
                "document_url": urljoin(BASE_URL, file_link) if file_link else None,
                "file_size": row.get("FILE_INFO"),
                "news_id": row.get("NEWS_ID"),
            }
        )
    return total_records, results


def search(
    opener: Any,
    stock: Dict[str, Any],
    start_date: date,
    end_date: date,
    category: Dict[str, str],
    language: str,
) -> Tuple[int, List[Dict[str, Any]]]:
    endpoint = BASE_URL + "/search/titleSearchServlet.do?"

    def search_range(
        batch_start: date, batch_end: date
    ) -> Tuple[int, List[Dict[str, Any]]]:
        params = {
            "sortDir": "0",
            "sortByOptions": "DateTime",
            "category": "0",
            "market": "SEHK",
            "stockId": str(stock.get("stockId") or ""),
            "documentType": "-1",
            "fromDate": batch_start.strftime("%Y%m%d"),
            "toDate": batch_end.strftime("%Y%m%d"),
            "title": "",
            "searchType": "0",
            "t1code": category["t1code"],
            "t2Gcode": category["t2Gcode"],
            "t2code": category["t2code"],
            "rowRange": "100",
            "lang": language,
        }
        response_text = fetch_text(opener, endpoint + urlencode(params))
        total_records, results = parse_search_response(response_text)

        if total_records > MAX_RESULTS_PER_REQUEST:
            if batch_start == batch_end:
                raise RuntimeError(
                    "HKEXnews reported {} results for {} on {}, exceeding its "
                    "{}-result request limit.".format(
                        total_records,
                        language,
                        batch_start.isoformat(),
                        MAX_RESULTS_PER_REQUEST,
                    )
                )

            midpoint = batch_start + (batch_end - batch_start) // 2
            earlier_total, earlier_results = search_range(batch_start, midpoint)
            later_total, later_results = search_range(
                midpoint + timedelta(days=1), batch_end
            )
            return (
                earlier_total + later_total,
                later_results + earlier_results,
            )

        if total_records > len(results):
            params["rowRange"] = str(total_records)
            response_text = fetch_text(opener, endpoint + urlencode(params))
            returned_total, results = parse_search_response(response_text)
            if returned_total != total_records:
                raise RuntimeError(
                    "HKEXnews reported a different result count while loading "
                    "all results."
                )

        if len(results) != total_records:
            raise RuntimeError(
                "HKEXnews returned {} of {} results for {} to {} in {}.".format(
                    len(results),
                    total_records,
                    batch_start.isoformat(),
                    batch_end.isoformat(),
                    language,
                )
            )
        return total_records, results

    return search_range(start_date, end_date)


def build_document(
    stock: Dict[str, Any],
    start_date: date,
    end_date: date,
    headline_category: Optional[str],
    total_records: int,
    results: List[Dict[str, Any]],
    chinese: bool = False,
) -> Dict[str, Any]:
    stock_name = stock["name"]
    if chinese and stock["code"] and results and results[0].get("stock_name"):
        stock_name = results[0]["stock_name"]

    return {
        "source": (
            "HKEXnews Title Search (Traditional Chinese)"
            if chinese
            else "HKEXnews Title Search"
        ),
        "retrieved_at": datetime.now(timezone.utc).isoformat(),
        "search_criteria": {
            "stock_code": stock["code"],
            "stock_name": stock_name,
            "from_date": start_date.isoformat(),
            "to_date": end_date.isoformat(),
            "headline_category": headline_category or "ALL",
        },
        "total_records_found": total_records,
        "results_returned": len(results),
        "results_complete": len(results) >= total_records,
        "results": results,
    }


def run_search(
    stock_code: Optional[str],
    start_date: date,
    end_date: date,
    headline_category: Optional[str] = None,
    category: Optional[Dict[str, str]] = None,
) -> Tuple[Dict[str, Any], Dict[str, Any]]:
    if start_date > end_date:
        raise ValueError("From date must not be later than to date.")
    if end_date > date.today():
        raise ValueError("--to-date must not be in the future.")

    opener = build_opener(HTTPCookieProcessor(http.cookiejar.CookieJar()))
    fetch_text(opener, INDEX_URL)
    fetch_text(opener, TITLE_SEARCH_URL)
    fetch_text(opener, CHINESE_INDEX_URL)
    fetch_text(opener, CHINESE_TITLE_SEARCH_URL)
    normalized_stock_code = (stock_code or "").strip()
    if normalized_stock_code:
        stock = resolve_stock(opener, normalized_stock_code)
    else:
        stock = {"code": "", "name": "All stocks", "stockId": ""}
    tier_one_en, tier_two_en = fetch_categories(opener, "e")
    tier_one_zh, tier_two_zh = fetch_categories(opener, "c")
    selected_category = category or resolve_category(
        headline_category, tier_one_en, tier_two_en
    )
    verify_chinese_category(selected_category, tier_one_zh, tier_two_zh)
    total_records_en, results_en = search(
        opener, stock, start_date, end_date, selected_category, "E"
    )
    total_records_zh, results_zh = search(
        opener, stock, start_date, end_date, selected_category, "ZH"
    )

    return (
        build_document(
            stock,
            start_date,
            end_date,
            headline_category,
            total_records_en,
            results_en,
        ),
        build_document(
            stock,
            start_date,
            end_date,
            headline_category,
            total_records_zh,
            results_zh,
            chinese=True,
        ),
    )


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Search HKEXnews announcements in English and Traditional Chinese "
            "and save separate JSON files."
        )
    )
    parser.add_argument(
        "--stock-code",
        help="Stock code, e.g. 00005; omit to search all stock codes",
    )
    parser.add_argument(
        "--from-date", required=True, help="Start date (YYYY-MM-DD or YYYY/MM/DD)"
    )
    parser.add_argument(
        "--to-date", required=True, help="End date (YYYY-MM-DD or YYYY/MM/DD)"
    )
    parser.add_argument(
        "--headline-category",
        help=(
            "Exact English headline category name or category code; the matching "
            "category is applied to both languages"
        ),
    )
    parser.add_argument(
        "--output",
        default="hkexnews_results.json",
        help="Output JSON path (default: hkexnews_results.json)",
    )
    parser.add_argument(
        "--chinese-output",
        default="hkexnews_results_zh.json",
        help="Traditional Chinese JSON path (default: hkexnews_results_zh.json)",
    )
    return parser


def main() -> int:
    parser = build_argument_parser()
    args = parser.parse_args()
    output = Path(args.output)
    chinese_output = Path(args.chinese_output)
    try:
        if output.resolve() == chinese_output.resolve():
            raise ValueError("--output and --chinese-output must be different paths.")

        start_date = parse_date(args.from_date, "--from-date")
        end_date = parse_date(args.to_date, "--to-date")
        if start_date > end_date:
            raise ValueError("--from-date must not be later than --to-date.")
        if end_date > date.today():
            raise ValueError("--to-date must not be in the future.")

        document_en, document_zh = run_search(
            args.stock_code,
            start_date,
            end_date,
            args.headline_category,
        )
    except (HTTPError, URLError, OSError, ValueError, RuntimeError) as exc:
        parser.error(str(exc))

    try:
        for output_path, document in (
            (output, document_en),
            (chinese_output, document_zh),
        ):
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_text(
                json.dumps(document, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
    except OSError as exc:
        parser.error("Could not write output JSON: {}".format(exc))

    print(
        "Saved {} result(s) ({} total reported) to {}".format(
            document_en["results_returned"],
            document_en["total_records_found"],
            output,
        )
    )
    print(
        "Saved {} Traditional Chinese result(s) ({} total reported) to {}".format(
            document_zh["results_returned"],
            document_zh["total_records_found"],
            chinese_output,
        )
    )
    if not document_en["results_complete"]:
        print(
            "Note: HKEXnews returned only part of the matching English results; "
            "the English JSON indicates results_complete=false.",
            file=sys.stderr,
        )
    if not document_zh["results_complete"]:
        print(
            "Note: HKEXnews returned only part of the matching Chinese results; "
            "the Chinese JSON indicates results_complete=false.",
            file=sys.stderr,
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())