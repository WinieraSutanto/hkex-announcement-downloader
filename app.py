import http.cookiejar
import calendar
from collections import defaultdict, deque
import io
import html
import json
import re
import zipfile
from datetime import date, datetime
from typing import Any, Deque, Dict, List, Optional, Tuple
from urllib.error import HTTPError, URLError
from urllib.request import HTTPCookieProcessor, build_opener

from openpyxl import Workbook
import streamlit as st
import streamlit_antd_components as sac

from scraper import (
    StockNotFoundError,
    fetch_bytes,
    fetch_categories,
    fetch_category_groups,
    run_search,
)


st.set_page_config(
    page_title="HKEXnews Search",
    page_icon="📄",
    layout="wide",
)

st.title("HKEXnews Search")
st.write(
    "Search HKEXnews in English and Traditional Chinese. "
    "Choose an English headline category; it will be applied to both searches."
)


@st.cache_data(ttl=3600)
def load_category_tree() -> List[Dict[str, Any]]:
    opener = build_opener(HTTPCookieProcessor(http.cookiejar.CookieJar()))
    tier_one, tier_two = fetch_categories(opener, "e")
    tier_two_groups = fetch_category_groups(opener, "e")
    children_by_parent: Dict[str, List[Dict[str, Any]]] = {}
    for record in tier_two:
        children_by_parent.setdefault(str(record["t1code"]), []).append(record)
    groups_by_parent: Dict[str, List[Dict[str, Any]]] = {}
    for record in tier_two_groups:
        groups_by_parent.setdefault(str(record["t1code"]), []).append(record)

    tree: List[Dict[str, Any]] = [
        {"label": "All categories", "value": "-2"}
    ]
    for record in tier_one:
        tier_one_code = str(record["code"])
        children: List[Dict[str, Any]] = [
            {"label": "All categories", "value": "__all__"}
        ]
        children_by_group: Dict[str, List[Dict[str, Any]]] = {}
        for child in children_by_parent.get(tier_one_code, []):
            children_by_group.setdefault(str(child["t2Gcode"]), []).append(child)
        for group in groups_by_parent.get(tier_one_code, []):
            group_code = str(group["code"])
            group_children: List[Dict[str, Any]] = [
                {"label": "All categories", "value": "__all__"}
            ]
            for child in children_by_group.get(group_code, []):
                group_children.append(
                    {"label": str(child["name"]), "value": str(child["code"])}
                )
            children.append(
                {
                    "label": str(group["name"]),
                    "value": group_code,
                    "children": group_children,
                }
            )
        tree.append(
            {
                "label": str(record["name"]),
                "value": tier_one_code,
                "children": children,
            }
        )
    return tree


def category_from_path(
    selected_path: List[str],
    category_tree: List[Dict[str, Any]],
) -> Dict[str, str] | None:
    if not selected_path or selected_path == ["All categories"]:
        return None

    level = category_tree
    selected_nodes: List[Dict[str, Any]] = []
    for label in selected_path:
        node = next((item for item in level if item["label"] == label), None)
        if node is None:
            return None
        selected_nodes.append(node)
        level = node.get("children", [])

    codes = [str(node["value"]) for node in selected_nodes]
    if len(codes) == 1:
        return {"t1code": codes[0], "t2Gcode": "-2", "t2code": "-2"}
    if len(codes) == 2:
        if codes[1] == "__all__":
            return {"t1code": codes[0], "t2Gcode": "-2", "t2code": "-2"}
        return {"t1code": codes[0], "t2Gcode": codes[1], "t2code": "-2"}
    if codes[2] == "__all__":
        return {"t1code": codes[0], "t2Gcode": codes[1], "t2code": "-2"}
    return {"t1code": codes[0], "t2Gcode": codes[1], "t2code": codes[2]}


def normalize_selected_path(selected_path: Any) -> List[str]:
    if not selected_path:
        return []
    if isinstance(selected_path, list):
        return [str(value) for value in selected_path]
    return [str(selected_path)]


def date_selector(
    label: str,
    key_prefix: str,
    min_date: date,
    max_date: date,
) -> date:
    st.markdown("**{}**".format(label))
    year_key = "{}_year".format(key_prefix)
    month_key = "{}_month".format(key_prefix)
    day_key = "{}_day".format(key_prefix)

    years = list(range(max_date.year, min_date.year - 1, -1))
    if year_key not in st.session_state or st.session_state[year_key] not in years:
        st.session_state[year_key] = max_date.year

    year_column, month_column, day_column = st.columns(3)
    year = year_column.selectbox("Year", years, key=year_key)

    first_month = min_date.month if year == min_date.year else 1
    last_month = max_date.month if year == max_date.year else 12
    months = list(range(first_month, last_month + 1))
    default_month = max_date.month if year == max_date.year else first_month
    if month_key not in st.session_state or st.session_state[month_key] not in months:
        st.session_state[month_key] = default_month
    month = month_column.selectbox("Month", months, key=month_key)

    first_day = (
        min_date.day
        if year == min_date.year and month == min_date.month
        else 1
    )
    last_day = calendar.monthrange(year, month)[1]
    if year == max_date.year and month == max_date.month:
        last_day = min(last_day, max_date.day)
    days = list(range(first_day, last_day + 1))
    default_day = (
        max_date.day
        if year == max_date.year and month == max_date.month
        else first_day
    )
    if day_key not in st.session_state or st.session_state[day_key] not in days:
        st.session_state[day_key] = default_day
    selected_day = day_column.selectbox(
        "Day",
        days,
        key=day_key,
        format_func=lambda value: "{:02d}".format(value),
    )

    return date(year, month, selected_day)


try:
    category_tree = load_category_tree()
except (HTTPError, URLError, OSError, ValueError, RuntimeError) as exc:
    st.error("Could not load headline categories: {}".format(exc))
    st.stop()

earliest_date = date(1999, 4, 1)
latest_date = date.today()

stock_code = st.text_input(
    "Stock Code (Mandatory if period is more than 1 month)",
    value="",
    placeholder="e.g. 00005",
    help=(
        "Leave blank to search all stock codes for date ranges of 31 days or less. "
        "Enter exactly 5 digits otherwise."
    ),
)
from_date = date_selector(
    "From date",
    "from_date",
    earliest_date,
    latest_date,
)
to_date = date_selector(
    "To date",
    "to_date",
    earliest_date,
    latest_date,
)
st.markdown("**Headline category**")
selected_path = normalize_selected_path(
    sac.cascader(
        items=category_tree,
        placeholder="All categories",
        index=0,
        clear=True,
    )
)
submitted = st.button("Search HKEXnews", type="primary")
stock_code_value = stock_code.strip()
if submitted:
    if stock_code_value and not re.fullmatch(r"[0-9]{5}", stock_code_value):
        st.warning("Stock code must be exactly 5 digits (for example, 00005).")
        submitted = False
    elif not stock_code_value and (to_date - from_date).days > 31:
        st.warning(
            "A stock code is required when the date range is longer than 31 days."
        )
        submitted = False


def combine_results(
    english_document: Dict[str, Any],
    chinese_document: Dict[str, Any],
) -> List[Dict[str, Any]]:
    english_results = english_document["results"]
    chinese_results = chinese_document["results"]

    def result_key(entry: Dict[str, Any]) -> Tuple[Tuple[str, ...], str, str]:
        stock_code = html.unescape(str(entry.get("stock_code") or ""))
        stock_codes = tuple(
            sorted(
                {
                    code.zfill(5) if code.isdigit() else code.casefold()
                    for code in (
                        part.strip()
                        for part in re.split(
                            r"<br\s*/?>|[,;]", stock_code, flags=re.IGNORECASE
                        )
                    )
                    if code
                }
            )
        )

        try:
            publication_date = date.fromisoformat(
                str(entry["publication_date"]).strip()
            ).isoformat()
            release_time = datetime.strptime(
                str(entry["release_time"]).strip(), "%d/%m/%Y %H:%M"
            ).strftime("%H:%M")
        except (KeyError, TypeError, ValueError) as exc:
            raise RuntimeError(
                "A search result has an invalid publication date or time."
            ) from exc

        return stock_codes, publication_date, release_time

    chinese_by_key: Dict[
        Tuple[Tuple[str, ...], str, str], Deque[int]
    ] = defaultdict(deque)
    for index, chinese in enumerate(chinese_results):
        chinese_by_key[result_key(chinese)].append(index)

    combined: List[Dict[str, Any]] = []
    matched_chinese_indices = set()

    def append_result(
        english: Optional[Dict[str, Any]],
        chinese: Optional[Dict[str, Any]],
    ) -> None:
        metadata_entry = english if english is not None else chinese
        if metadata_entry is None:
            raise RuntimeError("Cannot combine an empty English/Chinese result pair.")

        combined.append(
            {
                "company_name": metadata_entry.get("stock_name"),
                "publication_date": metadata_entry.get("publication_date"),
                "release_time": metadata_entry.get("release_time"),
                "stock_code": metadata_entry.get("stock_code"),
                "headline_category": metadata_entry.get("headline_category"),
                "headline_category_english": (
                    english.get("headline_category") if english is not None else None
                ),
                "metadata_language": (
                    "English" if english is not None else "Traditional Chinese"
                ),
                "headline_english": (
                    english.get("headline") if english is not None else None
                ),
                "headline_chinese": (
                    chinese.get("headline") if chinese is not None else None
                ),
                "english_entry_available": english is not None,
                "chinese_entry_available": chinese is not None,
                "english_document_url": (
                    english.get("document_url") if english is not None else None
                ),
                "chinese_document_url": (
                    chinese.get("document_url") if chinese is not None else None
                ),
                "english_file_size": (
                    english.get("file_size") if english is not None else None
                ),
                "chinese_file_size": (
                    chinese.get("file_size") if chinese is not None else None
                ),
            }
        )

    for english in english_results:
        candidates = chinese_by_key.get(result_key(english))
        chinese = None
        if candidates:
            chinese_index = candidates.popleft()
            matched_chinese_indices.add(chinese_index)
            chinese = chinese_results[chinese_index]
        append_result(english, chinese)

    for index, chinese in enumerate(chinese_results):
        if index not in matched_chinese_indices:
            append_result(None, chinese)

    return combined


def document_title(entry: Dict[str, Any]) -> str:
    title = str(entry["headline"] or "document")
    title = re.sub(r"[<>:\"/\\|?*]+", "-", title)
    title = re.sub(r"\s+", "-", title).strip(".-")
    title = re.sub(r"-{2,}", "-", title)
    return title or "document"


def document_stock_code(entry: Dict[str, Any]) -> str:
    stock_code = html.unescape(str(entry.get("stock_code") or ""))
    stock_code = re.sub(r"<br\s*/?>", "-", stock_code, flags=re.IGNORECASE)
    stock_code = re.sub(r"[^A-Za-z0-9-]+", "-", stock_code).strip("-")
    if stock_code.isdigit():
        stock_code = stock_code.zfill(5)
    return stock_code or "all-stocks"


def document_filename(language: str, entry: Dict[str, Any]) -> str:
    return "{}_{}-{}-{}.pdf".format(
        language,
        entry["publication_date"],
        document_stock_code(entry),
        document_title(entry),
    )


def archive_filename(
    english_entry: Dict[str, Any],
    chinese_entry: Dict[str, Any],
) -> str:
    return "{}-{}-{}-{}.zip".format(
        english_entry["publication_date"],
        document_stock_code(english_entry),
        document_title(english_entry),
        document_title(chinese_entry),
    )


def selected_archive_filename(language: Optional[str], stock_code: str) -> str:
    filename = ""
    if language:
        filename = language
    if stock_code:
        stock_code = document_stock_code({"stock_code": stock_code})
        filename = "{}_{}".format(filename, stock_code) if filename else stock_code
    if filename:
        filename += "-"
    return filename + "ANNOUNCEMENTS.zip"


def is_pdf_url(url: Any) -> bool:
    return (
        isinstance(url, str)
        and url.lower().split("#", 1)[0].split("?", 1)[0].endswith(".pdf")
    )


def is_excel_url(url: Any) -> bool:
    return (
        isinstance(url, str)
        and url.lower()
        .split("#", 1)[0]
        .split("?", 1)[0]
        .endswith((".xls", ".xlsx", ".xlsm", ".xlsb"))
    )


def is_multi_file_entry(entry: Dict[str, Any]) -> bool:
    document_name = "{} {} {}".format(
        entry.get("headline") or "",
        entry.get("document_url") or "",
        entry.get("file_size") or "",
    )
    return (
        re.search(
            r"(?:\(\s*(?:Multi-Files|多檔案)\s*\)|Multi-Files|多檔案)",
            document_name,
            re.IGNORECASE,
        )
        is not None
    )


def is_unavailable_html_entry(entry: Dict[str, Any]) -> bool:
    file_size = str(entry.get("file_size") or "")
    url = str(entry.get("document_url") or "").lower().split("#", 1)[0].split("?", 1)[0]
    return (
        "1kb" in file_size.casefold()
        and (
            "htm" in file_size.casefold()
            or url.endswith((".htm", ".html"))
        )
    )


def is_downloadable_pdf(entry: Dict[str, Any]) -> bool:
    return (
        not is_multi_file_entry(entry)
        and not is_unavailable_html_entry(entry)
        and is_pdf_url(entry.get("document_url"))
    )


def is_excel_only_entry(entry: Dict[str, Any]) -> bool:
    return not is_multi_file_entry(entry) and is_excel_url(entry.get("document_url"))


def link_with_display_text(url: Any, display_text: str) -> str:
    if not isinstance(url, str) or not url:
        return ""
    return "{}#{}".format(url.split("#", 1)[0], display_text)


def format_multivalue(value: Any) -> str:
    text = html.unescape(str(value or ""))
    text = re.sub(r"\s*<br\s*/?>\s*", ", ", text, flags=re.IGNORECASE)
    text = re.sub(r"\s*,\s*,+\s*", ", ", text)
    return text.strip(" ,")


def matches_title_filter(result: Dict[str, Any], query: str) -> bool:
    normalized_query = query.casefold()
    return not normalized_query or any(
        normalized_query in str(result.get(title_key) or "").casefold()
        for title_key in ("headline_english", "headline_chinese")
    )


def custom_download_filename(default_filename: str, custom_filename: str) -> str:
    if not custom_filename.strip():
        return default_filename

    _, extension = default_filename.rsplit(".", 1)
    custom_filename = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "-", custom_filename)
    custom_filename = custom_filename.strip(" .")
    if not custom_filename:
        return default_filename
    if custom_filename.casefold().endswith("." + extension.casefold()):
        custom_filename = custom_filename[: -(len(extension) + 1)]
    return "{}.{}".format(custom_filename, extension)


def download_button_with_custom_filename(
    label: str,
    data: Any,
    file_name: str,
    mime: str,
    key: str,
    disabled: bool = False,
) -> None:
    custom_filename_key = "custom_filename_{}".format(key)
    st.download_button(
        label,
        data=data,
        file_name=custom_download_filename(
            file_name,
            st.session_state.get(custom_filename_key, ""),
        ),
        mime=mime,
        key=key,
        disabled=disabled,
        on_click="ignore",
    )
    st.text_input(
        "Custom file name",
        placeholder="Input custom file name here",
        key=custom_filename_key,
    )


def combined_result_table_row(
    result: Dict[str, Any],
    labeled_links: bool = False,
) -> Dict[str, Any]:
    english_entry = selected_document_entry(result, "EN")
    chinese_entry = selected_document_entry(result, "TC")
    english_multi_file = is_multi_file_entry(english_entry)
    chinese_multi_file = is_multi_file_entry(chinese_entry)
    english_unavailable = is_unavailable_html_entry(english_entry)
    chinese_unavailable = is_unavailable_html_entry(chinese_entry)
    english_available = is_downloadable_pdf(english_entry)
    chinese_available = is_downloadable_pdf(chinese_entry)
    english_multiple_link = (
        english_multi_file
        and not english_unavailable
        and isinstance(english_entry["document_url"], str)
        and bool(english_entry["document_url"])
    )
    chinese_multiple_link = (
        chinese_multi_file
        and not chinese_unavailable
        and isinstance(chinese_entry["document_url"], str)
        and bool(chinese_entry["document_url"])
    )
    english_title = (
        "Match Uncertain/Corresponding Version not Found"
        if not result["english_entry_available"]
        else "There are multiple files"
        if english_multi_file and not english_unavailable
        else "No PDF exists. Only Excel files"
        if is_excel_only_entry(english_entry)
        else (result["headline_english"] or "Headline not available")
        if english_available
        else "Document not Available"
    )
    chinese_title = (
        "Match Uncertain/Corresponding Version not Found"
        if not result["chinese_entry_available"]
        else "There are multiple files"
        if chinese_multi_file and not chinese_unavailable
        else "No PDF exists. Only Excel files"
        if is_excel_only_entry(chinese_entry)
        else (result["headline_chinese"] or "Headline not available")
        if chinese_available
        else "Document not Available"
    )
    english_link = (
        link_with_display_text(
            english_entry["document_url"],
            "Link to Multiple PDFs" if english_multiple_link else "Link to PDF",
        )
        if english_multiple_link or english_available
        else ""
    )
    chinese_link = (
        link_with_display_text(
            chinese_entry["document_url"],
            "Link to Multiple PDFs" if chinese_multiple_link else "Link to PDF",
        )
        if chinese_multiple_link or chinese_available
        else ""
    )
    if not labeled_links:
        english_link = str(english_entry["document_url"] or "") if (
            english_multiple_link or english_available
        ) else ""
        chinese_link = str(chinese_entry["document_url"] or "") if (
            chinese_multiple_link or chinese_available
        ) else ""

    return {
        "Select": False,
        "Company name (English when available)": format_multivalue(
            result["company_name"]
        ),
        "Publication date": str(result["publication_date"] or ""),
        "Publication time": str(result["release_time"] or ""),
        "Stock code": format_multivalue(result["stock_code"]),
        "Headline category (English when available)": str(
            result["headline_category"] or ""
        ),
        "English title": str(english_title),
        "Chinese title": str(chinese_title),
        "English document": english_link,
        "Chinese document": chinese_link,
    }


def create_excel_file(combined_results: List[Dict[str, Any]]) -> bytes:
    headers = [
        "Company name (English when available)",
        "Publication date",
        "Publication time",
        "Stock code",
        "Headline category (English when available)",
        "English title",
        "Chinese title",
        "English document",
        "Chinese document",
    ]
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Announcements"
    worksheet.append(headers)
    for result in combined_results:
        row = combined_result_table_row(result)
        worksheet.append([row[header] for header in headers])

    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


def announcements_filename(search_criteria: Dict[str, Any]) -> str:
    start_date = date.fromisoformat(str(search_criteria["from_date"]))
    end_date = date.fromisoformat(str(search_criteria["to_date"]))
    stock_code_value = search_criteria.get("stock_code")
    stock_code = (
        document_stock_code({"stock_code": stock_code_value})
        if stock_code_value
        else ""
    )
    filename = "{}_{}-".format(start_date.isoformat(), end_date.isoformat())
    if stock_code:
        filename += "{}-".format(stock_code)
    return filename + "ANNOUNCEMENTS"


@st.cache_data(ttl=3600, show_spinner=False)
def load_document(url: str) -> bytes:
    opener = build_opener(HTTPCookieProcessor(http.cookiejar.CookieJar()))
    return fetch_bytes(opener, url)


def download_entry_documents(
    english_entry: Dict[str, Any],
    chinese_entry: Dict[str, Any],
    download_key: str,
) -> None:
    english_url = english_entry["document_url"]
    chinese_url = chinese_entry["document_url"]
    english_pdf = is_downloadable_pdf(english_entry)
    chinese_pdf = is_downloadable_pdf(chinese_entry)
    english_column, chinese_column, zip_column = st.columns(3)

    if english_pdf:
        with english_column:
            download_button_with_custom_filename(
                "Download English",
                lambda url=english_url: load_document(url),
                document_filename("EN", english_entry),
                "application/pdf",
                "english_download_{}".format(download_key),
            )
    else:
        english_column.caption("Document not Available")

    if chinese_pdf:
        with chinese_column:
            download_button_with_custom_filename(
                "Download Chinese",
                lambda url=chinese_url: load_document(url),
                document_filename("TC", chinese_entry),
                "application/pdf",
                "chinese_download_{}".format(download_key),
            )
    else:
        chinese_column.caption("Document not Available")

    if english_pdf and chinese_pdf:
        with zip_column:
            def create_archive() -> bytes:
                archive = io.BytesIO()
                with zipfile.ZipFile(
                    archive, "w", compression=zipfile.ZIP_DEFLATED
                ) as zip_file:
                    zip_file.writestr(
                        document_filename("EN", english_entry),
                        load_document(english_url),
                    )
                    zip_file.writestr(
                        document_filename("TC", chinese_entry),
                        load_document(chinese_url),
                    )

                return archive.getvalue()

            download_button_with_custom_filename(
                "Download both",
                create_archive,
                archive_filename(english_entry, chinese_entry),
                "application/zip",
                "zip_download_{}".format(download_key),
            )
    else:
        zip_column.caption("Both PDFs are not available")


def selected_document_entry(
    result: Dict[str, Any],
    language: str,
) -> Dict[str, Any]:
    is_english = language == "EN"
    return {
        "document_url": (
            result["english_document_url"]
            if is_english
            else result["chinese_document_url"]
        ),
        "headline": (
            result["headline_english"] if is_english else result["headline_chinese"]
        ),
        "publication_date": result["publication_date"],
        "stock_code": result["stock_code"],
        "file_size": (
            result.get("english_file_size")
            if is_english
            else result.get("chinese_file_size")
        ),
    }


def create_selected_archive(
    selected_results: List[Dict[str, Any]],
    language: Optional[str],
) -> bytes:
    archive = io.BytesIO()
    used_filenames: Dict[str, int] = {}
    with zipfile.ZipFile(archive, "w", compression=zipfile.ZIP_DEFLATED) as zip_file:
        for result in selected_results:
            languages = [language] if language else ["EN", "TC"]
            for document_language in languages:
                entry = selected_document_entry(result, document_language)
                if not is_downloadable_pdf(entry):
                    continue

                filename = document_filename(document_language, entry)
                duplicate_count = used_filenames.get(filename.casefold(), 0)
                used_filenames[filename.casefold()] = duplicate_count + 1
                if duplicate_count:
                    stem, extension = filename.rsplit(".", 1)
                    filename = "{}-{}.{}".format(
                        stem,
                        duplicate_count + 1,
                        extension,
                    )
                zip_file.writestr(
                    filename,
                    load_document(entry["document_url"]),
                )
    return archive.getvalue()


def show_combined_results(
    combined_results: List[Dict[str, Any]],
    table_key: str,
    export_filename: str,
    search_stock_code: str,
) -> None:
    with st.container(border=True):
        st.subheader("Combined English and Traditional Chinese results")
        st.caption("{} result(s) returned".format(len(combined_results)))
        if not combined_results:
            st.info("No matching entries were found.")
            return

        filter_input_key = "{}_title_filter_input".format(table_key)
        applied_filter_key = "{}_title_filter".format(table_key)
        page_widget_key = "{}_page".format(table_key)
        filter_input = st.text_input(
            "Title keyword",
            placeholder="Please input title for filtered search",
            key=filter_input_key,
        )
        if st.button("Filter", key="{}_title_filter_button".format(table_key)):
            st.session_state[applied_filter_key] = filter_input.strip()
            st.session_state[page_widget_key] = 0

        title_filter = st.session_state.get(applied_filter_key, "").casefold()
        filtered_results = [
            (index, result)
            for index, result in enumerate(combined_results)
            if matches_title_filter(result, title_filter)
        ]
        if not filtered_results:
            st.info("No results match the title filter.")
            return

        page_size = 50
        page_count = (len(filtered_results) + page_size - 1) // page_size
        st.session_state[page_widget_key] = min(
            st.session_state.get(page_widget_key, 0),
            page_count - 1,
        )
        page_index = st.session_state[page_widget_key]

        def set_page(page: int) -> None:
            st.session_state[page_widget_key] = page

        def set_selected_download_page(key: str, page: int) -> None:
            st.session_state[key] = page

        page_start = page_index * page_size
        page_end = min(page_start + page_size, len(filtered_results))

        selected_indices_key = "{}_selected_indices".format(table_key)
        editor_revision_key = "{}_editor_revision".format(table_key)
        if editor_revision_key not in st.session_state:
            st.session_state[editor_revision_key] = 0

        def select_results(result_indices: List[int]) -> None:
            selected_indices = set(st.session_state.get(selected_indices_key, []))
            selected_indices.update(result_indices)
            st.session_state[selected_indices_key] = sorted(selected_indices)
            st.session_state[editor_revision_key] += 1

        def deselect_results(result_indices: List[int]) -> None:
            selected_indices = set(st.session_state.get(selected_indices_key, []))
            selected_indices.difference_update(result_indices)
            st.session_state[selected_indices_key] = sorted(selected_indices)
            st.session_state[editor_revision_key] += 1

        def update_selected_indices(
            selected_key: str,
            editor_key: str,
            result_indices: List[int],
        ) -> None:
            editor_state = st.session_state.get(editor_key, {})
            edited_rows = editor_state.get("edited_rows", {})
            selected_indices = set(st.session_state.get(selected_key, []))
            for row_index, changes in edited_rows.items():
                if "Select" not in changes:
                    continue
                result_index = result_indices[int(row_index)]
                if changes["Select"]:
                    selected_indices.add(result_index)
                else:
                    selected_indices.discard(result_index)
            st.session_state[selected_key] = sorted(selected_indices)

        page_result_indices = [
            filtered_results[index][0] for index in range(page_start, page_end)
        ]
        selected_indices = set(st.session_state.get(selected_indices_key, []))
        table_rows = [
            {
                **combined_result_table_row(result, labeled_links=True),
                "Select": result_index in selected_indices,
            }
            for result_index, result in filtered_results[page_start:page_end]
        ]

        editor_key = "{}_editor_{}_{}".format(
            table_key,
            page_index,
            st.session_state[editor_revision_key],
        )
        st.data_editor(
            table_rows,
            hide_index=True,
            width="stretch",
            key=editor_key,
            on_change=update_selected_indices,
            args=(
                selected_indices_key,
                editor_key,
                page_result_indices,
            ),
            column_order=[
                "Select",
                "Company name (English when available)",
                "Publication date",
                "Publication time",
                "Stock code",
                "Headline category (English when available)",
                "English title",
                "Chinese title",
                "English document",
                "Chinese document",
            ],
            column_config={
                "Select": st.column_config.CheckboxColumn(
                    "Select",
                    help="Select entries to show their download controls below.",
                    default=False,
                ),
                "English document": st.column_config.LinkColumn(
                    "English document",
                    display_text=r"#(Link to (?:PDF|Multiple PDFs))",
                    disabled=True,
                ),
                "Chinese document": st.column_config.LinkColumn(
                    "Chinese document",
                    display_text=r"#(Link to (?:PDF|Multiple PDFs))",
                    disabled=True,
                ),
            },
            disabled=[
                "Company name (English when available)",
                "Publication date",
                "Publication time",
                "Stock code",
                "Headline category (English when available)",
                "English title",
                "Chinese title",
                "English document",
                "Chinese document",
            ],
        )

        (
            page_select_column,
            page_deselect_column,
            table_select_column,
            table_deselect_column,
            _,
        ) = st.columns([1, 1, 1, 1, 2])
        with page_select_column:
            st.button(
                "Select all on this page",
                key="select_page_{}".format(table_key),
                on_click=select_results,
                args=(page_result_indices,),
            )
        with page_deselect_column:
            st.button(
                "Deselect all on this page",
                key="deselect_page_{}".format(table_key),
                on_click=deselect_results,
                args=(page_result_indices,),
            )
        with table_select_column:
            st.button(
                "Select all in table",
                key="select_table_{}".format(table_key),
                on_click=select_results,
                args=([index for index, _ in filtered_results],),
            )
        with table_deselect_column:
            st.button(
                "Deselect all in table",
                key="deselect_table_{}".format(table_key),
                on_click=deselect_results,
                args=([index for index, _ in filtered_results],),
            )

        page_selector_column, previous_column, next_column = st.columns(
            [4, 1, 1],
            vertical_alignment="bottom",
        )
        with page_selector_column:
            page_index = st.selectbox(
                "Results page",
                options=list(range(page_count)),
                format_func=lambda page: "Page {} of {}".format(page + 1, page_count),
                key=page_widget_key,
            )
        with previous_column:
            st.button(
                "Previous",
                key="previous_page_{}".format(table_key),
                disabled=page_index == 0,
                on_click=set_page,
                args=(page_index - 1,),
            )
        with next_column:
            st.button(
                "Next",
                key="next_page_{}".format(table_key),
                disabled=page_index >= page_count - 1,
                on_click=set_page,
                args=(page_index + 1,),
            )

        st.caption(
            "Showing results {}–{} of {}.".format(
                page_start + 1,
                page_end,
                len(filtered_results),
            )
        )

    selected_result_indices = sorted(
        st.session_state.get(selected_indices_key, [])
    )
    if selected_result_indices:
        selected_results = [
            combined_results[index] for index in selected_result_indices
        ]
        with st.container(border=True):
            st.markdown('<div style="height: 1.25rem"></div>', unsafe_allow_html=True)
            st.markdown("#### Downloads for specific results")
            downloads_page_size = 3
            downloads_page_count = (
                len(selected_result_indices) + downloads_page_size - 1
            ) // downloads_page_size
            downloads_page_key = "{}_selected_downloads_page".format(table_key)
            st.session_state[downloads_page_key] = min(
                st.session_state.get(downloads_page_key, 0),
                downloads_page_count - 1,
            )
            downloads_page_index = st.session_state[downloads_page_key]

            downloads_start = downloads_page_index * downloads_page_size
            downloads_end = min(
                downloads_start + downloads_page_size,
                len(selected_result_indices),
            )
            for selected_result_index in selected_result_indices[
                downloads_start:downloads_end
            ]:
                result = combined_results[selected_result_index]
                st.markdown(
                    '<div style="font-size: 0.95rem; font-weight: 600; '
                    'margin: 1.25rem 0 0.75rem">{}</div>'.format(
                        html.escape(
                            "{} ({})".format(
                                format_multivalue(result["company_name"])
                                or "Selected result",
                                result["publication_date"],
                            )
                        )
                    ),
                    unsafe_allow_html=True,
                )
                download_entry_documents(
                    selected_document_entry(result, "EN"),
                    selected_document_entry(result, "TC"),
                    "{}_{}".format(table_key, selected_result_index + 1),
                )
                st.markdown('<div style="height: 1rem"></div>', unsafe_allow_html=True)

            st.caption(
                "Showing selected results {}–{} of {}.".format(
                    downloads_start + 1,
                    downloads_end,
                    len(selected_result_indices),
                )
            )
            if downloads_page_count > 1:
                (
                    downloads_page_column,
                    downloads_previous_column,
                    downloads_next_column,
                ) = st.columns([4, 1, 1], vertical_alignment="bottom")
                with downloads_page_column:
                    downloads_page_index = st.selectbox(
                        "Selected results page",
                        options=list(range(downloads_page_count)),
                        format_func=lambda page: "Page {} of {}".format(
                            page + 1,
                            downloads_page_count,
                        ),
                        key=downloads_page_key,
                    )
                with downloads_previous_column:
                    st.button(
                        "Previous",
                        key="previous_downloads_page_{}".format(table_key),
                        disabled=downloads_page_index == 0,
                        on_click=set_selected_download_page,
                        args=(downloads_page_key, downloads_page_index - 1),
                    )
                with downloads_next_column:
                    st.button(
                        "Next",
                        key="next_downloads_page_{}".format(table_key),
                        disabled=downloads_page_index >= downloads_page_count - 1,
                        on_click=set_selected_download_page,
                        args=(downloads_page_key, downloads_page_index + 1),
                    )

        if len(selected_results) > 1:
            with st.container(border=True):
                st.markdown('<div style="height: 1.25rem"></div>', unsafe_allow_html=True)
                st.markdown("#### Downloads for All Selected Results")
                english_available = any(
                    is_downloadable_pdf(selected_document_entry(result, "EN"))
                    for result in selected_results
                )
                chinese_available = any(
                    is_downloadable_pdf(selected_document_entry(result, "TC"))
                    for result in selected_results
                )
                any_document_available = english_available or chinese_available
                english_download, chinese_download, all_download = st.columns(3)
                with english_download:
                    download_button_with_custom_filename(
                        "Download All Selected English Results",
                        lambda: create_selected_archive(selected_results, "EN"),
                        selected_archive_filename("EN", search_stock_code),
                        "application/zip",
                        "selected_english_zip_{}".format(table_key),
                        disabled=not english_available,
                    )
                with chinese_download:
                    download_button_with_custom_filename(
                        "Download All Selected Chinese Results",
                        lambda: create_selected_archive(selected_results, "TC"),
                        selected_archive_filename("TC", search_stock_code),
                        "application/zip",
                        "selected_chinese_zip_{}".format(table_key),
                        disabled=not chinese_available,
                    )
                with all_download:
                    download_button_with_custom_filename(
                        "Download All Selected Results",
                        lambda: create_selected_archive(selected_results, None),
                        selected_archive_filename(None, search_stock_code),
                        "application/zip",
                        "selected_all_zip_{}".format(table_key),
                        disabled=not any_document_available,
                    )

    st.markdown('<div style="height: 1.5rem"></div>', unsafe_allow_html=True)
    with st.container(border=True):
        st.markdown("#### Download Table as Excel and JSON")
        excel_column, json_column = st.columns(2)
        with excel_column:
            download_button_with_custom_filename(
                "Download Excel",
                lambda: create_excel_file(combined_results),
                "{}.xlsx".format(export_filename),
                (
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                "excel_download_{}".format(table_key),
            )
        with json_column:
            download_button_with_custom_filename(
                "Download combined JSON",
                st.session_state["combined_json"],
                "{}.json".format(export_filename),
                "application/json",
                "json_download_{}".format(table_key),
            )

if submitted:
    search_generation = st.session_state.get("search_generation", 0) + 1
    st.session_state["search_generation"] = search_generation
    st.session_state.pop("combined_results", None)
    st.session_state.pop("combined_json", None)
    st.session_state.pop("announcements_filename", None)
    st.session_state.pop("search_stock_code", None)
    with st.spinner("Searching HKEXnews and loading results..."):
        try:
            english_document, chinese_document = run_search(
                stock_code=stock_code_value,
                start_date=from_date,
                end_date=to_date,
                headline_category=(
                    None if not selected_path or selected_path == ["All categories"]
                    else str(selected_path[-1])
                ),
                category=category_from_path(selected_path, category_tree),
            )
        except StockNotFoundError as exc:
            st.warning("Search failed: {} Please check the stock code.".format(exc))
        except (HTTPError, URLError, OSError, ValueError, RuntimeError) as exc:
            st.error("Search failed: {}".format(exc))
        else:
            try:
                combined_results = combine_results(english_document, chinese_document)
            except RuntimeError as exc:
                st.error(
                    "Search completed, but results could not be combined: {}".format(
                        exc
                    )
                )
            else:
                try:
                    combined_json = json.dumps(
                        {
                            "english_search": english_document["search_criteria"],
                            "chinese_search": chinese_document["search_criteria"],
                            "results": combined_results,
                        },
                        ensure_ascii=False,
                        indent=2,
                    ) + "\n"
                except (TypeError, ValueError) as exc:
                    st.error("Could not prepare search results: {}".format(exc))
                else:
                    export_filename = announcements_filename(
                        english_document["search_criteria"]
                    )
                    st.session_state["combined_results"] = combined_results
                    st.session_state["combined_json"] = combined_json
                    st.session_state["announcements_filename"] = export_filename
                    search_stock_code = (
                        str(english_document["search_criteria"].get("stock_code") or "")
                    )
                    st.session_state["search_stock_code"] = search_stock_code
                    show_combined_results(
                        combined_results,
                        "combined_results_table_{}".format(search_generation),
                        export_filename,
                        search_stock_code,
                    )
elif "combined_results" in st.session_state:
    show_combined_results(
        st.session_state["combined_results"],
        "combined_results_table_{}".format(
        st.session_state.get("search_generation", 0)
        ),
        st.session_state["announcements_filename"],
        st.session_state["search_stock_code"],
    )
