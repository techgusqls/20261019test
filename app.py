"""학교 기술실 기자재 관리 앱

Google Sheets를 데이터 저장소로 사용합니다. GOOGLE_SHEET_URL은 이 파일 상단에서 설정하세요.
비밀 인증정보는 st.secrets에서만 읽습니다.
"""
from __future__ import annotations

import hashlib
import json
import math
import re
import secrets as py_secrets
import threading
import uuid
from collections import defaultdict
from datetime import date, datetime
from typing import Any
from zoneinfo import ZoneInfo

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials
from gspread.exceptions import APIError, SpreadsheetNotFound, WorksheetNotFound

# -----------------------------------------------------------------------------
# 설정: 사용할 Google Sheets 문서의 URL을 여기에 붙여 넣으세요.
# 예: GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/문서ID/edit"
# -----------------------------------------------------------------------------
GOOGLE_SHEET_URL = "여기에_구글_시트_URL_입력"

APP_TITLE = "기술실 기자재 관리"
KST = ZoneInfo("Asia/Seoul")
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SHEET_HEADERS: dict[str, list[str]] = {
    "기자재목록": [
        "관리번호", "기자재명", "분류", "세부 품목", "보유 수량", "대여 중 수량",
        "수리 중 수량", "분실 수량", "보관 위치", "상태", "구입일", "구입 가격",
        "제조사 및 모델명", "비고", "등록일", "활성화 여부", "최종 수정일",
    ],
    "대여반납기록": [
        "거래 ID", "관리번호", "기자재명", "대여 수량", "반납 완료 수량", "대여자",
        "학년 및 반", "대여일", "반납 예정일", "실제 반납일", "비고", "거래 유형",
        "원거래 ID", "처리일시", "반납 상태",
    ],
    "수리기록": [
        "수리 ID", "관리번호", "기자재명", "수리 접수일", "수리 내용", "수리 상태",
        "수리 완료일", "처리 메모", "수리 수량",
    ],
    "변경이력": [
        "이력 ID", "변경 시각", "작업 유형", "관리번호", "변경 내용", "작업자",
    ],
}

ITEM_STATES = ["사용 가능", "대여 중", "수리 필요", "분실", "폐기"]
CATEGORIES = [
    "공구", "전기·전자", "측정·계측", "목공", "재봉·섬유", "조리", "안전·보호",
    "재료·소모품", "기계·설비", "기타",
]
PROCESS_LOCK = threading.RLock()

st.set_page_config(page_title=APP_TITLE, page_icon="🧰", layout="wide")
st.markdown(
    """
    <style>
      .block-container { padding-top: 1.25rem; padding-bottom: 2.5rem; }
      [data-testid="stMetric"] { border: 1px solid rgba(128,128,128,.22); padding: 14px 16px; border-radius: 12px; }
      div[data-testid="stDataFrame"] { width: 100%; }
      h1, h2, h3 { letter-spacing: -0.025em; }
    </style>
    """,
    unsafe_allow_html=True,
)


def now_kst() -> datetime:
    return datetime.now(KST)


def today_kst() -> date:
    return now_kst().date()


def date_text(value: Any) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    try:
        return pd.to_datetime(value).date().isoformat()
    except Exception:
        return str(value).strip()


def parse_date(value: Any) -> date | None:
    if value is None or str(value).strip() == "":
        return None
    try:
        return pd.to_datetime(value).date()
    except Exception:
        return None


def as_int(value: Any, default: int = 0) -> int:
    if value is None or str(value).strip() == "":
        return default
    try:
        number = float(str(value).replace(",", "").strip())
        if not math.isfinite(number):
            return default
        return int(number)
    except (TypeError, ValueError):
        return default


def as_bool(value: Any, default: bool = True) -> bool:
    if isinstance(value, bool):
        return value
    text_value = str(value).strip().lower()
    if text_value in {"true", "1", "yes", "y", "활성", "활성화", "사용"}:
        return True
    if text_value in {"false", "0", "no", "n", "비활성", "비활성화", "중지"}:
        return False
    return default


def new_id(prefix: str) -> str:
    return f"{prefix}-{uuid.uuid4().hex[:12].upper()}"


def suggest_item_id(items: list[dict[str, Any]]) -> str:
    prefix = f"TR-{today_kst():%Y%m%d}-"
    used = {str(row.get("관리번호", "")).strip() for row in items}
    counter = 1
    while f"{prefix}{counter:03d}" in used:
        counter += 1
    return f"{prefix}{counter:03d}"


def get_secret_credentials() -> dict[str, Any]:
    try:
        section = st.secrets["gcp_service_account"]
        credentials_info = dict(section)
    except Exception as exc:
        raise RuntimeError(
            "Google Sheets 인증 정보가 없습니다. `.streamlit/secrets.toml` 또는 Streamlit Cloud의 Secrets에 "
            "`[gcp_service_account]` 설정을 등록해 주세요."
        ) from exc
    required = {"type", "project_id", "private_key", "client_email", "token_uri"}
    missing = sorted(required - set(credentials_info))
    if missing:
        raise RuntimeError("서비스 계정 설정에서 다음 항목을 찾을 수 없습니다: " + ", ".join(missing))
    return credentials_info


@st.cache_resource(show_spinner=False)
def connect_spreadsheet(sheet_url: str, credential_fingerprint: str, _credentials_info: dict[str, Any]):
    """연결 객체만 캐시합니다. 실제 시트 데이터는 매 실행 때 다시 읽습니다."""
    credentials = Credentials.from_service_account_info(_credentials_info, scopes=SCOPES)
    client = gspread.authorize(credentials)
    return client.open_by_url(sheet_url)


class SheetsStore:
    def __init__(self, spreadsheet):
        self.spreadsheet = spreadsheet
        self.worksheets: dict[str, Any] = {}

    def ensure_sheets(self) -> None:
        existing = {ws.title: ws for ws in self.spreadsheet.worksheets()}
        for title, required_headers in SHEET_HEADERS.items():
            if title in existing:
                ws = existing[title]
            else:
                # 새 탭만 생성합니다. 다른 기존 탭이나 데이터는 삭제하지 않습니다.
                ws = self.spreadsheet.add_worksheet(
                    title=title, rows=max(1000, len(required_headers) + 10),
                    cols=max(30, len(required_headers) + 5),
                )
            values = ws.get_all_values()
            if not values or not any(str(cell).strip() for cell in values[0]):
                ws.update(values=[required_headers], range_name="A1")
            else:
                current_headers = [str(cell).strip() for cell in values[0]]
                missing_headers = [h for h in required_headers if h not in current_headers]
                if missing_headers:
                    # 기존 컬럼과 데이터를 건드리지 않고 누락된 컬럼만 오른쪽에 추가합니다.
                    current_headers.extend(missing_headers)
                    ws.update(values=[current_headers], range_name="A1")
            self.worksheets[title] = ws

    def read_sheet(self, title: str) -> list[dict[str, Any]]:
        ws = self.worksheets[title]
        values = ws.get_all_values()
        if not values:
            return []
        headers = [str(h).strip() for h in values[0]]
        records: list[dict[str, Any]] = []
        for row_num, row in enumerate(values[1:], start=2):
            if not any(str(cell).strip() for cell in row):
                continue
            record: dict[str, Any] = {
                header: (row[index] if index < len(row) else "")
                for index, header in enumerate(headers)
                if header
            }
            record["_row"] = row_num
            records.append(record)
        return records

    def load_all(self) -> dict[str, list[dict[str, Any]]]:
        # 이 함수는 각 시트의 최신값을 직접 조회합니다. 데이터 캐시는 쓰지 않습니다.
        return {title: self.read_sheet(title) for title in SHEET_HEADERS}

    @staticmethod
    def _cell(value: Any) -> dict[str, Any]:
        if value is None:
            return {"userEnteredValue": {"stringValue": ""}}
        if isinstance(value, bool):
            return {"userEnteredValue": {"boolValue": value}}
        if isinstance(value, int) and not isinstance(value, bool):
            return {"userEnteredValue": {"numberValue": value}}
        if isinstance(value, float) and math.isfinite(value):
            return {"userEnteredValue": {"numberValue": value}}
        return {"userEnteredValue": {"stringValue": str(value)}}

    def atomic_write(self, writes: list[tuple[str, dict[str, Any]]]) -> None:
        """여러 시트 행 변경을 하나의 spreadsheets.batchUpdate API 요청으로 적용합니다."""
        if not writes:
            return
        rows_cache: dict[str, list[list[str]]] = {}
        next_append_row: dict[str, int] = {}
        requests: list[dict[str, Any]] = []
        for title, record in writes:
            ws = self.worksheets[title]
            if title not in rows_cache:
                rows_cache[title] = ws.get_all_values()
                next_append_row[title] = len(rows_cache[title]) + 1
            sheet_values = rows_cache[title]
            headers = [str(h).strip() for h in (sheet_values[0] if sheet_values else SHEET_HEADERS[title])]
            existing_row = record.get("_row")
            if existing_row is not None:
                row_num = int(existing_row)
            else:
                row_num = next_append_row[title]
                next_append_row[title] += 1
            row_cells = []
            for header in headers:
                if not header:
                    row_cells.append(self._cell(""))
                else:
                    row_cells.append(self._cell(record.get(header, "")))
            requests.append(
                {
                    "updateCells": {
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": row_num - 1,
                            "endRowIndex": row_num,
                            "startColumnIndex": 0,
                            "endColumnIndex": len(headers),
                        },
                        "rows": [{"values": row_cells}],
                        "fields": "userEnteredValue",
                    }
                }
            )
        # Google Sheets API는 단일 batchUpdate 안의 요청들을 원자적으로 적용합니다.
        # 단, 읽기-검증-쓰기 전체를 다른 앱 인스턴스와 직렬화하는 트랜잭션은 아닙니다.
        self.spreadsheet.batch_update({"requests": requests})


@st.cache_resource(show_spinner=False)
def cached_store(sheet_url: str, credential_fingerprint: str, _credentials_info: dict[str, Any]) -> SheetsStore:
    """Sheets 연결·Worksheet 객체와 초기 헤더 검증 결과를 캐시합니다. 행 데이터는 캐시하지 않습니다."""
    spreadsheet = connect_spreadsheet(sheet_url, credential_fingerprint, _credentials_info)
    store = SheetsStore(spreadsheet)
    store.ensure_sheets()
    return store


def get_store() -> SheetsStore:
    if not GOOGLE_SHEET_URL.startswith("https://docs.google.com/spreadsheets/"):
        raise RuntimeError("`app.py` 상단의 `GOOGLE_SHEET_URL`에 실제 Google Sheets URL을 입력해 주세요.")
    credentials_info = get_secret_credentials()
    fingerprint = hashlib.sha256(
        json.dumps(credentials_info, sort_keys=True, ensure_ascii=False).encode("utf-8")
    ).hexdigest()
    return cached_store(GOOGLE_SHEET_URL, fingerprint, credentials_info)


def audit_record(task: str, item_id: str, details: str, operator: str) -> dict[str, Any]:
    return {
        "이력 ID": new_id("LOG"),
        "변경 시각": now_kst().isoformat(timespec="seconds"),
        "작업 유형": task,
        "관리번호": item_id,
        "변경 내용": details,
        "작업자": operator.strip() or "미입력",
    }


def loan_events(loan_records: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], dict[str, int]]:
    """대여 이벤트와 부분 반납 이벤트에서 거래별 반납 수량을 계산합니다.

    거래 유형 컬럼이 없는 구버전 행도 최대한 읽도록 보완합니다.
    """
    loans: list[dict[str, Any]] = []
    return_totals: dict[str, int] = defaultdict(int)
    for row in loan_records:
        transaction_id = str(row.get("거래 ID", "")).strip()
        event_type = str(row.get("거래 유형", "")).strip()
        borrowed = as_int(row.get("대여 수량"))
        returned = as_int(row.get("반납 완료 수량"))
        if event_type == "반납" or (not event_type and borrowed == 0 and returned > 0):
            original_id = str(row.get("원거래 ID", "")).strip()
            if original_id:
                return_totals[original_id] += returned
        elif event_type == "대여" or borrowed > 0:
            loan = dict(row)
            loan["_transaction_id"] = transaction_id
            loan["_borrowed"] = borrowed
            # 구버전에서는 반납 누계가 대여 행에 같이 기록될 수 있습니다.
            loan["_legacy_returned"] = returned if event_type != "대여" else 0
            loans.append(loan)
    return loans, return_totals


def get_loan_status(loan_records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    loans, return_totals = loan_events(loan_records)
    result = []
    for loan in loans:
        transaction_id = loan.get("_transaction_id", "")
        borrowed = as_int(loan.get("_borrowed"))
        returned = as_int(loan.get("_legacy_returned")) + return_totals.get(transaction_id, 0)
        outstanding = max(0, borrowed - returned)
        loan["_returned"] = min(borrowed, returned)
        loan["_outstanding"] = outstanding
        loan["_is_overdue"] = bool(
            outstanding > 0
            and parse_date(loan.get("반납 예정일")) is not None
            and parse_date(loan.get("반납 예정일")) < today_kst()
        )
        result.append(loan)
    return result


def active_repair_quantities(repairs: list[dict[str, Any]]) -> dict[str, int]:
    totals: dict[str, int] = defaultdict(int)
    for repair in repairs:
        status = str(repair.get("수리 상태", "접수")).strip()
        if status not in {"완료", "수리 완료", "취소"}:
            totals[str(repair.get("관리번호", "")).strip()] += max(1, as_int(repair.get("수리 수량"), 1))
    return totals


def loaned_quantities(loan_records: list[dict[str, Any]]) -> dict[str, int]:
    totals: dict[str, int] = defaultdict(int)
    for loan in get_loan_status(loan_records):
        totals[str(loan.get("관리번호", "")).strip()] += loan["_outstanding"]
    return totals


def availability(item: dict[str, Any], loans: list[dict[str, Any]], repairs: list[dict[str, Any]]) -> int:
    item_id = str(item.get("관리번호", "")).strip()
    total = max(0, as_int(item.get("보유 수량")))
    if not as_bool(item.get("활성화 여부", True)) or str(item.get("상태", "")).strip() == "폐기":
        return 0
    borrowed = loaned_quantities(loans).get(item_id, 0)
    repair_count = active_repair_quantities(repairs).get(item_id, 0)
    lost = max(0, as_int(item.get("분실 수량")))
    return max(0, total - borrowed - repair_count - lost)


def enrich_items(
    items: list[dict[str, Any]], loans: list[dict[str, Any]], repairs: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    borrowed_counts = loaned_quantities(loans)
    repair_counts = active_repair_quantities(repairs)
    result = []
    for original in items:
        item = dict(original)
        item_id = str(item.get("관리번호", "")).strip()
        total = max(0, as_int(item.get("보유 수량")))
        borrowed = borrowed_counts.get(item_id, 0)
        repair_count = repair_counts.get(item_id, 0)
        lost = max(0, as_int(item.get("분실 수량")))
        is_active = as_bool(item.get("활성화 여부", True))
        raw_state = str(item.get("상태", "사용 가능")).strip()
        available = availability(item, loans, repairs)
        if not is_active:
            state = "비활성"
        elif raw_state == "폐기":
            state = "폐기"
        elif available > 0:
            state = "사용 가능"
        elif borrowed > 0:
            state = "대여 중"
        elif repair_count > 0:
            state = "수리 필요"
        elif lost > 0:
            state = "분실"
        else:
            state = "사용 가능"
        item["보유 수량"] = total
        item["대여 중 수량"] = borrowed
        item["수리 중 수량"] = repair_count
        item["분실 수량"] = lost
        item["사용 가능 수량"] = available
        item["표시 상태"] = state
        item["활성화 여부"] = "활성" if is_active else "비활성"
        item["구입 가격"] = as_int(item.get("구입 가격"))
        result.append(item)
    return result


def find_item(items: list[dict[str, Any]], item_id: str) -> dict[str, Any] | None:
    wanted = str(item_id).strip()
    for item in items:
        if str(item.get("관리번호", "")).strip() == wanted:
            return item
    return None


def fmt_item(item: dict[str, Any]) -> str:
    return f"{item.get('관리번호', '')} · {item.get('기자재명', '')} (대여 가능 {as_int(item.get('사용 가능 수량'))}개)"


def fmt_loan(loan: dict[str, Any]) -> str:
    due = loan.get("반납 예정일", "미정") or "미정"
    return (
        f"{loan.get('_transaction_id', '')} · {loan.get('기자재명', '')} · "
        f"{loan.get('대여자', '')} · 미반납 {loan.get('_outstanding', 0)}개 · 예정 {due}"
    )


def to_df(records: list[dict[str, Any]], columns: list[str] | None = None) -> pd.DataFrame:
    cleaned = [{k: v for k, v in row.items() if not str(k).startswith("_") or (columns is not None and k in columns and k != "_row")} for row in records]
    df = pd.DataFrame(cleaned)
    if columns is not None:
        for col in columns:
            if col not in df.columns:
                df[col] = ""
        df = df[columns]
    return df


def safe_csv(df: pd.DataFrame, filename: str, label: str) -> None:
    payload = df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(label, data=payload, file_name=filename, mime="text/csv", use_container_width=True)


def operator_name() -> str:
    return str(st.session_state.get("operator_name", "")).strip() or "미입력"


def ensure_capacity(item: dict[str, Any], loans: list[dict[str, Any]], repairs: list[dict[str, Any]], new_total: int) -> None:
    item_id = str(item.get("관리번호", "")).strip()
    borrowed = loaned_quantities(loans).get(item_id, 0)
    repair_count = active_repair_quantities(repairs).get(item_id, 0)
    lost = max(0, as_int(item.get("분실 수량")))
    if new_total < borrowed + repair_count + lost:
        raise ValueError(
            f"보유 수량은 대여 중({borrowed}) + 수리 중({repair_count}) + 분실({lost})인 "
            f"수량의 합계({borrowed + repair_count + lost})보다 작을 수 없습니다."
        )


def current_records(store: SheetsStore) -> dict[str, list[dict[str, Any]]]:
    # 대여/수리 처리 직전 재고 확인용 최신 스냅샷.
    return store.load_all()


def handle_loan(
    store: SheetsStore, item_id: str, quantity: int, borrower: str, grade_class: str,
    loan_date: date, due_date: date, note: str, operator: str, transaction_id: str,
) -> None:
    with PROCESS_LOCK:
        data = current_records(store)
        if any(str(r.get("거래 ID", "")).strip() == transaction_id for r in data["대여반납기록"]):
            raise ValueError("이미 처리된 대여 요청입니다. 화면을 새로고침한 뒤 기록을 확인해 주세요.")
        items = data["기자재목록"]
        item = find_item(items, item_id)
        if item is None or not as_bool(item.get("활성화 여부", True)):
            raise ValueError("선택한 기자재를 찾을 수 없거나 비활성 상태입니다.")
        available = availability(item, data["대여반납기록"], data["수리기록"])
        if quantity <= 0:
            raise ValueError("대여 수량은 1개 이상이어야 합니다.")
        if quantity > available:
            raise ValueError(f"최신 대여 가능 수량은 {available}개입니다. 화면을 확인한 뒤 다시 시도해 주세요.")
        if due_date < loan_date:
            raise ValueError("반납 예정일은 대여일보다 빠를 수 없습니다.")
        now_text = now_kst().isoformat(timespec="seconds")
        borrowed_after = loaned_quantities(data["대여반납기록"]).get(item_id, 0) + quantity
        item["대여 중 수량"] = borrowed_after
        item["수리 중 수량"] = active_repair_quantities(data["수리기록"]).get(item_id, 0)
        item["최종 수정일"] = now_text
        loan_record = {
            "거래 ID": transaction_id, "관리번호": item_id,
            "기자재명": str(item.get("기자재명", "")), "대여 수량": quantity,
            "반납 완료 수량": 0, "대여자": borrower.strip(), "학년 및 반": grade_class.strip(),
            "대여일": loan_date.isoformat(), "반납 예정일": due_date.isoformat(),
            "실제 반납일": "", "비고": note.strip(), "거래 유형": "대여",
            "원거래 ID": "", "처리일시": now_text, "반납 상태": "미반납",
        }
        store.atomic_write([
            ("기자재목록", item),
            ("대여반납기록", loan_record),
            ("변경이력", audit_record("대여", item_id, f"{borrower.strip()} / {quantity}개 대여 / 거래 ID {transaction_id}", operator)),
        ])


def handle_return(
    store: SheetsStore, transaction_id: str, quantity: int, return_date: date,
    condition: str, note: str, operator: str, return_event_id: str,
) -> None:
    with PROCESS_LOCK:
        data = current_records(store)
        if any(str(r.get("거래 ID", "")).strip() == return_event_id for r in data["대여반납기록"]):
            raise ValueError("이미 처리된 반납 요청입니다. 화면을 새로고침한 뒤 기록을 확인해 주세요.")
        statuses = get_loan_status(data["대여반납기록"])
        loan = next((r for r in statuses if r.get("_transaction_id") == transaction_id), None)
        if loan is None or loan.get("_outstanding", 0) <= 0:
            raise ValueError("반납할 대여 건이 없거나 이미 반납 완료된 기록입니다.")
        outstanding = as_int(loan.get("_outstanding"))
        if quantity <= 0 or quantity > outstanding:
            raise ValueError(f"반납 수량은 1~{outstanding}개로 입력해 주세요.")
        item_id = str(loan.get("관리번호", "")).strip()
        item = find_item(data["기자재목록"], item_id)
        if item is None:
            raise ValueError("대여 기록에 연결된 기자재가 기자재목록에서 발견되지 않습니다. 데이터를 점검해 주세요.")
        now_text = now_kst().isoformat(timespec="seconds")
        new_outstanding = outstanding - quantity
        returned_total = as_int(loan.get("_returned")) + quantity
        item["대여 중 수량"] = max(0, loaned_quantities(data["대여반납기록"]).get(item_id, 0) - quantity)
        item["수리 중 수량"] = active_repair_quantities(data["수리기록"]).get(item_id, 0)
        item["최종 수정일"] = now_text
        return_record = {
            "거래 ID": return_event_id, "관리번호": item_id,
            "기자재명": str(loan.get("기자재명", "")), "대여 수량": 0,
            "반납 완료 수량": quantity, "대여자": str(loan.get("대여자", "")),
            "학년 및 반": str(loan.get("학년 및 반", "")), "대여일": str(loan.get("대여일", "")),
            "반납 예정일": str(loan.get("반납 예정일", "")), "실제 반납일": return_date.isoformat(),
            "비고": note.strip(), "거래 유형": "반납", "원거래 ID": transaction_id,
            "처리일시": now_text, "반납 상태": condition,
        }
        status_text = "전체 반납" if new_outstanding == 0 else f"부분 반납 (미반납 {new_outstanding}개)"
        store.atomic_write([
            ("기자재목록", item),
            ("대여반납기록", return_record),
            ("변경이력", audit_record("반납", item_id, f"거래 ID {transaction_id}: {quantity}개 반납, {status_text}, 상태 {condition}", operator)),
        ])


def handle_repair_start(
    store: SheetsStore, item_id: str, quantity: int, repair_date: date, description: str,
    note: str, operator: str, repair_id: str,
) -> None:
    with PROCESS_LOCK:
        data = current_records(store)
        if any(str(r.get("수리 ID", "")).strip() == repair_id for r in data["수리기록"]):
            raise ValueError("이미 처리된 수리 접수 요청입니다. 기록을 새로고침해 주세요.")
        item = find_item(data["기자재목록"], item_id)
        if item is None or not as_bool(item.get("활성화 여부", True)):
            raise ValueError("선택한 기자재를 찾을 수 없거나 비활성 상태입니다.")
        available = availability(item, data["대여반납기록"], data["수리기록"])
        if quantity <= 0 or quantity > available:
            raise ValueError(f"수리 접수 수량은 1~{available}개로 입력해 주세요.")
        now_text = now_kst().isoformat(timespec="seconds")
        active_after = active_repair_quantities(data["수리기록"]).get(item_id, 0) + quantity
        item["수리 중 수량"] = active_after
        item["대여 중 수량"] = loaned_quantities(data["대여반납기록"]).get(item_id, 0)
        item["최종 수정일"] = now_text
        repair = {
            "수리 ID": repair_id, "관리번호": item_id, "기자재명": str(item.get("기자재명", "")),
            "수리 접수일": repair_date.isoformat(), "수리 내용": description.strip(),
            "수리 상태": "접수", "수리 완료일": "", "처리 메모": note.strip(), "수리 수량": quantity,
        }
        store.atomic_write([
            ("기자재목록", item), ("수리기록", repair),
            ("변경이력", audit_record("수리 접수", item_id, f"{quantity}개 / {description.strip()} / 수리 ID {repair_id}", operator)),
        ])


def handle_repair_finish(
    store: SheetsStore, repair_id: str, finish_date: date, note: str, operator: str,
) -> None:
    with PROCESS_LOCK:
        data = current_records(store)
        repair = next((r for r in data["수리기록"] if str(r.get("수리 ID", "")).strip() == repair_id), None)
        if repair is None:
            raise ValueError("선택한 수리 기록을 찾을 수 없습니다.")
        if str(repair.get("수리 상태", "")).strip() in {"완료", "수리 완료", "취소"}:
            raise ValueError("이미 완료 또는 취소된 수리 기록입니다.")
        item_id = str(repair.get("관리번호", "")).strip()
        item = find_item(data["기자재목록"], item_id)
        if item is None:
            raise ValueError("수리 기록에 연결된 기자재를 찾을 수 없습니다.")
        repair["수리 상태"] = "완료"
        repair["수리 완료일"] = finish_date.isoformat()
        old_note = str(repair.get("처리 메모", "")).strip()
        repair["처리 메모"] = " / ".join(part for part in (old_note, note.strip()) if part)
        active_after = active_repair_quantities(
            [r for r in data["수리기록"] if str(r.get("수리 ID", "")).strip() != repair_id]
        ).get(item_id, 0)
        item["수리 중 수량"] = active_after
        item["대여 중 수량"] = loaned_quantities(data["대여반납기록"]).get(item_id, 0)
        item["최종 수정일"] = now_kst().isoformat(timespec="seconds")
        store.atomic_write([
            ("수리기록", repair), ("기자재목록", item),
            ("변경이력", audit_record("수리 완료", item_id, f"수리 ID {repair_id}, 수량 {as_int(repair.get('수리 수량'), 1)}개 / {note.strip()}", operator)),
        ])


def set_lost_quantity(
    store: SheetsStore, item_id: str, change: int, operator: str, reason: str,
) -> None:
    with PROCESS_LOCK:
        data = current_records(store)
        item = find_item(data["기자재목록"], item_id)
        if item is None or not as_bool(item.get("활성화 여부", True)):
            raise ValueError("선택한 기자재를 찾을 수 없거나 비활성 상태입니다.")
        old_lost = max(0, as_int(item.get("분실 수량")))
        if change == 0:
            raise ValueError("변경 수량을 확인해 주세요.")
        if change < 0 and old_lost < abs(change):
            raise ValueError(f"현재 분실 수량은 {old_lost}개이므로 그보다 많이 복구할 수 없습니다.")
        if change > 0:
            available = availability(item, data["대여반납기록"], data["수리기록"])
            if change > available:
                raise ValueError(f"분실 처리 가능한 수량은 현재 {available}개입니다.")
        item["분실 수량"] = old_lost + change
        item["대여 중 수량"] = loaned_quantities(data["대여반납기록"]).get(item_id, 0)
        item["수리 중 수량"] = active_repair_quantities(data["수리기록"]).get(item_id, 0)
        item["최종 수정일"] = now_kst().isoformat(timespec="seconds")
        kind = "분실 처리" if change > 0 else "분실 복구"
        description = f"{abs(change)}개 / {reason.strip()} / 분실 수량 {old_lost} → {old_lost + change}개"
        store.atomic_write([
            ("기자재목록", item), ("변경이력", audit_record(kind, item_id, description, operator)),
        ])


def filter_items(
    items: list[dict[str, Any]], query: str, category: str, location: str, state: str, include_inactive: bool
) -> list[dict[str, Any]]:
    result = list(items)
    if not include_inactive:
        result = [r for r in result if r.get("활성화 여부") == "활성"]
    query = query.strip().lower()
    if query:
        result = [
            r for r in result if query in str(r.get("기자재명", "")).lower()
            or query in str(r.get("관리번호", "")).lower()
        ]
    if category != "전체":
        result = [r for r in result if str(r.get("분류", "")) == category]
    if location != "전체":
        result = [r for r in result if str(r.get("보관 위치", "")) == location]
    if state != "전체":
        result = [r for r in result if str(r.get("표시 상태", "")) == state]
    return result


def render_dashboard(data: dict[str, list[dict[str, Any]]]) -> None:
    st.title("기술실 기자재 현황")
    items = enrich_items(data["기자재목록"], data["대여반납기록"], data["수리기록"])
    active_items = [r for r in items if r.get("활성화 여부") == "활성"]
    borrowed_total = sum(as_int(r.get("대여 중 수량")) for r in active_items)
    repair_total = sum(as_int(r.get("수리 중 수량")) for r in active_items)
    lost_total = sum(as_int(r.get("분실 수량")) for r in active_items)
    available_total = sum(as_int(r.get("사용 가능 수량")) for r in active_items)
    total_stock = sum(as_int(r.get("보유 수량")) for r in active_items)
    metrics = st.columns(5)
    metric_values = [
        ("전체 기자재 수량", total_stock), ("사용 가능 수량", available_total),
        ("대여 중 수량", borrowed_total), ("수리 필요 수량", repair_total), ("분실 수량", lost_total),
    ]
    for col, (label, value) in zip(metrics, metric_values):
        col.metric(label, f"{value:,}개")
    left, right = st.columns(2)
    with left:
        st.subheader("분류별 보유 현황")
        cat_df = pd.DataFrame(active_items)
        if not cat_df.empty:
            by_category = cat_df.groupby("분류", dropna=False)["보유 수량"].sum().sort_values(ascending=False)
            st.bar_chart(by_category)
            st.dataframe(by_category.rename("보유 수량").reset_index(), hide_index=True, use_container_width=True)
        else:
            st.info("등록된 기자재가 없습니다.")
    with right:
        st.subheader("최근 등록된 기자재")
        recent = sorted(active_items, key=lambda r: str(r.get("등록일", "")), reverse=True)[:8]
        if recent:
            st.dataframe(to_df(recent, ["등록일", "관리번호", "기자재명", "분류", "보유 수량", "보관 위치"]), hide_index=True, use_container_width=True)
        else:
            st.info("최근 등록 내역이 없습니다.")
    st.subheader("최근 대여 및 반납 내역")
    recent_events = sorted(data["대여반납기록"], key=lambda r: str(r.get("처리일시", r.get("대여일", ""))), reverse=True)[:10]
    if recent_events:
        st.dataframe(to_df(recent_events, ["처리일시", "거래 유형", "관리번호", "기자재명", "대여 수량", "반납 완료 수량", "대여자", "원거래 ID", "반납 상태"]), hide_index=True, use_container_width=True)
    else:
        st.info("대여 및 반납 내역이 없습니다.")
    overdue = [r for r in get_loan_status(data["대여반납기록"]) if r.get("_is_overdue")]
    st.subheader("연체 현황")
    if overdue:
        st.warning(f"반납 예정일이 지난 대여 건이 {len(overdue)}건 있습니다.")
        st.dataframe(to_df(overdue, ["_transaction_id", "관리번호", "기자재명", "대여자", "학년 및 반", "반납 예정일", "_outstanding"]), hide_index=True, use_container_width=True)
    else:
        st.success("현재 연체된 대여 건이 없습니다.")


def render_inventory(data: dict[str, list[dict[str, Any]]], store: SheetsStore, operator: str) -> None:
    st.title("기자재 목록 조회 및 관리")
    items = enrich_items(data["기자재목록"], data["대여반납기록"], data["수리기록"])
    categories = sorted({str(r.get("분류", "")).strip() for r in items if str(r.get("분류", "")).strip()})
    locations = sorted({str(r.get("보관 위치", "")).strip() for r in items if str(r.get("보관 위치", "")).strip()})
    c1, c2, c3 = st.columns([2, 1, 1])
    with c1:
        query = st.text_input("기자재명 또는 관리번호 검색", placeholder="예: 드릴, TR-20261019-001")
    with c2:
        category = st.selectbox("분류", ["전체"] + categories)
    with c3:
        state = st.selectbox("상태", ["전체", "사용 가능", "대여 중", "수리 필요", "분실", "폐기", "비활성"])
    c4, c5, c6 = st.columns([1, 1, 1])
    with c4:
        location = st.selectbox("보관 위치", ["전체"] + locations)
    with c5:
        sort_by = st.selectbox("정렬", ["재고 수량순", "기자재명순", "관리번호순", "등록일순"])
    with c6:
        include_inactive = st.checkbox("비활성 기자재 포함", value=False)
    filtered = filter_items(items, query, category, location, state, include_inactive)
    if sort_by == "재고 수량순":
        filtered.sort(key=lambda r: as_int(r.get("사용 가능 수량")), reverse=True)
    elif sort_by == "기자재명순":
        filtered.sort(key=lambda r: str(r.get("기자재명", "")).lower())
    elif sort_by == "관리번호순":
        filtered.sort(key=lambda r: str(r.get("관리번호", "")))
    else:
        filtered.sort(key=lambda r: str(r.get("등록일", "")), reverse=True)
    st.caption(f"검색 결과 {len(filtered)}건")
    show_cols = ["관리번호", "기자재명", "분류", "세부 품목", "보유 수량", "대여 중 수량", "사용 가능 수량", "수리 중 수량", "분실 수량", "보관 위치", "표시 상태"]
    if filtered:
        st.dataframe(to_df(filtered, show_cols), hide_index=True, use_container_width=True)
    else:
        st.info("검색 조건에 맞는 기자재가 없습니다.")

    if not items:
        return
    st.divider()
    st.subheader("기자재 상세 정보 및 수정")
    active_ids = [str(r.get("관리번호", "")) for r in items]
    selected_id = st.selectbox("상세 정보를 확인할 기자재", active_ids, format_func=lambda x: f"{x} · {find_item(items, x).get('기자재명', '') if find_item(items, x) else ''}")
    selected = find_item(data["기자재목록"], selected_id)
    selected_enriched = find_item(items, selected_id)
    if selected is None or selected_enriched is None:
        return
    with st.expander("상세 정보 보기", expanded=True):
        detail_cols = st.columns(4)
        detail_cols[0].write(f"**상태:** {selected_enriched.get('표시 상태', '')}")
        detail_cols[1].write(f"**보유:** {as_int(selected.get('보유 수량'))}개")
        detail_cols[2].write(f"**대여 가능:** {selected_enriched.get('사용 가능 수량', 0)}개")
        detail_cols[3].write(f"**구입 가격:** {as_int(selected.get('구입 가격')):,}원")
        st.write(f"**제조사·모델:** {selected.get('제조사 및 모델명', '') or '-'}")
        st.write(f"**비고:** {selected.get('비고', '') or '-'}")
    with st.form("edit_inventory_form"):
        e1, e2, e3 = st.columns(3)
        with e1:
            edit_name = st.text_input("기자재명 *", value=str(selected.get("기자재명", "")))
            edit_category = st.text_input("분류 *", value=str(selected.get("분류", "")))
            edit_subtype = st.text_input("세부 품목", value=str(selected.get("세부 품목", "")))
            edit_total = st.number_input("보유 수량 *", min_value=0, step=1, value=max(0, as_int(selected.get("보유 수량"))))
        with e2:
            edit_location = st.text_input("보관 위치", value=str(selected.get("보관 위치", "")))
            raw_status = str(selected.get("상태", "사용 가능"))
            edit_status = st.selectbox("기본 상태", ["사용 가능", "폐기"], index=1 if raw_status == "폐기" else 0)
            old_purchase_date = parse_date(selected.get("구입일")) or today_kst()
            edit_purchase_date = st.date_input("구입일", value=old_purchase_date)
            edit_price = st.number_input("구입 가격(원)", min_value=0, step=1000, value=max(0, as_int(selected.get("구입 가격"))))
        with e3:
            edit_model = st.text_input("제조사 또는 모델명", value=str(selected.get("제조사 및 모델명", "")))
            edit_note = st.text_area("비고", value=str(selected.get("비고", "")), height=130)
        update_button = st.form_submit_button("수정 내용 저장", type="primary", use_container_width=True)
    if update_button:
        if not edit_name.strip() or not edit_category.strip():
            st.error("기자재명과 분류는 필수 입력 항목입니다.")
        else:
            try:
                fresh = current_records(store)
                current = find_item(fresh["기자재목록"], selected_id)
                if current is None:
                    raise ValueError("최신 데이터에서 기자재를 찾을 수 없습니다. 새로고침 후 다시 시도해 주세요.")
                ensure_capacity(current, fresh["대여반납기록"], fresh["수리기록"], int(edit_total))
                before = {k: current.get(k, "") for k in ["기자재명", "분류", "세부 품목", "보유 수량", "보관 위치", "상태", "구입일", "구입 가격", "제조사 및 모델명", "비고"]}
                current.update({
                    "기자재명": edit_name.strip(), "분류": edit_category.strip(),
                    "세부 품목": edit_subtype.strip(), "보유 수량": int(edit_total),
                    "보관 위치": edit_location.strip(), "상태": edit_status,
                    "구입일": edit_purchase_date.isoformat(), "구입 가격": int(edit_price),
                    "제조사 및 모델명": edit_model.strip(), "비고": edit_note.strip(),
                    "대여 중 수량": loaned_quantities(fresh["대여반납기록"]).get(selected_id, 0),
                    "수리 중 수량": active_repair_quantities(fresh["수리기록"]).get(selected_id, 0),
                    "최종 수정일": now_kst().isoformat(timespec="seconds"),
                })
                changed = [f"{k}: {before[k]} → {current.get(k, '')}" for k in before if str(before[k]) != str(current.get(k, ''))]
                store.atomic_write([
                    ("기자재목록", current),
                    ("변경이력", audit_record("기자재 수정", selected_id, "; ".join(changed) or "저장 버튼 실행(변경 사항 없음)", operator)),
                ])
                st.success("기자재 정보를 저장했습니다.")
                st.rerun()
            except Exception as exc:
                st.error(f"수정 실패: {exc}")
    with st.expander("기자재 비활성화", expanded=False):
        st.warning("대여·수리 이력은 보존됩니다. 비활성화한 기자재는 신규 대여와 일반 목록에서 제외됩니다.")
        confirm = st.checkbox(f"{selected_id} 기자재를 비활성화하는 데 동의합니다.", key=f"deactivate_confirm_{selected_id}")
        if st.button("기자재 비활성화", disabled=not confirm, key=f"deactivate_button_{selected_id}"):
            try:
                fresh = current_records(store)
                current = find_item(fresh["기자재목록"], selected_id)
                if current is None:
                    raise ValueError("기자재를 찾을 수 없습니다.")
                current["활성화 여부"] = False
                current["최종 수정일"] = now_kst().isoformat(timespec="seconds")
                store.atomic_write([
                    ("기자재목록", current),
                    ("변경이력", audit_record("기자재 비활성화", selected_id, "기자재를 비활성화함. 이력은 보존됨.", operator)),
                ])
                st.success("기자재를 비활성화했습니다.")
                st.rerun()
            except Exception as exc:
                st.error(f"비활성화 실패: {exc}")


def render_register(data: dict[str, list[dict[str, Any]]], store: SheetsStore, operator: str) -> None:
    st.title("기자재 등록")
    items = enrich_items(data["기자재목록"], data["대여반납기록"], data["수리기록"])
    suggested_id = suggest_item_id(data["기자재목록"])
    with st.form("register_item_form"):
        auto_id = st.checkbox("관리번호 자동 생성 사용", value=True)
        item_id_input = st.text_input("기자재 관리번호 *", value=suggested_id, help="자동 생성 해제 시 직접 입력할 수 있습니다.")
        a, b, c = st.columns(3)
        with a:
            item_name = st.text_input("기자재명 *", placeholder="예: 충전 드릴")
            category = st.selectbox("분류 *", CATEGORIES)
            subtype = st.text_input("세부 품목", placeholder="예: 18V 무선 드릴")
            quantity = st.number_input("보유 수량 *", min_value=1, value=1, step=1)
        with b:
            location = st.text_input("보관 위치", placeholder="예: 기술실 2번 공구장")
            status = st.selectbox("기본 상태", ["사용 가능", "폐기"])
            purchase_date = st.date_input("구입일", value=today_kst())
            price = st.number_input("구입 가격(원)", min_value=0, value=0, step=1000)
        with c:
            model = st.text_input("제조사 또는 모델명")
            note = st.text_area("비고", height=160)
        submit = st.form_submit_button("기자재 등록", type="primary", use_container_width=True)
    if submit:
        final_id = suggested_id if auto_id else item_id_input.strip()
        if not final_id or not item_name.strip():
            st.error("관리번호와 기자재명은 필수 입력 항목입니다.")
        elif not category:
            st.error("분류를 선택해 주세요.")
        elif int(quantity) < 1:
            st.error("보유 수량은 1개 이상이어야 합니다.")
        else:
            try:
                with PROCESS_LOCK:
                    fresh = current_records(store)
                    if any(str(r.get("관리번호", "")).strip().casefold() == final_id.casefold() for r in fresh["기자재목록"]):
                        raise ValueError("이미 사용 중인 관리번호입니다. 다른 번호를 입력해 주세요.")
                    now_text = now_kst().isoformat(timespec="seconds")
                    new_item = {
                        "관리번호": final_id, "기자재명": item_name.strip(), "분류": category,
                        "세부 품목": subtype.strip(), "보유 수량": int(quantity), "대여 중 수량": 0,
                        "수리 중 수량": 0, "분실 수량": 0, "보관 위치": location.strip(),
                        "상태": status, "구입일": purchase_date.isoformat(), "구입 가격": int(price),
                        "제조사 및 모델명": model.strip(), "비고": note.strip(), "등록일": today_kst().isoformat(),
                        "활성화 여부": True, "최종 수정일": now_text,
                    }
                    store.atomic_write([
                        ("기자재목록", new_item),
                        ("변경이력", audit_record("기자재 등록", final_id, f"{item_name.strip()} / {int(quantity)}개 등록", operator)),
                    ])
                st.success(f"기자재를 등록했습니다. 관리번호: {final_id}")
                st.rerun()
            except Exception as exc:
                st.error(f"등록 실패: {exc}")


def render_loans(data: dict[str, list[dict[str, Any]]], store: SheetsStore, operator: str) -> None:
    st.title("대여 및 반납")
    items = enrich_items(data["기자재목록"], data["대여반납기록"], data["수리기록"])
    active_items = [r for r in items if r.get("활성화 여부") == "활성" and as_int(r.get("사용 가능 수량")) > 0]
    current_loans = [r for r in get_loan_status(data["대여반납기록"]) if r.get("_outstanding", 0) > 0]
    borrow_tab, return_tab, status_tab = st.tabs(["기자재 대여", "기자재 반납", "대여 현황"])
    with borrow_tab:
        if not active_items:
            st.info("현재 대여 가능한 기자재가 없습니다. 기자재 재고 및 상태를 확인해 주세요.")
        else:
            with st.form("loan_form"):
                selected_item_id = st.selectbox("대여 기자재 *", [str(r.get("관리번호")) for r in active_items], format_func=lambda x: fmt_item(find_item(active_items, x) or {}))
                selected_item = find_item(active_items, selected_item_id) or {}
                max_qty = max(1, as_int(selected_item.get("사용 가능 수량")))
                quantity = st.number_input("대여 수량 *", min_value=1, max_value=max_qty, value=1, step=1)
                borrower = st.text_input("대여자 이름 *")
                grade_class = st.text_input("소속 학년 및 반", placeholder="예: 2학년 3반")
                c1, c2 = st.columns(2)
                with c1:
                    loan_date = st.date_input("대여일 *", value=today_kst())
                with c2:
                    due_date = st.date_input("반납 예정일 *", value=today_kst())
                note = st.text_area("비고", key="loan_note")
                loan_submit = st.form_submit_button("대여 처리", type="primary", use_container_width=True)
            if loan_submit:
                if not borrower.strip():
                    st.error("대여자 이름을 입력해 주세요.")
                else:
                    st.session_state.setdefault("loan_submission_id", new_id("LOAN"))
                    transaction_id = st.session_state["loan_submission_id"]
                    try:
                        handle_loan(store, selected_item_id, int(quantity), borrower, grade_class, loan_date, due_date, note, operator, transaction_id)
                        st.session_state["loan_submission_id"] = new_id("LOAN")
                        st.success(f"대여를 처리했습니다. 거래 ID: {transaction_id}")
                        st.rerun()
                    except Exception as exc:
                        st.error(f"대여 처리 실패: {exc}")
    with return_tab:
        if not current_loans:
            st.info("현재 미반납 대여 건이 없습니다.")
        else:
            with st.form("return_form"):
                transaction_ids = [str(r.get("_transaction_id", "")) for r in current_loans]
                selected_transaction_id = st.selectbox("반납할 대여 기록 *", transaction_ids, format_func=lambda x: fmt_loan(next((r for r in current_loans if r.get("_transaction_id") == x), {})))
                selected_loan = next((r for r in current_loans if r.get("_transaction_id") == selected_transaction_id), {})
                max_return = max(1, as_int(selected_loan.get("_outstanding")))
                return_qty = st.number_input("이번 반납 수량 *", min_value=1, max_value=max_return, value=max_return, step=1)
                return_date = st.date_input("실제 반납일 *", value=today_kst())
                condition = st.selectbox("반납 상태", ["정상", "파손 확인", "구성품 누락", "추가 점검 필요"])
                return_note = st.text_area("반납 비고", key="return_note")
                return_submit = st.form_submit_button("반납 처리", type="primary", use_container_width=True)
            if return_submit:
                st.session_state.setdefault("return_submission_id", new_id("RET"))
                return_event_id = st.session_state["return_submission_id"]
                try:
                    handle_return(store, selected_transaction_id, int(return_qty), return_date, condition, return_note, operator, return_event_id)
                    st.session_state["return_submission_id"] = new_id("RET")
                    st.success(f"반납을 처리했습니다. 기록 ID: {return_event_id}")
                    st.rerun()
                except Exception as exc:
                    st.error(f"반납 처리 실패: {exc}")
    with status_tab:
        if current_loans:
            status_rows = []
            for row in current_loans:
                status_rows.append({
                    "거래 ID": row.get("_transaction_id", ""), "관리번호": row.get("관리번호", ""),
                    "기자재명": row.get("기자재명", ""), "대여자": row.get("대여자", ""),
                    "학년 및 반": row.get("학년 및 반", ""), "대여 수량": row.get("_borrowed", 0),
                    "반납 완료 수량": row.get("_returned", 0), "미반납 수량": row.get("_outstanding", 0),
                    "대여일": row.get("대여일", ""), "반납 예정일": row.get("반납 예정일", ""),
                    "상태": "연체" if row.get("_is_overdue") else "대여 중",
                })
            st.dataframe(pd.DataFrame(status_rows), hide_index=True, use_container_width=True)
            overdue_count = sum(1 for r in current_loans if r.get("_is_overdue"))
            if overdue_count:
                st.warning(f"연체 {overdue_count}건을 확인했습니다.")
        else:
            st.success("현재 대여 중인 기자재가 없습니다.")


def render_repair_loss(data: dict[str, list[dict[str, Any]]], store: SheetsStore, operator: str) -> None:
    st.title("수리 및 분실 관리")
    items = enrich_items(data["기자재목록"], data["대여반납기록"], data["수리기록"])
    active_items = [r for r in items if r.get("활성화 여부") == "활성"]
    repair_tab, loss_tab, history_tab = st.tabs(["수리 접수 및 처리", "분실·복구", "수리 현황"])
    with repair_tab:
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("수리 접수")
            can_repair = [r for r in active_items if as_int(r.get("사용 가능 수량")) > 0]
            if not can_repair:
                st.info("수리 접수할 수 있는 대여 가능 재고가 없습니다.")
            else:
                with st.form("repair_start_form"):
                    repair_item_id = st.selectbox("기자재 *", [str(r.get("관리번호")) for r in can_repair], format_func=lambda x: fmt_item(find_item(can_repair, x) or {}))
                    selected_repair_item = find_item(can_repair, repair_item_id) or {}
                    repair_max = max(1, as_int(selected_repair_item.get("사용 가능 수량")))
                    repair_qty = st.number_input("수리 접수 수량 *", min_value=1, max_value=repair_max, value=1, step=1)
                    repair_date = st.date_input("수리 접수일", value=today_kst())
                    repair_description = st.text_area("수리 내용 *", placeholder="고장 증상, 필요한 조치 등")
                    repair_note = st.text_input("처리 메모")
                    repair_submit = st.form_submit_button("수리 접수", type="primary", use_container_width=True)
                if repair_submit:
                    if not repair_description.strip():
                        st.error("수리 내용을 입력해 주세요.")
                    else:
                        st.session_state.setdefault("repair_submission_id", new_id("REP"))
                        repair_id = st.session_state["repair_submission_id"]
                        try:
                            handle_repair_start(store, repair_item_id, int(repair_qty), repair_date, repair_description, repair_note, operator, repair_id)
                            st.session_state["repair_submission_id"] = new_id("REP")
                            st.success(f"수리를 접수했습니다. 수리 ID: {repair_id}")
                            st.rerun()
                        except Exception as exc:
                            st.error(f"수리 접수 실패: {exc}")
        with c2:
            st.subheader("수리 완료 처리")
            active_repairs = [r for r in data["수리기록"] if str(r.get("수리 상태", "접수")).strip() not in {"완료", "수리 완료", "취소"}]
            if not active_repairs:
                st.info("처리 대기 중인 수리 기록이 없습니다.")
            else:
                repair_ids = [str(r.get("수리 ID", "")) for r in active_repairs]
                with st.form("repair_finish_form"):
                    selected_repair_id = st.selectbox("완료할 수리 기록", repair_ids, format_func=lambda x: f"{x} · {(next((r for r in active_repairs if str(r.get('수리 ID')) == x), {})).get('기자재명', '')} · {(next((r for r in active_repairs if str(r.get('수리 ID')) == x), {})).get('수리 내용', '')}")
                    finish_date = st.date_input("수리 완료일", value=today_kst())
                    finish_note = st.text_input("완료 메모")
                    finish_submit = st.form_submit_button("수리 완료 처리", use_container_width=True)
                if finish_submit:
                    try:
                        handle_repair_finish(store, selected_repair_id, finish_date, finish_note, operator)
                        st.success("수리 완료 처리했습니다.")
                        st.rerun()
                    except Exception as exc:
                        st.error(f"수리 완료 처리 실패: {exc}")
    with loss_tab:
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("분실 처리")
            if active_items:
                with st.form("loss_form"):
                    loss_item_id = st.selectbox("분실 기자재", [str(r.get("관리번호")) for r in active_items], format_func=lambda x: fmt_item(find_item(active_items, x) or {}))
                    loss_qty = st.number_input("분실 수량", min_value=1, value=1, step=1)
                    loss_reason = st.text_input("분실 사유 또는 메모")
                    loss_submit = st.form_submit_button("분실 처리", use_container_width=True)
                if loss_submit:
                    try:
                        set_lost_quantity(store, loss_item_id, int(loss_qty), operator, loss_reason or "분실 확인")
                        st.success("분실 수량을 반영했습니다.")
                        st.rerun()
                    except Exception as exc:
                        st.error(f"분실 처리 실패: {exc}")
        with c2:
            st.subheader("분실 복구")
            lost_items = [r for r in active_items if as_int(r.get("분실 수량")) > 0]
            if lost_items:
                with st.form("found_form"):
                    found_item_id = st.selectbox("복구할 기자재", [str(r.get("관리번호")) for r in lost_items], format_func=lambda x: f"{x} · {(find_item(lost_items, x) or {}).get('기자재명', '')} (분실 {as_int((find_item(lost_items, x) or {}).get('분실 수량'))}개)")
                    found_max = as_int((find_item(lost_items, found_item_id) or {}).get("분실 수량"))
                    found_qty = st.number_input("복구 수량", min_value=1, max_value=max(1, found_max), value=1, step=1)
                    found_reason = st.text_input("복구 메모")
                    found_submit = st.form_submit_button("분실 복구", use_container_width=True)
                if found_submit:
                    try:
                        set_lost_quantity(store, found_item_id, -int(found_qty), operator, found_reason or "분실품 회수")
                        st.success("분실 수량을 복구했습니다.")
                        st.rerun()
                    except Exception as exc:
                        st.error(f"분실 복구 실패: {exc}")
            else:
                st.info("분실 처리된 기자재가 없습니다.")
        st.caption("분실 수량을 늘리거나 줄이면 변경이력에 자동으로 기록됩니다. 폐기 상태는 기자재 목록의 기본 상태에서 설정할 수 있습니다.")
    with history_tab:
        if data["수리기록"]:
            repair_df = to_df(data["수리기록"], ["수리 ID", "관리번호", "기자재명", "수리 접수일", "수리 내용", "수리 수량", "수리 상태", "수리 완료일", "처리 메모"])
            st.dataframe(repair_df.sort_values("수리 접수일", ascending=False), hide_index=True, use_container_width=True)
        else:
            st.info("수리 기록이 없습니다.")


def render_audit(data: dict[str, list[dict[str, Any]]]) -> None:
    st.title("변경 이력")
    logs = data["변경이력"]
    items = data["기자재목록"]
    ids = sorted({str(r.get("관리번호", "")).strip() for r in items if str(r.get("관리번호", "")).strip()})
    operators = sorted({str(r.get("작업자", "")).strip() for r in logs if str(r.get("작업자", "")).strip()})
    c1, c2, c3 = st.columns(3)
    with c1:
        filter_id = st.selectbox("관리번호 필터", ["전체"] + ids)
    with c2:
        task_types = sorted({str(r.get("작업 유형", "")).strip() for r in logs if str(r.get("작업 유형", "")).strip()})
        filter_task = st.selectbox("작업 유형", ["전체"] + task_types)
    with c3:
        filter_operator = st.selectbox("작업자", ["전체"] + operators)
    filtered = list(logs)
    if filter_id != "전체":
        filtered = [r for r in filtered if str(r.get("관리번호", "")) == filter_id]
    if filter_task != "전체":
        filtered = [r for r in filtered if str(r.get("작업 유형", "")) == filter_task]
    if filter_operator != "전체":
        filtered = [r for r in filtered if str(r.get("작업자", "")) == filter_operator]
    start_date, end_date = st.columns(2)
    with start_date:
        start = st.date_input("변경일 시작", value=today_kst().replace(day=1))
    with end_date:
        end = st.date_input("변경일 종료", value=today_kst())
    if start > end:
        st.warning("시작일이 종료일보다 늦습니다.")
    else:
        filtered = [r for r in filtered if (parse_date(r.get("변경 시각")) is None or start <= parse_date(r.get("변경 시각")) <= end)]
    filtered.sort(key=lambda r: str(r.get("변경 시각", "")), reverse=True)
    st.caption(f"조회 결과 {len(filtered)}건")
    st.dataframe(to_df(filtered, SHEET_HEADERS["변경이력"]), hide_index=True, use_container_width=True)


def render_exports(data: dict[str, list[dict[str, Any]]]) -> None:
    st.title("데이터 내보내기")
    st.write("필터 조건에 해당하는 기자재 목록 또는 전체 대여·반납 내역을 CSV 파일로 내려받을 수 있습니다. 한글 호환을 위해 UTF-8-SIG 인코딩을 사용합니다.")
    items = enrich_items(data["기자재목록"], data["대여반납기록"], data["수리기록"])
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("기자재 목록 CSV")
        query = st.text_input("기자재명 또는 관리번호", key="export_query")
        categories = sorted({str(r.get("분류", "")).strip() for r in items if str(r.get("분류", "")).strip()})
        category = st.selectbox("분류 필터", ["전체"] + categories, key="export_category")
        locations = sorted({str(r.get("보관 위치", "")).strip() for r in items if str(r.get("보관 위치", "")).strip()})
        location = st.selectbox("보관 위치 필터", ["전체"] + locations, key="export_location")
        state = st.selectbox("상태 필터", ["전체", "사용 가능", "대여 중", "수리 필요", "분실", "폐기", "비활성"], key="export_state")
        include_inactive = st.checkbox("비활성 기자재 포함", key="export_inactive")
        selected = filter_items(items, query, category, location, state, include_inactive)
        item_cols = ["관리번호", "기자재명", "분류", "세부 품목", "보유 수량", "대여 중 수량", "사용 가능 수량", "수리 중 수량", "분실 수량", "보관 위치", "표시 상태", "구입일", "구입 가격", "제조사 및 모델명", "비고", "등록일", "활성화 여부"]
        st.caption(f"내보낼 기자재: {len(selected)}건")
        safe_csv(to_df(selected, item_cols), f"기자재목록_{today_kst().isoformat()}.csv", "기자재 목록 CSV 다운로드")
    with c2:
        st.subheader("대여·반납 기록 CSV")
        loan_df = to_df(data["대여반납기록"], SHEET_HEADERS["대여반납기록"])
        if not loan_df.empty and "처리일시" in loan_df:
            loan_df = loan_df.sort_values("처리일시", ascending=False)
        st.caption(f"내보낼 거래 이벤트: {len(loan_df)}건 (대여와 각 반납 이벤트가 별도 행으로 저장됨)")
        safe_csv(loan_df, f"대여반납기록_{today_kst().isoformat()}.csv", "대여·반납 기록 CSV 다운로드")
        st.subheader("수리 기록 CSV")
        safe_csv(to_df(data["수리기록"], SHEET_HEADERS["수리기록"]), f"수리기록_{today_kst().isoformat()}.csv", "수리 기록 CSV 다운로드")
        st.subheader("변경 이력 CSV")
        safe_csv(to_df(data["변경이력"], SHEET_HEADERS["변경이력"]), f"변경이력_{today_kst().isoformat()}.csv", "변경 이력 CSV 다운로드")


def authenticate() -> bool:
    try:
        password = str(st.secrets["app_password"])
    except Exception:
        st.error("앱 접속 비밀번호 설정이 없습니다. Secrets에 `app_password = \"충분히_긴_비밀번호\"`를 설정해 주세요.")
        st.stop()
        return False
    if not password or password.lower().startswith("change-this"):
        st.error("기본 예시 비밀번호가 설정되어 있습니다. `.streamlit/secrets.toml` 또는 Streamlit Cloud Secrets에서 강력한 접속 비밀번호로 변경해 주세요.")
        st.stop()
    if st.session_state.get("authenticated"):
        return True
    st.title("🧰 기술실 기자재 관리")
    st.write("학교 기자재 관리 시스템입니다. 사용 권한이 있는 사람만 접속해 주세요.")
    with st.form("app_login_form"):
        entered = st.text_input("접속 비밀번호", type="password")
        submitted = st.form_submit_button("접속", type="primary", use_container_width=True)
    if submitted:
        if py_secrets.compare_digest(entered, password):
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("비밀번호가 일치하지 않습니다.")
    st.caption("공유 비밀번호는 사용자별 권한을 구분하지 않습니다. 학교 내부에서 승인된 사람에게만 공유하세요.")
    return False


def main() -> None:
    if not authenticate():
        st.stop()
    try:
        store = get_store()
        data = store.load_all()
    except SpreadsheetNotFound:
        st.error("Google Sheets 문서를 찾을 수 없습니다. `GOOGLE_SHEET_URL`이 정확한지, 서비스 계정 이메일에 문서를 공유했는지 확인해 주세요.")
        st.stop()
    except WorksheetNotFound:
        st.error("필요한 시트 탭을 찾을 수 없습니다. 앱을 다시 실행해 초기화하거나 Google Sheets 문서의 권한을 확인해 주세요.")
        st.stop()
    except APIError as exc:
        st.error("Google Sheets API 요청에 실패했습니다. 서비스 계정 권한, Google Sheets API 활성화 여부, 호출 제한을 확인해 주세요.")
        st.code(str(exc))
        st.stop()
    except Exception as exc:
        message = str(exc)
        if "permission" in message.lower() or "403" in message:
            st.error("Google Sheets 접근 권한이 없습니다. 서비스 계정 이메일을 해당 문서의 편집자로 공유했는지 확인해 주세요.")
        elif "private_key" in message.lower() or "credential" in message.lower() or "secrets" in message.lower():
            st.error("인증 설정을 읽거나 검증하지 못했습니다. Secrets의 `[gcp_service_account]` 항목을 다시 확인해 주세요.")
        else:
            st.error(f"앱을 초기화하지 못했습니다: {message}")
        st.stop()

    st.sidebar.title("🧰 기술실 관리")
    st.sidebar.caption("Google Sheets 실시간 연동")
    st.sidebar.text_input("작업자", key="operator_name", placeholder="예: 홍길동")
    if st.sidebar.button("최신 데이터 새로고침", use_container_width=True):
        st.rerun()
    if st.sidebar.button("로그아웃", use_container_width=True):
        st.session_state["authenticated"] = False
        st.rerun()
    page = st.sidebar.radio(
        "메뉴",
        ["대시보드", "기자재 목록", "기자재 등록", "대여 및 반납", "수리 및 분실 관리", "변경 이력", "데이터 내보내기"],
    )
    st.sidebar.divider()
    st.sidebar.caption("보호된 정보가 포함될 수 있으므로 접속 비밀번호를 외부에 공개하지 마세요.")

    operator = operator_name()
    try:
        if page == "대시보드":
            render_dashboard(data)
        elif page == "기자재 목록":
            render_inventory(data, store, operator)
        elif page == "기자재 등록":
            render_register(data, store, operator)
        elif page == "대여 및 반납":
            render_loans(data, store, operator)
        elif page == "수리 및 분실 관리":
            render_repair_loss(data, store, operator)
        elif page == "변경 이력":
            render_audit(data)
        elif page == "데이터 내보내기":
            render_exports(data)
    except Exception as exc:
        st.error(f"화면을 처리하는 중 문제가 발생했습니다: {exc}")
        st.caption("데이터가 예상과 다르게 보이면 최신 데이터 새로고침 후 다시 확인하고, 필요한 경우 변경이력과 Google Sheets 원본을 점검해 주세요.")


if __name__ == "__main__":
    main()
