from __future__ import annotations

import os
import re
import traceback
import threading
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from datetime import datetime, timedelta

import openpyxl
from fastapi import FastAPI, HTTPException, Depends, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from fastapi.responses import HTMLResponse
from openpyxl.worksheet.worksheet import Worksheet
import secrets

# OneDrive 동기화 모듈 import
try:
    from onedrive_sync import sync_onedrive_file
except ImportError:
    # 로컬 개발 환경에서는 동기화 없이 진행
    def sync_onedrive_file(*args, **kwargs):
        return True

BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILENAME = "★26SS MLB 생산스케쥴_DASHBOARD.xlsx"
DEFAULT_WORKBOOK = BASE_DIR / EXCEL_FILENAME

# 환경 변수
FILE_PATH = Path(os.getenv("SUMMARY_EXCEL", str(DEFAULT_WORKBOOK)))
SHEET_NAME = os.getenv("SUMMARY_SHEET", "수량 기준")
ONEDRIVE_SHARE_LINK = os.getenv("ONEDRIVE_FILE_URL", "")
DASHBOARD_PASSWORD = os.getenv("DASHBOARD_PASSWORD", "admin123")  # 기본값, 배포 시 변경 필수
SYNC_INTERVAL = int(os.getenv("SYNC_INTERVAL", "3600"))  # 기본 1시간

# 기본값 (주차를 찾을 수 없을 때 사용)
DEFAULT_WEEK1 = 48
DEFAULT_WEEK2 = 49

# 고정 컬럼: 총 수량
TOTAL_QTY_COL = "C"
TOTAL_QTY_COL_DETAIL = "P"

# 비밀번호 인증
security = HTTPBasic()

def verify_password(credentials: HTTPBasicCredentials = Depends(security)) -> bool:
    """비밀번호 인증"""
    correct_password = DASHBOARD_PASSWORD
    is_correct = secrets.compare_digest(credentials.password, correct_password)
    if not is_correct:
        raise HTTPException(
            status_code=401,
            detail="Invalid password",
            headers={"WWW-Authenticate": "Basic"},
        )
    return True


def _extract_week_from_header(cell_value: Any) -> Optional[int]:
    """셀 값에서 주차 번호를 추출합니다. 'xx주차' 형식을 찾습니다."""
    if cell_value is None:
        return None
    
    # 문자열로 변환
    text = str(cell_value).strip()
    
    # 'xx주차' 패턴 찾기 (예: "49주차", "49 주차" 등)
    match = re.search(r'(\d+)\s*주차', text)
    if match:
        try:
            return int(match.group(1))
        except (ValueError, AttributeError):
            return None
    
    return None


def _find_week_numbers(ws: Worksheet) -> Tuple[int, int]:
    """엑셀 시트의 헤더 행에서 주차 번호를 찾습니다."""
    week_numbers = []
    
    # 헤더 행을 확인 (일반적으로 1-4행)
    for row in range(1, 5):
        # D열부터 L열까지 확인 (첫 번째 주차 그룹)
        for col_letter in ["D", "E", "F", "G", "H", "I", "J", "K", "L"]:
            try:
                cell_value = ws[f"{col_letter}{row}"].value
                week_num = _extract_week_from_header(cell_value)
                if week_num is not None and week_num not in week_numbers:
                    week_numbers.append(week_num)
                    if len(week_numbers) >= 2:
                        break
            except Exception:
                continue
        if len(week_numbers) >= 2:
            break
    
    # 주차를 찾았으면 정렬 (첫 번째가 금주, 두 번째가 차주)
    if len(week_numbers) >= 2:
        week_numbers.sort()
        return tuple(week_numbers[:2])
    elif len(week_numbers) == 1:
        # 한 개만 찾았으면 차주는 +1
        return (week_numbers[0], week_numbers[0] + 1)
    else:
        # 찾지 못했으면 기본값 사용
        return (DEFAULT_WEEK1, DEFAULT_WEEK2)


def _col_num_to_letter(n: int) -> str:
    """열 번호를 엑셀 열 문자로 변환 (1=A, 2=B, ..., 27=AA)"""
    result = ""
    while n > 0:
        n -= 1
        result = chr(65 + (n % 26)) + result
        n //= 26
    return result


def _col_letter_to_num(col: str) -> int:
    """엑셀 열 문자를 열 번호로 변환 (A=1, B=2, ..., AA=27)"""
    result = 0
    for char in col.upper():
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result


def _build_value_columns(week1: int, week2: int, total_qty_col: str = "C", first_week_target_col: str = "D") -> Tuple[Tuple[str, str], ...]:
    """주차 번호에 따라 VALUE_COLUMNS를 동적으로 생성합니다.
    
    Args:
        week1: 첫 번째 주차 번호 (금주)
        week2: 두 번째 주차 번호 (차주)
        total_qty_col: 총 수량 컬럼 (예: "C" 또는 "P")
        first_week_target_col: 첫 번째 주차의 target 컬럼 (예: "D" 또는 "Q")
    """
    # 첫 번째 주차 target 컬럼의 열 번호
    base_col_num = _col_letter_to_num(first_week_target_col)
    
    columns = [
        ("total_qty", total_qty_col),
        (f"target_{week1}", _col_num_to_letter(base_col_num)),      # D 또는 Q
        (f"actual_{week1}", _col_num_to_letter(base_col_num + 1)),  # E 또는 R
        (f"diff_{week1}", _col_num_to_letter(base_col_num + 2)),    # F 또는 S
        (f"target_{week1}_pct", _col_num_to_letter(base_col_num + 3)),  # G 또는 T
        (f"actual_{week1}_pct", _col_num_to_letter(base_col_num + 4)),  # H 또는 U
        (f"target_{week2}", _col_num_to_letter(base_col_num + 5)),      # I 또는 V
        (f"actual_{week2}", _col_num_to_letter(base_col_num + 6)),      # J 또는 W
        (f"target_{week2}_pct", _col_num_to_letter(base_col_num + 7)),  # K 또는 X
        (f"actual_{week2}_pct", _col_num_to_letter(base_col_num + 8)),  # L 또는 Y
    ]
    
    return tuple(columns)


# 동적으로 생성할 예정이므로 None으로 초기화
VALUE_COLUMNS: Optional[Tuple[Tuple[str, str], ...]] = None
DETAIL_VALUE_COLUMNS: Optional[Tuple[Tuple[str, str], ...]] = None
WEEK1: Optional[int] = None
WEEK2: Optional[int] = None

# 데이터 캐시 시스템
_data_cache: Optional[Dict[str, Any]] = None
_cache_timestamp: Optional[datetime] = None
_cache_lock = threading.Lock()

# 업데이트 시간 설정 (오전 11시)
UPDATE_HOUR = 11
UPDATE_MINUTE = 0

BLOCK_LAYOUT = (
    ("nations", {"rows": range(5, 10), "label_key": "code", "label_col": "B"}),
    ("items", {"rows": range(15, 19), "label_key": "item", "label_col": "B"}),
    (
        "categories",
        {
            "rows": range(24, 80),
            "label_key": "category",
            "label_col": "B",
            "stop_on_blank": True,
            "blank_tolerance": 2,
        },
    ),
    (
        "sub_categories",
        {
            "rows": range(5, 60),
            "label_key": "subcategory",
            "label_col": "O",
            "use_detail_columns": True,  # DETAIL_VALUE_COLUMNS 사용 플래그
            "stop_on_blank": True,
            "blank_tolerance": 3,
        },
    ),
)

app = FastAPI(title="26SS Quantity Summary API")

# CORS 설정 - 프론트엔드 도메인만 허용하도록 제한 가능
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 배포 후 프론트엔드 도메인으로 제한 권장
    allow_methods=["*"],
    allow_headers=["*"],
)


def _extract_block(ws: Worksheet, config: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Generic reader for a contiguous table that shares the same value columns."""

    payload: List[Dict[str, Any]] = []
    blank_streak = 0
    # 동적 컬럼이 설정되지 않은 경우 기본값 사용 (하위 호환성)
    default_columns = _build_value_columns(DEFAULT_WEEK1, DEFAULT_WEEK2, "C", "D")
    
    # value_columns가 config에 명시적으로 설정되어 있으면 그것을 우선 사용
    if "value_columns" in config and config["value_columns"] is not None:
        columns = config["value_columns"]
    else:
        columns = VALUE_COLUMNS or default_columns
    
    label_col = config.get("label_col", "B")
    label_key = config.get("label_key", "label")
    rows = config.get("rows", [])

    # 배치로 셀 읽기 최적화 (메모리 효율성 향상)
    for row in config["rows"]:
        try:
            label = ws[f"{label_col}{row}"].value

            if label in (None, ""):
                if config.get("stop_on_blank"):
                    blank_streak += 1
                    if blank_streak >= config.get("blank_tolerance", 1):
                        break
                continue

            blank_streak = 0
            entry = {label_key: label}

            # 한 행의 모든 셀을 한 번에 읽기 (메모리 효율성)
            for key, column in columns:
                try:
                    cell_value = ws[f"{column}{row}"].value
                    
                    # #VALUE! 같은 엑셀 오류 문자열 처리
                    if isinstance(cell_value, str) and cell_value.startswith("#"):
                        entry[key] = None
                    elif cell_value is None:
                        entry[key] = None
                    elif isinstance(cell_value, (int, float)):
                        entry[key] = cell_value
                    elif isinstance(cell_value, str):
                        cleaned = cell_value.strip().replace(",", "").replace(" ", "")
                        if not cleaned or cleaned.startswith("#"):
                            entry[key] = None
                        else:
                            try:
                                entry[key] = float(cleaned) if "." in cleaned else int(cleaned)
                            except (ValueError, TypeError):
                                entry[key] = None
                    else:
                        entry[key] = cell_value
                except Exception:
                    entry[key] = None

            payload.append(entry)
        except Exception:
            continue

    return payload


def ensure_excel_file() -> Path:
    """엑셀 파일이 존재하는지 확인하고, OneDrive에서 동기화합니다."""
    # OneDrive 링크가 설정되어 있고 파일이 없으면 동기화 시도
    if ONEDRIVE_SHARE_LINK and not FILE_PATH.exists():
        sync_onedrive_file(ONEDRIVE_SHARE_LINK, FILE_PATH, sync_interval=SYNC_INTERVAL, force_download=False)
    
    # 파일이 여전히 없으면 에러
    if not FILE_PATH.exists():
        raise FileNotFoundError(f"Excel file not found at {FILE_PATH}. OneDrive sync may have failed.")
    
    return FILE_PATH


def should_update_cache() -> bool:
    """캐시를 업데이트해야 하는지 확인합니다.
    - 캐시가 없으면 업데이트
    - 오전 11시 이후이고 오늘 업데이트하지 않았으면 업데이트
    - 단순 시간 기반 체크만 수행 (파일 변경 시간 체크 제거)
    """
    global _cache_timestamp
    
    now = datetime.now()
    
    # 캐시가 없으면 업데이트
    if _cache_timestamp is None:
        return True
    
    # 오늘 날짜
    today = now.date()
    cache_date = _cache_timestamp.date()
    
    # 오늘 업데이트했으면 스킵 (다음날 11시까지 캐시 사용)
    if cache_date == today:
        return False
    
    # 오늘 업데이트하지 않았고, 오전 11시 이후면 업데이트
    if now.hour >= UPDATE_HOUR:
        return True
    
    return False


def update_cache() -> None:
    """캐시를 업데이트합니다. 매일 11시에만 실행됩니다."""
    global _data_cache, _cache_timestamp
    
    # 기존 캐시 백업 (에러 발생 시 복구용)
    old_cache = _data_cache
    old_timestamp = _cache_timestamp
    
    with _cache_lock:
        try:
            # OneDrive 동기화 (파일이 없는 경우만)
            if ONEDRIVE_SHARE_LINK and not FILE_PATH.exists():
                sync_onedrive_file(ONEDRIVE_SHARE_LINK, FILE_PATH, sync_interval=3600, force_download=False)
            
            # 데이터 로드
            quantity_data = load_summary("수량 기준")
            style_count_data = load_summary("스타일수 기준")
            
            # 데이터 유효성 검사
            if not quantity_data or not isinstance(quantity_data, dict):
                raise ValueError("Invalid quantity data")
            if not style_count_data or not isinstance(style_count_data, dict):
                raise ValueError("Invalid style_count data")
            
            # 필수 키 확인
            required_keys = ["nations", "items", "categories", "week_info"]
            if not all(key in quantity_data for key in required_keys):
                raise ValueError("Missing required keys in quantity data")
            if not all(key in style_count_data for key in required_keys):
                raise ValueError("Missing required keys in style_count data")
            
            _data_cache = {
                "quantity": quantity_data,
                "style_count": style_count_data,
            }
            _cache_timestamp = datetime.now()
            
        except Exception as e:
            # 에러 발생 시 기존 캐시 복구
            if old_cache is not None:
                _data_cache = old_cache
                _cache_timestamp = old_timestamp
            else:
                # 기존 캐시도 없으면 None 유지 (다음 요청 시 직접 로드)
                pass


# 백그라운드 업데이트 플래그
_updating_cache = False
_update_lock = threading.Lock()

def get_cached_data(sheet_name: str) -> Dict[str, Any]:
    """캐시된 데이터를 반환합니다. 빠른 응답을 위해 캐시를 우선 사용합니다."""
    global _data_cache, _updating_cache
    
    # 캐시가 있으면 즉시 반환 (가장 빠름 - 파일 읽기 없음)
    if _data_cache is not None:
        if sheet_name == "수량 기준":
            cached = _data_cache.get("quantity")
            if cached and isinstance(cached, dict) and len(cached) > 0:
                return cached
        elif sheet_name == "스타일수 기준":
            cached = _data_cache.get("style_count")
            if cached and isinstance(cached, dict) and len(cached) > 0:
                return cached
    
    # 캐시가 없으면 직접 로드 (최초 로드 시에만 - 이후에는 백그라운드에서 업데이트)
    try:
        return load_summary(sheet_name)
    except Exception as e:
        # 최소한의 구조라도 반환하여 프론트엔드 에러 방지
        return {
            "nations": [],
            "items": [],
            "categories": [],
            "sub_categories": [],
            "week_info": {
                "current_week": DEFAULT_WEEK1,
                "next_week": DEFAULT_WEEK2,
            },
            "sheet_name": sheet_name,
        }


def load_summary(sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """엑셀 파일에서 데이터를 로드하고 주차 정보를 포함하여 반환합니다."""
    global VALUE_COLUMNS, DETAIL_VALUE_COLUMNS, WEEK1, WEEK2
    
    target_sheet = sheet_name or SHEET_NAME
    
    # 파일 존재 확인 및 동기화
    excel_path = ensure_excel_file()
    
    workbook = None
    try:
        # read_only=True로 메모리 사용량 최소화
        workbook = openpyxl.load_workbook(excel_path, data_only=True, read_only=True, keep_links=False)
    except PermissionError as exc:
        error_msg = (
            f"엑셀 파일에 접근할 수 없습니다. 파일이 다른 프로그램에서 열려있거나 "
            f"OneDrive 동기화 중일 수 있습니다. 원본 에러: {exc}"
        )
        raise RuntimeError(error_msg) from exc
    except Exception as exc:
        raise RuntimeError(f"Failed to open workbook: {exc}") from exc

    try:
        available_sheets = workbook.sheetnames
        
        if target_sheet not in available_sheets:
            raise RuntimeError(
                f"Worksheet '{target_sheet}' not found. Available sheets: {', '.join(available_sheets)}"
            )
        
        sheet = workbook[target_sheet]
        
    except KeyError as exc:
        if workbook:
            workbook.close()
        raise RuntimeError(f"Worksheet '{target_sheet}' not found in workbook") from exc

    try:
        # 헤더에서 주차 정보 추출
        WEEK1, WEEK2 = _find_week_numbers(sheet)
        
        # 주차 정보에 따라 동적으로 컬럼 생성
        VALUE_COLUMNS = _build_value_columns(WEEK1, WEEK2, "C", "D")
        DETAIL_VALUE_COLUMNS = _build_value_columns(WEEK1, WEEK2, "P", "Q")
        
        data = {}
        for name, config in BLOCK_LAYOUT:
            try:
                current_config = config.copy()
                if config.get("use_detail_columns"):
                    current_config["value_columns"] = DETAIL_VALUE_COLUMNS
                elif "value_columns" not in current_config:
                    current_config["value_columns"] = VALUE_COLUMNS
                
                extracted = _extract_block(sheet, current_config)
                data[name] = extracted
            except Exception as e:
                data[name] = []
        
        result = {
            **data,
            "week_info": {
                "current_week": WEEK1,
                "next_week": WEEK2,
            },
            "sheet_name": target_sheet,
        }
        
        return result
    except Exception as e:
        error_msg = f"Unexpected error in load_summary: {str(e)}"
        raise RuntimeError(error_msg) from e
    finally:
        if workbook:
            try:
                workbook.close()
            except Exception:
                pass


@app.get("/health")
def healthcheck() -> Dict[str, str]:
    return {"status": "ok"}


@app.get("/api/sheets")
def list_sheets() -> Dict[str, Any]:
    """엑셀 파일의 시트 목록을 반환 (디버깅용)"""
    try:
        excel_path = ensure_excel_file()
        
        workbook = openpyxl.load_workbook(excel_path, read_only=True)
        sheets = workbook.sheetnames
        workbook.close()
        
        return {
            "file_path": str(excel_path),
            "sheets": sheets,
            "current_sheet": SHEET_NAME,
            "sheet_exists": SHEET_NAME in sheets,
            "file_exists": True,
            "file_readable": True
        }
    except Exception as exc:
        return {
            "error": str(exc),
            "error_type": type(exc).__name__,
            "sheets": []
        }


@app.get("/api/quantity")
def get_quantity_summary(_: bool = Depends(verify_password)) -> Dict[str, Any]:
    """수량 기준 데이터를 주차 정보와 함께 반환합니다. (캐시 사용)"""
    try:
        return get_cached_data("수량 기준")
    except Exception as exc:
        error_detail = f"Unexpected error: {str(exc)}"
        print(f"Unexpected error in /api/quantity: {error_detail}")
        print(traceback.format_exc())
        raise HTTPException(status_code=500, detail=error_detail) from exc


@app.get("/api/style-count")
def get_style_count_summary(_: bool = Depends(verify_password)) -> Dict[str, Any]:
    """스타일수 기준 데이터를 주차 정보와 함께 반환합니다. (캐시 사용)"""
    try:
        return get_cached_data("스타일수 기준")
    except Exception as exc:
        error_detail = f"Unexpected error: {str(exc)}"
        print(f"Unexpected error in /api/style-count: {error_detail}")
        print(traceback.format_exc())
        raise HTTPException(status_code=500, detail=error_detail) from exc


@app.on_event("startup")
async def startup_event():
    """서버 시작 시 초기 캐시 업데이트 및 백그라운드 스레드 시작"""
    # 초기 캐시 업데이트 (캐시가 없으면)
    if _data_cache is None:
        update_cache()
    
    # 백그라운드 스레드에서 매일 11시에 업데이트 체크
    def background_update_check():
        """백그라운드에서 매일 11시에 업데이트를 수행합니다."""
        global _updating_cache
        
        while True:
            now = datetime.now()
            
            # 다음 업데이트 시간 계산 (오늘 11시 또는 내일 11시)
            if now.hour < UPDATE_HOUR:
                # 오늘 11시까지 대기
                next_update = now.replace(hour=UPDATE_HOUR, minute=UPDATE_MINUTE, second=0, microsecond=0)
            else:
                # 내일 11시까지 대기
                next_update = (now + timedelta(days=1)).replace(hour=UPDATE_HOUR, minute=UPDATE_MINUTE, second=0, microsecond=0)
            
            wait_seconds = (next_update - now).total_seconds()
            
            # 다음 업데이트 시간까지 대기 (최대 1시간씩 체크)
            while wait_seconds > 0:
                sleep_time = min(wait_seconds, 3600)  # 최대 1시간씩 체크
                time.sleep(sleep_time)
                wait_seconds -= sleep_time
                
                # 중간에 업데이트 시간이 되었는지 확인
                now = datetime.now()
                if now.hour >= UPDATE_HOUR and should_update_cache():
                    break
            
            # 업데이트 시간이 되었는지 확인
            try:
                with _update_lock:
                    if should_update_cache() and not _updating_cache:
                        _updating_cache = True
                        update_cache()
                        _updating_cache = False
            except Exception:
                _updating_cache = False
                pass  # 에러 발생해도 계속 실행
    
    thread = threading.Thread(target=background_update_check, daemon=True)
    thread.start()

