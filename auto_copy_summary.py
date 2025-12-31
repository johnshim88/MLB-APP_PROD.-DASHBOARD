"""
SUMMARY 파일 자동 복사 스크립트
회사 OneDrive의 SUMMARY 파일을 Google Drive로 자동 복사합니다.
Windows 작업 스케줄러와 함께 사용하여 주기적으로 실행할 수 있습니다.
"""
import shutil
from pathlib import Path
import time
from datetime import datetime
import sys

# 경로 설정
# 회사 OneDrive 경로 (원본 파일)
SOURCE_FILE = Path(r"C:\Users\AD1060\OneDrive - F&F\F_SO_ MLB 소싱팀 - 26SS\생산스케쥴\DASHBOARD\★26SS MLB 생산스케쥴_DASHBOARD_V2.xlsx")

# Google Drive 경로 (복사본 저장 위치)
DEST_FILE = Path(r"G:\내 드라이브\MLB PROD DASHBOARD\★26SS MLB 생산스케쥴_DASHBOARD_V2.xlsx")

# 로그 파일 경로 (선택사항)
LOG_FILE = Path(r"G:\내 드라이브\MLB PROD DASHBOARD\copy_log.txt")

# 강제 덮어쓰기 옵션 (True: 항상 복사, False: 최신일 때만 복사)
FORCE_COPY = True  # True로 설정하면 항상 덮어쓰기


def log_message(message: str, to_console: bool = True, to_file: bool = True):
    """메시지를 콘솔과 로그 파일에 기록"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_msg = f"[{timestamp}] {message}"
    
    if to_console:
        print(log_msg)
    
    if to_file and LOG_FILE:
        try:
            LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
            with open(LOG_FILE, "a", encoding="utf-8") as f:
                f.write(log_msg + "\n")
        except Exception as e:
            print(f"로그 파일 쓰기 오류: {e}")


def copy_if_newer():
    """원본 파일을 복사 (FORCE_COPY 옵션에 따라 강제 복사 또는 조건부 복사)"""
    try:
        # 원본 파일 존재 확인
        if not SOURCE_FILE.exists():
            log_message(f"원본 파일을 찾을 수 없습니다: {SOURCE_FILE}")
            return False
        
        # Google Drive 폴더 생성
        DEST_FILE.parent.mkdir(parents=True, exist_ok=True)
        
        # 강제 복사 모드인 경우
        if FORCE_COPY:
            # 파일 복사 (덮어쓰기)
            shutil.copy2(SOURCE_FILE, DEST_FILE)
            log_message(f"파일 강제 복사 완료: {DEST_FILE}")
            log_message(f"  원본 수정 시간: {datetime.fromtimestamp(SOURCE_FILE.stat().st_mtime)}")
            log_message(f"  복사본 수정 시간: {datetime.fromtimestamp(DEST_FILE.stat().st_mtime)}")
            return True
        
        # 조건부 복사 모드 (기존 로직)
        should_copy = False
        reason = ""
        
        if not DEST_FILE.exists():
            should_copy = True
            reason = "복사본 파일이 없음"
        else:
            # 파일 수정 시간 비교
            source_mtime = SOURCE_FILE.stat().st_mtime
            dest_mtime = DEST_FILE.stat().st_mtime
            
            if source_mtime > dest_mtime:
                should_copy = True
                reason = f"원본이 더 최신 (원본: {datetime.fromtimestamp(source_mtime)}, 복사본: {datetime.fromtimestamp(dest_mtime)})"
            else:
                reason = "복사본이 이미 최신 상태"
        
        if should_copy:
            # 파일 복사
            shutil.copy2(SOURCE_FILE, DEST_FILE)
            log_message(f"파일 복사 완료: {DEST_FILE}")
            log_message(f"  이유: {reason}")
            log_message(f"  원본 수정 시간: {datetime.fromtimestamp(SOURCE_FILE.stat().st_mtime)}")
            log_message(f"  복사본 수정 시간: {datetime.fromtimestamp(DEST_FILE.stat().st_mtime)}")
            return True
        else:
            log_message(f"파일 복사 건너뜀: {reason}")
            return False
            
    except PermissionError as e:
        log_message(f"권한 오류: 파일이 다른 프로그램에서 열려있을 수 있습니다. {e}")
        return False
    except Exception as e:
        log_message(f"오류 발생: {e}")
        import traceback
        log_message(f"상세 오류:\n{traceback.format_exc()}")
        return False


def main():
    """메인 함수"""
    log_message("=" * 60)
    log_message("SUMMARY 파일 자동 복사 스크립트 시작")
    log_message(f"원본 파일: {SOURCE_FILE}")
    log_message(f"복사본 파일: {DEST_FILE}")
    log_message(f"강제 복사 모드: {'ON' if FORCE_COPY else 'OFF'}")
    log_message("=" * 60)
    
    success = copy_if_newer()
    
    log_message("=" * 60)
    if success:
        log_message("스크립트 완료: 파일 복사 성공")
    else:
        log_message("스크립트 완료: 파일 복사 불필요 또는 실패")
    log_message("=" * 60)
    
    return 0 if success else 1


if __name__ == "__main__":
    try:
        exit_code = main()
        sys.exit(exit_code)
    except KeyboardInterrupt:
        log_message("사용자에 의해 중단됨")
        sys.exit(1)
    except Exception as e:
        log_message(f"예상치 못한 오류: {e}")
        import traceback
        log_message(traceback.format_exc())
        sys.exit(1)

