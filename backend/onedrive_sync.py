"""OneDrive 파일 동기화 모듈"""
import os
import time
import httpx
from pathlib import Path
from typing import Optional
import traceback


def download_from_onedrive_share_link(share_link: str, output_path: Path) -> bool:
    """OneDrive/SharePoint 공유 링크에서 파일을 다운로드합니다.
    
    Args:
        share_link: OneDrive/SharePoint 공유 링크 
            (예: https://1drv.ms/x/..., https://onedrive.live.com/..., https://*.sharepoint.com/...)
        output_path: 저장할 파일 경로
    
    Returns:
        성공 여부
    """
    try:
        # URL에 프로토콜이 없으면 https:// 추가
        if share_link and not share_link.startswith(("http://", "https://")):
            share_link = "https://" + share_link
            print(f"Added protocol to URL: {share_link}")
        
        download_url = None
        
        # SharePoint 링크 처리
        if "sharepoint.com" in share_link:
            print(f"Detected SharePoint link: {share_link}")
            # SharePoint 링크를 직접 다운로드 링크로 변환
            # :f: (폴더) 링크는 :x: (파일) 링크로 변환 시도
            if ":f:" in share_link:
                print("Warning: Folder link detected. Please use file direct link (with :x:) instead.")
                # 폴더 링크는 파일 링크로 변환 시도 (작동하지 않을 수 있음)
                share_link = share_link.replace(":f:", ":x:")
            
            # SharePoint 직접 다운로드 링크로 변환
            # SharePoint 링크는 리다이렉트를 따라가서 실제 다운로드 URL 찾기
            with httpx.Client(follow_redirects=True, timeout=30.0) as client:
                try:
                    # 먼저 링크에 접근하여 리다이렉트 확인
                    response = client.get(share_link)
                    final_url = str(response.url)
                    print(f"SharePoint redirect URL: {final_url}")
                    
                    # SharePoint 직접 다운로드 링크 생성
                    # 형식: https://*.sharepoint.com/:x:/s/... -> 다운로드 가능
                    # 또는 download.aspx 링크로 변환
                    if ":x:" in final_url or ":x:" in share_link:
                        # :x: 링크는 직접 다운로드 가능
                        # 다운로드 파라미터 추가
                        if "?download=1" not in final_url and "download.aspx" not in final_url:
                            # 직접 다운로드 링크 생성
                            if "?" in final_url:
                                download_url = final_url + "&download=1"
                            else:
                                download_url = final_url + "?download=1"
                        else:
                            download_url = final_url
                    else:
                        download_url = final_url
                except Exception as e:
                    print(f"Error following SharePoint redirect: {e}")
                    # 리다이렉트 실패 시 원본 링크 사용
                    download_url = share_link.replace(":f:", ":x:")
        
        # OneDrive 1drv.ms 링크 처리
        elif "1drv.ms" in share_link:
            print(f"Detected OneDrive 1drv.ms link: {share_link}")
            # 1drv.ms 링크는 리다이렉트를 따라가야 함
            with httpx.Client(follow_redirects=True, timeout=30.0) as client:
                response = client.get(share_link)
                # 리다이렉트된 URL에서 다운로드 링크 추출
                final_url = str(response.url)
                print(f"OneDrive redirect URL: {final_url}")
                # onedrive.live.com 링크로 변환
                if "onedrive.live.com" in final_url:
                    # 다운로드 링크로 변환 (embed를 download로 변경)
                    download_url = final_url.replace("/embed?", "/download?")
                else:
                    download_url = final_url
        
        # OneDrive live.com 링크 처리
        elif "onedrive.live.com" in share_link:
            print(f"Detected OneDrive live.com link: {share_link}")
            # 이미 onedrive.live.com 링크인 경우
            download_url = share_link.replace("/embed?", "/download?")
            if "download" not in download_url:
                # embed가 없으면 download 추가
                download_url = share_link.replace("?", "/download?")
        
        # Google Sheets 링크 처리 (docs.google.com/spreadsheets)
        elif "docs.google.com/spreadsheets" in share_link:
            print(f"Detected Google Sheets link: {share_link}")
            # Google Sheets 공유 링크에서 파일 ID 추출
            # 형식: https://docs.google.com/spreadsheets/d/FILE_ID/edit?usp=sharing
            file_id = None
            if "/spreadsheets/d/" in share_link:
                parts = share_link.split("/spreadsheets/d/")
                if len(parts) > 1:
                    file_id = parts[1].split("/")[0].split("?")[0]
            
            if file_id:
                # Google Sheets를 Excel 형식으로 다운로드
                download_url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
                print(f"Converted Google Sheets link to Excel download URL: {download_url}")
            else:
                print(f"Warning: Could not extract file ID from Google Sheets link: {share_link}")
                download_url = share_link
        
        # Google Drive 링크 처리
        elif "drive.google.com" in share_link:
            print(f"Detected Google Drive link: {share_link}")
            # Google Drive 공유 링크에서 파일 ID 추출
            # 형식: https://drive.google.com/file/d/FILE_ID/view?usp=sharing
            # 또는: https://drive.google.com/open?id=FILE_ID
            file_id = None
            if "/file/d/" in share_link:
                # /file/d/FILE_ID/ 형식
                parts = share_link.split("/file/d/")
                if len(parts) > 1:
                    file_id = parts[1].split("/")[0].split("?")[0]
            elif "id=" in share_link:
                # ?id=FILE_ID 형식
                from urllib.parse import urlparse, parse_qs
                parsed = urlparse(share_link)
                params = parse_qs(parsed.query)
                if "id" in params:
                    file_id = params["id"][0]
            
            if file_id:
                # Google Drive 직접 다운로드 링크 생성
                # 큰 파일의 경우 confirm 파라미터가 필요할 수 있음
                download_url = f"https://drive.google.com/uc?export=download&id={file_id}&confirm=t"
                print(f"Converted Google Drive link to download URL: {download_url}")
            else:
                print(f"Warning: Could not extract file ID from Google Drive link: {share_link}")
                # 파일 ID를 찾을 수 없으면 원본 링크 사용 시도
                download_url = share_link
        
        else:
            # 다른 형식의 링크는 그대로 사용
            print(f"Using link as-is: {share_link}")
            download_url = share_link
        
        if not download_url:
            print("Error: Could not determine download URL")
            return False
        
        print(f"Attempting to download from: {download_url}")
        
        # 파일 다운로드 (Google Drive 바이러스 스캔 페이지 처리 포함)
        with httpx.Client(follow_redirects=True, timeout=60.0) as client:
            response = client.get(download_url, follow_redirects=True)
            
            # Google Drive 바이러스 스캔 경고 페이지 처리
            if "drive.google.com" in download_url and response.status_code == 200:
                content_type = response.headers.get("content-type", "").lower()
                # HTML 응답이면 바이러스 스캔 페이지일 수 있음
                if "text/html" in content_type:
                    print("Detected HTML response (possibly virus scan warning), extracting download link...")
                    # HTML에서 실제 다운로드 링크 추출 시도
                    import re
                    html_content = response.text
                    # "downloadUrl" 또는 "uc-download-link" 찾기
                    download_pattern = r'href="(/uc\?export=download[^"]+)"'
                    matches = re.findall(download_pattern, html_content)
                    if matches:
                        actual_download_url = "https://drive.google.com" + matches[0]
                        print(f"Found actual download URL in HTML: {actual_download_url}")
                        # confirm 파라미터 추가
                        if "confirm=" not in actual_download_url:
                            actual_download_url += "&confirm=t"
                        response = client.get(actual_download_url, follow_redirects=True)
            
            if response.status_code == 200:
                output_path.parent.mkdir(parents=True, exist_ok=True)
                
                # 파일 유효성 검사: Excel 파일은 ZIP 형식이어야 함
                content = response.content
                # ZIP 파일 시그니처 확인 (PK\x03\x04)
                is_zip = content[:2] == b'PK'
                
                if not is_zip:
                    # HTML 에러 페이지일 가능성
                    content_text = content[:500].decode('utf-8', errors='ignore')
                    print(f"ERROR: Downloaded file is not a valid Excel file (ZIP signature not found)")
                    print(f"File content preview (first 500 chars): {content_text}")
                    return False
                
                with open(output_path, "wb") as f:
                    f.write(content)
                
                # 파일 크기 확인
                file_size = output_path.stat().st_size
                print(f"Successfully downloaded file to {output_path} (size: {file_size} bytes)")
                
                if file_size < 1000:  # 1KB 미만이면 의심스러움
                    print(f"WARNING: Downloaded file is very small ({file_size} bytes), might be an error page")
                    # 파일 내용 확인
                    with open(output_path, "rb") as f:
                        preview = f.read(100).decode('utf-8', errors='ignore')
                        if "<html" in preview.lower() or "<!doctype" in preview.lower():
                            print("ERROR: Downloaded file appears to be HTML, not Excel")
                            return False
                
                return True
            else:
                print(f"Failed to download file. Status code: {response.status_code}")
                print(f"Response headers: {dict(response.headers)}")
                # 응답 본문 일부 읽기 (에러 메시지 확인용)
                try:
                    error_text = response.text[:500] if hasattr(response, 'text') else "N/A"
                    print(f"Error response: {error_text}")
                except:
                    pass
                return False
    except Exception as e:
        print(f"Error downloading from OneDrive/SharePoint: {e}")
        print(traceback.format_exc())
        return False


def sync_onedrive_file(
    share_link: str,
    local_path: Path,
    sync_interval: int = 3600,
    force_download: bool = False
) -> bool:
    """OneDrive 파일을 주기적으로 동기화합니다.
    
    Args:
        share_link: OneDrive 공유 링크
        local_path: 로컬 저장 경로
        sync_interval: 동기화 주기 (초 단위, 기본 1시간)
        force_download: 강제 다운로드 여부
    
    Returns:
        동기화 성공 여부
    """
    # 파일이 존재하고 최근에 다운로드되었으면 스킵
    if not force_download and local_path.exists():
        file_age = time.time() - local_path.stat().st_mtime
        if file_age < sync_interval:
            print(f"File is recent (age: {file_age:.0f}s), skipping download")
            return True
    
    # 파일 다운로드
    return download_from_onedrive_share_link(share_link, local_path)


if __name__ == "__main__":
    # 테스트용
    share_link = os.getenv("ONEDRIVE_FILE_URL", "")
    if share_link:
        local_path = Path("★26SS MLB 생산스케쥴_DASHBOARD.xlsx")
        sync_onedrive_file(share_link, local_path, force_download=True)
    else:
        print("ONEDRIVE_FILE_URL environment variable not set")

