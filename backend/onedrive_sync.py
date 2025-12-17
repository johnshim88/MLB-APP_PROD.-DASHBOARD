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
                    response = client.get(share_link, allow_redirects=True)
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
        
        else:
            # 다른 형식의 링크는 그대로 사용
            print(f"Using link as-is: {share_link}")
            download_url = share_link
        
        if not download_url:
            print("Error: Could not determine download URL")
            return False
        
        print(f"Attempting to download from: {download_url}")
        
        # 파일 다운로드
        with httpx.Client(follow_redirects=True, timeout=60.0) as client:
            with client.stream("GET", download_url) as response:
                if response.status_code == 200:
                    output_path.parent.mkdir(parents=True, exist_ok=True)
                    with open(output_path, "wb") as f:
                        for chunk in response.iter_bytes():
                            f.write(chunk)
                    print(f"Successfully downloaded file to {output_path}")
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

