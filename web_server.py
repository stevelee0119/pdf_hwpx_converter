import os
import uuid
import shutil
import asyncio
from fastapi import FastAPI, UploadFile, File, Form, BackgroundTasks, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from core_converter import UniversalConverter, HwpAutomationEngine

app = FastAPI(title="Google Docs & PDF to HWPX Converter API")

# 전역 작업 상태 관리 저장소
# task_id -> {status, progress, message, result_file, estimated_seconds}
tasks_status = {}

# 임시 저장소 폴더 생성
UPLOAD_DIR = os.path.abspath("temp_uploads")
OUTPUT_DIR = os.path.abspath("converted_outputs")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 정적 파일 제공 설정 (HTML/CSS/JS)
# static 폴더가 실제 존재하는지 확인 후 마운트
os.makedirs("static", exist_ok=True)
app.mount("/static", StaticFiles(directory="static"), name="static")


@app.get("/", response_class=HTMLResponse)
def read_root():
    """메인 웹 페이지 반환"""
    index_path = os.path.join("static", "index.html")
    if os.path.exists(index_path):
        with open(index_path, "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())
    return HTMLResponse(content="<h1>HwpxConverter Server is Running</h1><p>static/index.html 파일을 찾을 수 없습니다.</p>")


def run_conversion_task(task_id: str, input_path: str, output_path: str, use_automation: bool, translate_to_ko: bool, is_temp_input: bool):
    """
    백그라운드 스레드 또는 비동기 태스크에서 호출되는 실제 문서 변환 함수입니다.
    """
    tasks_status[task_id]["status"] = "processing"
    
    def progress_callback(status_msg, percent):
        tasks_status[task_id]["progress"] = percent
        tasks_status[task_id]["message"] = status_msg
        
        # 진행도에 따른 예상 시간 갱신 (예: 총 15초 예상 기준 역산)
        total_estimate = tasks_status[task_id]["estimated_seconds"]
        remaining = max(1, int(total_estimate * (1 - percent / 100)))
        tasks_status[task_id]["remaining_seconds"] = remaining

    try:
        # 실제 변환 처리 호출
        result_path = UniversalConverter.convert(
            input_source=input_path,
            output_path=output_path,
            use_automation=use_automation,
            translate_to_ko=translate_to_ko,
            progress_callback=progress_callback
        )
        
        tasks_status[task_id]["status"] = "completed"
        tasks_status[task_id]["progress"] = 100
        tasks_status[task_id]["message"] = "변환 완료!"
        tasks_status[task_id]["result_file"] = result_path
        tasks_status[task_id]["original_name"] = os.path.basename(result_path)
        tasks_status[task_id]["remaining_seconds"] = 0

    except Exception as e:
        tasks_status[task_id]["status"] = "failed"
        tasks_status[task_id]["message"] = f"변환 실패: {str(e)}"
        tasks_status[task_id]["remaining_seconds"] = 0
    finally:
        # 변환이 끝난 후 임시 입력 파일 삭제 (구글 Docs 다운로드본이나 로컬 업로드본)
        if is_temp_input and os.path.exists(input_path):
            try:
                os.remove(input_path)
            except Exception as ex:
                print(f"Warning: Failed to remove temp input {input_path}: {ex}")


@app.get("/api/engines")
def api_engines():
    """사용 가능한 변환 엔진의 상태를 반환합니다."""
    from core_converter import LibBasedEngine
    return {
        "hwp_automation": HwpAutomationEngine.is_available(),
        "pure_lib": LibBasedEngine.is_available()
    }


@app.post("/api/convert")
def api_convert(
    background_tasks: BackgroundTasks,
    url: str = Form(None),
    file: UploadFile = File(None),
    use_automation: str = Form("true"),
    translate_to_ko: str = Form("false")
):
    """
    구글 Docs URL 또는 로컬 파일을 업로드받아 비동기 백그라운드 변환 태스크를 생성합니다.
    """
    # 1. 입력 유효성 검사
    if not url and not file:
        raise HTTPException(status_code=400, detail="Google Docs / Drive URL 또는 로컬 변환 대상 파일(PDF/DOCX) 중 하나는 반드시 업로드해야 합니다.")

    task_id = str(uuid.uuid4())
    is_automation = use_automation.lower() == "true"
    is_translation = translate_to_ko.lower() == "true"
    
    # 2. 로컬 한글 오피스 설치 여부 검사
    if is_automation and not HwpAutomationEngine.is_available():
        # 한글이 없으면 오픈소스 엔진으로 강제 우회 시도
        is_automation = False
        print(f"[Server] HwpAutomationEngine is not available. Falling back to Pure Library Engine.")

    # 3. 입력 소스 정의 및 임시 저장
    is_temp_input = False
    input_source = ""
    original_filename = "converted_document"
    
    if file:
        # 업로드된 파일을 로컬 임시 폴더에 저장
        file_ext = os.path.splitext(file.filename)[1].lower()
        if file_ext not in [".pdf", ".docx"]:
            raise HTTPException(status_code=400, detail="PDF 또는 DOCX 파일 형식만 지원합니다.")
            
        temp_input_path = os.path.join(UPLOAD_DIR, f"{task_id}{file_ext}")
        with open(temp_input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
            
        input_source = temp_input_path
        is_temp_input = True
        original_filename = os.path.splitext(file.filename)[0]
    else:
        # Google Docs URL 사용
        input_source = url
        # 구글 Docs ID에서 앞 8글자를 파일명으로 추출 시도
        doc_id = UniversalConverter.extract_document_id(url) if hasattr(UniversalConverter, 'extract_document_id') else None
        if not doc_id:
            # UniversalConverter에서 정의되지 않은 경우 백업 파싱
            from core_converter import GoogleDocsDownloader
            doc_id = GoogleDocsDownloader.extract_document_id(url)
            
        if doc_id:
            original_filename = f"gdocs_{doc_id[:8]}"

    # 4. 결과 파일 경로 설정
    fmt = "HWPX"

    out_file_name = f"{original_filename}.{fmt.lower()}"
    output_path = os.path.join(OUTPUT_DIR, f"{task_id}_{out_file_name}")

    # 5. 예상 시간 책정
    # PDF는 복잡하여 15초 정도, DOCX/구글Docs는 10초 정도로 기본 세팅
    ext = os.path.splitext(file.filename)[1].lower() if file else ".docx"
    estimated = 15 if ext == ".pdf" else 10
    
    # 6. 태스크 상태 초기화
    tasks_status[task_id] = {
        "status": "pending",
        "progress": 0,
        "message": "대기 중...",
        "result_file": None,
        "estimated_seconds": estimated,
        "remaining_seconds": estimated,
        "original_name": out_file_name
    }

    # 7. 백그라운드 태스크 추가
    background_tasks.add_task(
        run_conversion_task,
        task_id=task_id,
        input_path=input_source,
        output_path=output_path,
        use_automation=is_automation,
        translate_to_ko=is_translation,
        is_temp_input=is_temp_input
    )

    return {
        "task_id": task_id,
        "estimated_seconds": estimated
    }


@app.get("/api/status/{task_id}")
def api_status(task_id: str):
    """
    특정 태스크의 현재 진행 상태를 조회합니다.
    """
    if task_id not in tasks_status:
        raise HTTPException(status_code=404, detail="해당 작업을 찾을 수 없습니다.")
    return tasks_status[task_id]


@app.get("/api/download/{task_id}")
def api_download(task_id: str):
    """
    변환이 완료된 파일을 다운로드합니다.
    """
    if task_id not in tasks_status:
        raise HTTPException(status_code=404, detail="해당 작업을 찾을 수 없습니다.")
        
    task_info = tasks_status[task_id]
    if task_info["status"] != "completed" or not task_info["result_file"]:
        raise HTTPException(status_code=400, detail="변환이 아직 완료되지 않았거나 실패한 작업입니다.")
        
    file_path = task_info["result_file"]
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="결과 파일이 서버에서 삭제되었거나 찾을 수 없습니다.")

    return FileResponse(
        path=file_path,
        filename=task_info["original_name"],
        media_type="application/octet-stream"
    )


if __name__ == "__main__":
    import uvicorn
    # 로컬 네트워크에서도 접속 가능하도록 0.0.0.0 포트 8000에 서비스 구동
    uvicorn.run("web_server:app", host="0.0.0.0", port=8000, reload=True)
