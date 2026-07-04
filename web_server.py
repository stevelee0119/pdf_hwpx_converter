import os
import uuid
import shutil
import asyncio
import sqlite3
from fastapi import FastAPI, UploadFile, File, Form, BackgroundTasks, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from core_converter import UniversalConverter, HwpAutomationEngine

app = FastAPI(title="Google Docs & PDF to HWPX Converter API")

# 임시 저장소 및 결과물 폴더 경로 설정
UPLOAD_DIR = os.path.abspath("temp_uploads")
OUTPUT_DIR = os.path.abspath("converted_outputs")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 정적 파일 제공 설정
os.makedirs("static", exist_ok=True)
app.mount("/static", StaticFiles(directory="static"), name="static")

# SQLite 데이터베이스 경로
DB_PATH = os.path.abspath("tasks.db")

def init_db():
    """다중 프로세스 환경에서도 상태를 동기화하기 위해 SQLite 데이터베이스 테이블을 생성합니다."""
    conn = sqlite3.connect(DB_PATH, timeout=20.0)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS tasks (
            task_id TEXT PRIMARY KEY,
            status TEXT,
            progress INTEGER,
            message TEXT,
            result_file TEXT,
            estimated_seconds INTEGER,
            remaining_seconds INTEGER,
            original_name TEXT
        )
    """)
    conn.commit()
    conn.close()

# 서버 구동 전 DB 초기화
init_db()

def get_task_status(task_id: str):
    """DB에서 특정 작업의 상태 정보를 조회하여 반환합니다."""
    conn = sqlite3.connect(DB_PATH, timeout=20.0)
    cursor = conn.cursor()
    cursor.execute("""
        SELECT status, progress, message, result_file, estimated_seconds, remaining_seconds, original_name 
        FROM tasks WHERE task_id = ?
    """, (task_id,))
    row = cursor.fetchone()
    conn.close()
    if not row:
        return None
    return {
        "status": row[0],
        "progress": row[1],
        "message": row[2],
        "result_file": row[3],
        "estimated_seconds": row[4],
        "remaining_seconds": row[5],
        "original_name": row[6]
    }

def set_task_status_pending(task_id: str, estimated_seconds: int, original_name: str):
    """새 작업을 대기 상태(pending)로 등록합니다."""
    conn = sqlite3.connect(DB_PATH, timeout=20.0)
    cursor = conn.cursor()
    cursor.execute("""
        INSERT OR REPLACE INTO tasks 
        (task_id, status, progress, message, result_file, estimated_seconds, remaining_seconds, original_name)
        VALUES (?, 'pending', 0, '대기 중...', NULL, ?, ?, ?)
    """, (task_id, estimated_seconds, estimated_seconds, original_name))
    conn.commit()
    conn.close()

def update_task_status_processing(task_id: str):
    """작업의 상태를 처리 중(processing)으로 갱신합니다."""
    conn = sqlite3.connect(DB_PATH, timeout=20.0)
    cursor = conn.cursor()
    cursor.execute("UPDATE tasks SET status = 'processing' WHERE task_id = ?", (task_id,))
    conn.commit()
    conn.close()

def update_task_progress(task_id: str, progress: int, message: str, remaining_seconds: int):
    """작업의 진행 상태와 메시지를 주기적으로 기록합니다."""
    conn = sqlite3.connect(DB_PATH, timeout=20.0)
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE tasks 
        SET progress = ?, message = ?, remaining_seconds = ? 
        WHERE task_id = ?
    """, (progress, message, remaining_seconds, task_id))
    conn.commit()
    conn.close()

def set_task_status_completed(task_id: str, result_file: str, original_name: str):
    """작업이 성공적으로 완수(completed)되었음을 갱신합니다."""
    conn = sqlite3.connect(DB_PATH, timeout=20.0)
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE tasks 
        SET status = 'completed', progress = 100, message = '변환 완료!', result_file = ?, original_name = ?, remaining_seconds = 0
        WHERE task_id = ?
    """, (result_file, original_name, task_id))
    conn.commit()
    conn.close()

def set_task_status_failed(task_id: str, error_msg: str):
    """작업 도중 오류가 발생(failed)했음을 갱신합니다."""
    conn = sqlite3.connect(DB_PATH, timeout=20.0)
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE tasks 
        SET status = 'failed', message = ?, remaining_seconds = 0
        WHERE task_id = ?
    """, (error_msg, task_id))
    conn.commit()
    conn.close()


@app.get("/", response_class=HTMLResponse)
@app.head("/", response_class=HTMLResponse)
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
    update_task_status_processing(task_id)
    
    # 초기에 설정된 작업 메타 정보 조회
    task_info = get_task_status(task_id)
    total_estimate = task_info["estimated_seconds"] if task_info else 15

    def progress_callback(status_msg, percent):
        remaining = max(1, int(total_estimate * (1 - percent / 100)))
        update_task_progress(task_id, percent, status_msg, remaining)

    try:
        # 실제 변환 처리 호출
        result_path = UniversalConverter.convert(
            input_source=input_path,
            output_path=output_path,
            use_automation=use_automation,
            translate_to_ko=translate_to_ko,
            progress_callback=progress_callback
        )
        
        set_task_status_completed(task_id, result_path, os.path.basename(result_path))

    except Exception as e:
        set_task_status_failed(task_id, f"변환 실패: {str(e)}")
    finally:
        # 변환이 끝난 후 임시 입력 파일 삭제
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
    if not url and not file:
        raise HTTPException(status_code=400, detail="Google Docs / Drive URL 또는 로컬 변환 대상 파일(PDF/DOCX) 중 하나는 반드시 업로드해야 합니다.")

    task_id = str(uuid.uuid4())
    is_automation = use_automation.lower() == "true"
    is_translation = translate_to_ko.lower() == "true"
    
    if is_automation and not HwpAutomationEngine.is_available():
        is_automation = False
        print(f"[Server] HwpAutomationEngine is not available. Falling back to Pure Library Engine.")

    is_temp_input = False
    input_source = ""
    original_filename = "converted_document"
    
    if file:
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
        input_source = url
        doc_id = UniversalConverter.extract_document_id(url) if hasattr(UniversalConverter, 'extract_document_id') else None
        if not doc_id:
            from core_converter import GoogleDocsDownloader
            doc_id = GoogleDocsDownloader.extract_document_id(url)
            
        if doc_id:
            original_filename = f"gdocs_{doc_id[:8]}"

    fmt = "HWPX"
    out_file_name = f"{original_filename}.{fmt.lower()}"
    output_path = os.path.join(OUTPUT_DIR, f"{task_id}_{out_file_name}")

    ext = os.path.splitext(file.filename)[1].lower() if file else ".docx"
    estimated = 15 if ext == ".pdf" else 10
    
    # 6. 태스크 상태 초기화 (SQLite DB에 저장)
    set_task_status_pending(task_id, estimated, out_file_name)

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
    task_info = get_task_status(task_id)
    if not task_info:
        raise HTTPException(status_code=404, detail="해당 작업을 찾을 수 없습니다.")
    return task_info


@app.get("/api/download/{task_id}")
def api_download(task_id: str):
    """
    변환이 완료된 파일을 다운로드합니다.
    """
    task_info = get_task_status(task_id)
    if not task_info:
        raise HTTPException(status_code=404, detail="해당 작업을 찾을 수 없습니다.")
        
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
    uvicorn.run("web_server:app", host="0.0.0.0", port=8000, reload=True)
