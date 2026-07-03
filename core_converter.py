import os
import re
import requests
import traceback

# Optional imports for Libre Engine
try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    from hwpx import HwpxDocument
except ImportError:
    HwpxDocument = None

try:
    from docx import Document as DocxDocument
except ImportError:
    DocxDocument = None

try:
    from pdf2docx import Converter
except ImportError:
    Converter = None

try:
    from deep_translator import GoogleTranslator
except ImportError:
    GoogleTranslator = None

# Optional import for win32com
try:
    import win32com.client as win32
except ImportError:
    win32 = None

# Downloader for Google Docs (DOCX) URLs
class GoogleDocsDownloader:
    """
    Google Docs URL에서 문서를 파싱하고 DOCX 형태로 로컬에 다운로드하는 클래스입니다.
    """
    @staticmethod
    def extract_document_id(url):
        # 구글 Docs URL에서 Document ID 추출하는 정규식
        # 예: https://docs.google.com/document/d/1Xxxxx_xxxx/edit
        match = re.search(r"/document/d/([a-zA-Z0-9-_]+)", url)
        if match:
            return match.group(1)
        return None

    @classmethod
    def download_as_docx(cls, url, output_path):
        doc_id = cls.extract_document_id(url)
        if not doc_id:
            raise ValueError("올바른 Google Docs URL이 아닙니다. '/document/d/문서ID' 형식이 포함되어야 합니다.")
        export_url = f"https://docs.google.com/document/d/{doc_id}/export?format=docx"
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"}
        response = requests.get(export_url, headers=headers, timeout=30)
        if response.status_code == 200:
            with open(output_path, 'wb') as f:
                f.write(response.content)
            return output_path
        elif response.status_code in (401, 403):
            raise PermissionError("구글 Docs 문서에 접근할 권한이 없습니다. 링크 공유 설정을 '링크가 있는 모든 사용자에게 공개'로 변경해 주세요.")
        else:
            raise Exception(f"Google Docs 다운로드 실패 (HTTP {response.status_code}). URL을 확인해 주세요.")

# Downloader for public Google Drive PDF files
class GoogleDriveDownloader:
    """Download a PDF from a public Google Drive share link."""
    @staticmethod
    def extract_file_id(url):
        # Pattern: https://drive.google.com/file/d/<ID>/view
        match = re.search(r"/d/([a-zA-Z0-9_-]+)", url)
        if match:
            return match.group(1)
        # Alternate pattern with id parameter
        match = re.search(r"id=([a-zA-Z0-9_-]+)", url)
        if match:
            return match.group(1)
        raise ValueError("구글 드라이브 URL에서 파일 ID를 추출할 수 없습니다.")

    @staticmethod
    def download_as_pdf(url, output_path):
        file_id = GoogleDriveDownloader.extract_file_id(url)
        export_url = f"https://drive.google.com/uc?export=download&id={file_id}"
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"}
        response = requests.get(export_url, headers=headers, stream=True, timeout=30)
        if response.status_code == 200:
            with open(output_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            return output_path
        else:
            raise Exception(f"Google Drive PDF 다운로드 실패 (HTTP {response.status_code}). URL을 확인해 주세요.")

class DocxTranslator:
    """
    DOCX 파일의 문단과 표를 순회하며 텍스트를 한국어로 번역하는 클래스입니다.
    """
    @staticmethod
    def translate_docx(docx_path, output_path, progress_callback=None):
        if GoogleTranslator is None:
            raise RuntimeError("deep-translator 라이브러리가 필요합니다. 'pip install deep-translator'를 진행해 주세요.")
        if DocxDocument is None:
            raise RuntimeError("python-docx 라이브러리가 필요합니다.")

        doc = DocxDocument(docx_path)
        translator = GoogleTranslator(source='auto', target='ko')

        # 문단 번역
        for p in doc.paragraphs:
            text = p.text.strip()
            if text:
                try:
                    translated = translator.translate(text)
                    if translated:
                        # 기존 run들을 지우고 번역된 텍스트를 새 run으로 추가
                        for r in p.runs:
                            r.text = ''
                        p.add_run(translated)
                except Exception as e:
                    print(f"[DocxTranslator] 문단 번역 실패: {e}")

        # 표 번역
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        text = p.text.strip()
                        if text:
                            try:
                                translated = translator.translate(text)
                                if translated:
                                    for r in p.runs:
                                        r.text = ''
                                    p.add_run(translated)
                            except Exception as e:
                                print(f"[DocxTranslator] 표 텍스트 번역 실패: {e}")
                                
        doc.save(output_path)


class HwpAutomationEngine:
    """
    Windows 환경에서 한컴오피스 한글 프로그램의 API(win32com)를 직접 제어하여 고품질 변환을 수행하는 클래스입니다.
    """
    @staticmethod
    def is_available():
        if win32 is None:
            return False
        try:
            # 한글 객체 생성 테스트
            hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            hwp.Quit()
            return True
        except Exception:
            return False

    @classmethod
    def convert_file(cls, input_path, output_path):
        """
        win32com을 활용해 입력을 HWPX로 변환합니다.
        지원 입력: PDF, DOCX
        """
        if not cls.is_available():
            raise RuntimeError("한컴오피스가 설치되어 있지 않거나, 윈도우 환경이 아닙니다. OLE 자동화를 사용할 수 없습니다.")

        input_abs = os.path.abspath(input_path)
        output_abs = os.path.abspath(output_path)
        ext = os.path.splitext(input_abs)[1].lower()
        format_type = "HWPX"

        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        try:
            # 보안 경고창 우회를 위한 모듈 등록 시도
            try:
                hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
            except Exception as e:
                print(f"[Automation] Security DLL register failed: {e}")

            opened = False
            if ext == ".pdf":
                # PDF 열기 시도
                try:
                    # 1순위: FileOpenPDF 액션
                    pset = hwp.HParameterSet.HFileOpenSave
                    hwp.HAction.GetDefault("FileOpenPDF", pset.HSet)
                    pset.filename = input_abs
                    result = hwp.HAction.Execute("FileOpenPDF", pset.HSet)
                    if result:
                        opened = True
                except Exception as e:
                    print(f"[Automation] FileOpenPDF Action failed: {e}, falling back to Open()")
                
                if not opened:
                    # 2순위: PDF 포맷 필터 인자를 이용한 일반 Open
                    opened = hwp.Open(input_abs, "PDF", "")
            elif ext == ".docx":
                # DOCX 열기
                opened = hwp.Open(input_abs, "docx", "")
            else:
                # 일반 문서 열기
                opened = hwp.Open(input_abs)

            if not opened:
                raise Exception(f"한글 프로그램에서 입력 파일({os.path.basename(input_path)})을 열지 못했습니다.")

            # HWPX 또는 HWP 형식으로 저장
            saved = False
            try:
                hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
                hwp.HParameterSet.HFileOpenSave.filename = output_abs
                hwp.HParameterSet.HFileOpenSave.Format = format_type # "HWPX" or "HWP"
                result = hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
                if result:
                    saved = True
            except Exception as e:
                print(f"[Automation] FileSaveAs_S action failed: {e}, falling back to SaveAs()")

            if not saved:
                # 예비 저장 메서드 호출
                saved = hwp.SaveAs(output_abs, format_type, "")

            if not saved:
                raise Exception("변환 파일 저장에 실패했습니다.")

        except Exception as e:
            print(f"[Automation] Error during conversion: {e}")
            traceback.print_exc()
            raise e
        finally:
            try:
                hwp.Quit()
            except Exception:
                pass


class LibBasedEngine:
    """
    한컴오피스가 없는 환경에서 오픈소스 파서를 활용해 텍스트 기반 HWPX를 직접 구축하는 클래스입니다.
    """
    @staticmethod
    def is_available():
        return HwpxDocument is not None

    @classmethod
    def _extract_page_text(cls, page):
        """
        pdfplumber를 이용해 페이지 텍스트를 추출할 때 공백(띄어쓰기) 누락을 방지하는 다중 알고리즘 구현
        """
        # 1차: x_tolerance 정밀 조정을 통한 기본 텍스트 추출 시도
        orig_text = page.extract_text(x_tolerance=1.5, y_tolerance=3) or ""

        # 2차: extract_words 기반 좌표 재구성 (단어 간 띄어쓰기 강제 보장)
        try:
            words = page.extract_words(x_tolerance=1.5, y_tolerance=3)
        except Exception:
            words = []

        if not words:
            return orig_text

        # y좌표(top) 기준 그룹화 (tolerance 3pt)
        lines_dict = {}
        for word in words:
            top_key = round(word['top'] / 3.0) * 3.0
            if top_key not in lines_dict:
                lines_dict[top_key] = []
            lines_dict[top_key].append(word)

        sorted_tops = sorted(lines_dict.keys())
        reconstructed_lines = []
        for top in sorted_tops:
            line_words = sorted(lines_dict[top], key=lambda w: w['x0'])
            line_text = " ".join([w['text'] for w in line_words])
            reconstructed_lines.append(line_text)

        reconstructed_text = "\n".join(reconstructed_lines)

        # 공백 개수 비교를 통해 띄어쓰기가 더 잘 유지된 결과물 선택
        if orig_text.count(" ") >= reconstructed_text.count(" "):
            return orig_text
        return reconstructed_text

    @classmethod
    def convert_pdf_to_hwpx(cls, pdf_path, output_path):
        if not cls.is_available():
            raise RuntimeError("python-hwpx 라이브러리가 로드되지 않았습니다.")
        if pdfplumber is None:
            raise RuntimeError("pdfplumber 라이브러리가 필요합니다. 'pip install pdfplumber'를 진행해 주세요.")

        doc = HwpxDocument.new()

        # Temporary directory for extracted images
        import tempfile, shutil
        temp_dir = tempfile.mkdtemp()
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for i, page in enumerate(pdf.pages):
                    # ---- Text extraction ----
                    text = cls._extract_page_text(page)
                    if text:
                        for line in text.split('\n'):
                            doc.add_paragraph(line.strip())
                    # ---- Image extraction ----
                    if page.images:
                        for img_index, img_dict in enumerate(page.images):
                            try:
                                # Render image using pdfplumber's to_image
                                pil_img = page.to_image(resolution=150).original
                                img_path = os.path.join(temp_dir, f"page_{i}_img_{img_index}.png")
                                pil_img.save(img_path)
                                doc.add_image(img_path)
                            except Exception as e:
                                print(f"[LibBasedEngine] 이미지 추출 실패: {e}")
                    # ---- Table extraction ----
                    try:
                        tables = page.extract_tables()
                        for table in tables:
                            if table:
                                doc.add_table(table)
                    except Exception as e:
                        print(f"[LibBasedEngine] 테이블 추출 실패: {e}")
                    # 페이지 구분을 위한 간단한 문단 추가
                    if i < len(pdf.pages) - 1:
                        doc.add_paragraph("--------------------------------------------------------------------------------")
            doc.save_to_path(output_path)
        finally:
            # Clean up temporary image files
            shutil.rmtree(temp_dir, ignore_errors=True)

    @classmethod
    def convert_docx_to_hwpx(cls, docx_path, output_path):
        if not cls.is_available():
            raise RuntimeError("python-hwpx 라이브러리가 로드되지 않았습니다.")
        if DocxDocument is None:
            raise RuntimeError("python-docx 라이브러리가 필요합니다. 'pip install python-docx'를 진행해 주세요.")

        doc = HwpxDocument.new()
        docx_doc = DocxDocument(docx_path)
        
        for para in docx_doc.paragraphs:
            # 빈 줄을 포함해 본문 문단 추가
            doc.add_paragraph(para.text)
            
        doc.save_to_path(output_path)


class UniversalConverter:
    """
    엔진 선택 및 입력을 조합하여 전체 변환 흐름을 중재하는 메인 컨트롤러 클래스입니다.
    """
    @classmethod
    def convert(cls, input_source, output_path, use_automation=True, translate_to_ko=False, progress_callback=None):
        """
        input_source: 구글 Docs URL, Google Drive URL 또는 로컬 파일 경로 (PDF, DOCX)
        output_path: 변환 결과가 저장될 HWPX 파일 경로
        use_automation: True 시 한글 OLE 자동화 사용, False 시 오픈소스 라이브러리 사용
        translate_to_ko: True 시 영어를 한국어로 자동 번역
        """
        temp_files = []
        try:
            # 1단계: 구글 Docs URL 확인 및 다운로드
            is_url = input_source.startswith("http://") or input_source.startswith("https://")
            local_input = input_source
            
            if is_url:
                if progress_callback:
                    progress_callback("URL에서 파일 다운로드 중...", 20)
                
                if "drive.google.com" in input_source:
                    temp_pdf = output_path + ".temp.pdf"
                    GoogleDriveDownloader.download_as_pdf(input_source, temp_pdf)
                    temp_files.append(temp_pdf)
                    local_input = temp_pdf
                else:
                    temp_docx = output_path + ".temp.docx"
                    GoogleDocsDownloader.download_as_docx(input_source, temp_docx)
                    temp_files.append(temp_docx)
                    local_input = temp_docx

            if progress_callback:
                progress_callback("파일 변환 준비 중...", 40)

            # 2단계: 변환 방식 선택 및 실행
            ext = os.path.splitext(local_input)[1].lower()
            
            # PDF 입력인 경우, 고품질 레이아웃 보존을 위해 먼저 DOCX로 변환
            if ext == ".pdf":
                if Converter is None:
                    raise RuntimeError("pdf2docx 라이브러리가 필요합니다. 'pip install pdf2docx'를 진행해 주세요.")
                if progress_callback:
                    progress_callback("PDF 레이아웃을 보존하기 위해 중간 변환 중...", 55)
                temp_pdf_docx = output_path + ".converted.docx"
                cv = Converter(local_input)
                cv.convert(temp_pdf_docx, start=0, end=None)
                cv.close()
                temp_files.append(temp_pdf_docx)
                local_input = temp_pdf_docx
                ext = ".docx"
            
            # 2.5단계: 번역 옵션 적용
            if translate_to_ko:
                if ext == ".docx":
                    if progress_callback:
                        progress_callback("문서 내용을 한국어로 번역 중입니다...", 65)
                    temp_translated = output_path + ".translated.docx"
                    DocxTranslator.translate_docx(local_input, temp_translated, progress_callback)
                    temp_files.append(temp_translated)
                    local_input = temp_translated
                else:
                    print("Warning: 현재 번역 기능은 DOCX 변환 과정에서만 지원됩니다.")

            if use_automation:
                if not HwpAutomationEngine.is_available():
                    raise RuntimeError("한컴오피스 자동화 엔진을 사용할 수 없는 환경입니다. (Windows OS & 한컴오피스 필수)")
                
                if progress_callback:
                    progress_callback("한컴오피스 연동 엔진으로 문서 변환 중...", 70)
                HwpAutomationEngine.convert_file(local_input, output_path)
            else:
                if progress_callback:
                    progress_callback("오픈소스 엔진으로 텍스트 파싱 및 HWPX 생성 중...", 70)
                
                if ext == ".docx":
                    LibBasedEngine.convert_docx_to_hwpx(local_input, output_path)
                else:
                    raise ValueError(f"지원하지 않는 입력 형식입니다: {ext}")

            if progress_callback:
                progress_callback(f"변환 완료! 결과 파일: {os.path.basename(output_path)}", 100)
                
            return output_path

        finally:
            # 임시 파일 정리
            for temp_file in temp_files:
                if os.path.exists(temp_file):
                    try:
                        os.remove(temp_file)
                    except Exception as e:
                        print(f"Warning: Failed to delete temp file {temp_file}: {e}")
