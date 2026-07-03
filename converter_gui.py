import sys
import os
import traceback
from PySide6.QtCore import Qt, QThread, Signal, Slot, QSize
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QRadioButton, QButtonGroup,
    QProgressBar, QPlainTextEdit, QFileDialog, QFrame, QStackedWidget
)
from PySide6.QtGui import QFont, QColor, QPalette, QLinearGradient, QBrush, QIcon

from core_converter import UniversalConverter, HwpAutomationEngine, LibBasedEngine, GoogleDocsDownloader


class ConversionWorker(QThread):
    """
    GUI 프리징을 방지하기 위해 백그라운드에서 파일 변환을 진행하는 스레드입니다.
    """
    progress = Signal(str, int)  # (상태 메시지, 퍼센트)
    finished = Signal(bool, str)  # (성공 여부, 결과 메시지 또는 에러 메시지)

    def __init__(self, input_source, output_path, use_automation, translate_to_ko):
        super().__init__()
        self.input_source = input_source
        self.output_path = output_path
        self.use_automation = use_automation
        self.translate_to_ko = translate_to_ko

    def run(self):
        try:
            def progress_callback(status, percent):
                self.progress.emit(status, percent)

            result_path = UniversalConverter.convert(
                input_source=self.input_source,
                output_path=self.output_path,
                use_automation=self.use_automation,
                translate_to_ko=self.translate_to_ko,
                progress_callback=progress_callback
            )
            self.finished.emit(True, f"변환이 완료되었습니다!\n파일 경로: {result_path}")
        except Exception as e:
            err_msg = str(e)
            # 상세 트레이스백은 콘솔 출력
            traceback.print_exc()
            self.finished.emit(False, f"에러 발생: {err_msg}")


class DropArea(QFrame):
    """
    PDF 또는 DOCX 파일을 드래그앤드롭으로 입력받는 영역 컴포넌트입니다.
    """
    file_dropped = Signal(str)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.init_ui()

    def init_ui(self):
        self.setObjectName("DropArea")
        self.setFrameStyle(QFrame.StyledPanel | QFrame.Sunken)
        self.setMinimumHeight(150)
        
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)

        self.icon_label = QLabel("📥")
        self.icon_label.setFont(QFont("Segoe UI", 36))
        self.icon_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.icon_label)

        self.text_label = QLabel("여기에 PDF 또는 DOCX 파일을 드래그 앤 드롭하세요\n(또는 클릭하여 파일 선택)")
        self.text_label.setFont(QFont("Malgun Gothic", 11))
        self.text_label.setAlignment(Qt.AlignCenter)
        self.text_label.setStyleSheet("color: #a0a5c0; line-height: 1.5;")
        layout.addWidget(self.text_label)

        # 스타일링 적용 (보더 및 라운드 코너)
        self.setStyleSheet("""
            #DropArea {
                border: 2px dashed #4f5372;
                border-radius: 12px;
                background-color: #242533;
            }
            #DropArea:hover {
                border: 2px dashed #00ADB5;
                background-color: #2c2d3f;
            }
        """)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet("""
                #DropArea {
                    border: 2px dashed #00ADB5;
                    background-color: #2c2d3f;
                }
            """)

    def dragLeaveEvent(self, event):
        self.setStyleSheet("""
            #DropArea {
                border: 2px dashed #4f5372;
                background-color: #242533;
            }
        """)

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if os.path.exists(file_path):
                ext = os.path.splitext(file_path)[1].lower()
                if ext in ['.pdf', '.docx']:
                    self.file_dropped.emit(file_path)
                    break
        self.setStyleSheet("""
            #DropArea {
                border: 2px dashed #4f5372;
                background-color: #242533;
            }
        """)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            # 클릭 시에도 파일 탐색기 열기 트리거를 유도할 수 있도록 시그널 대신 빈 스트링 전달
            self.file_dropped.emit("")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.selected_file_path = ""
        self.worker = None
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Google Docs & PDF to HWPX Converter")
        self.resize(700, 750)
        self.setMinimumSize(600, 650)

        # 메인 윈도우 스타일 및 배경 설정 (다크 모드 프리미엄 테마)
        self.setStyleSheet("""
            QMainWindow {
                background-color: #1a1b26;
            }
            QWidget {
                color: #f8f8f2;
                font-family: "Malgun Gothic", "Segoe UI", sans-serif;
            }
            QLabel {
                font-size: 10pt;
            }
            QLineEdit {
                background-color: #242533;
                border: 1px solid #3c3f58;
                border-radius: 6px;
                padding: 8px 12px;
                color: #ffffff;
                selection-background-color: #00ADB5;
            }
            QLineEdit:focus {
                border: 1px solid #00ADB5;
            }
            QPushButton {
                background-color: #3e4260;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: bold;
                min-height: 20px;
            }
            QPushButton:hover {
                background-color: #4f5372;
            }
            QPushButton:pressed {
                background-color: #2f3248;
            }
            QRadioButton {
                spacing: 8px;
            }
            QRadioButton::indicator {
                width: 18px;
                height: 18px;
                border-radius: 9px;
            }
            QRadioButton::indicator::unchecked {
                border: 2px solid #5a5f80;
                background: none;
            }
            QRadioButton::indicator::checked {
                border: 2px solid #00ADB5;
                background-color: #00ADB5;
            }
            QProgressBar {
                border: 1px solid #3c3f58;
                border-radius: 6px;
                text-align: center;
                background-color: #242533;
            }
            QProgressBar::chunk {
                background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 #00ADB5, stop:1 #393E46);
                border-radius: 5px;
            }
            QPlainTextEdit {
                background-color: #12131a;
                border: 1px solid #242533;
                border-radius: 8px;
                color: #a9b1d6;
                font-family: "Consolas", "Courier New", monospace;
                font-size: 9.5pt;
            }
        """)

        # 중앙 위젯 및 기본 레이아웃
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(25, 25, 25, 25)
        main_layout.setSpacing(20)

        # 1. Header Banner
        header_layout = QVBoxLayout()
        header_title = QLabel("HWPX / HWP 문서 변환기")
        header_title.setFont(QFont("Malgun Gothic", 20, QFont.Bold))
        header_title.setStyleSheet("color: #00ADB5;")
        header_subtitle = QLabel("Google Docs 주소 또는 PDF 파일을 한글 문서로 손쉽게 변환해 보세요.")
        header_subtitle.setFont(QFont("Malgun Gothic", 10))
        header_subtitle.setStyleSheet("color: #787a91;")
        
        header_layout.addWidget(header_title)
        header_layout.addWidget(header_subtitle)
        main_layout.addLayout(header_layout)

        # 2. Input Type Selector Buttons (Tab 역할)
        tab_layout = QHBoxLayout()
        self.btn_gdocs_tab = QPushButton("🔗 Google Docs / Drive 링크 변환")
        self.btn_local_tab = QPushButton("📂 로컬 파일 변환 (PDF / DOCX)")
        
        self.btn_gdocs_tab.setCheckable(True)
        self.btn_local_tab.setCheckable(True)
        self.btn_gdocs_tab.setChecked(True)
        
        # 버튼 그룹 지정하여 배타적 선택
        self.tab_group = QButtonGroup(self)
        self.tab_group.addButton(self.btn_gdocs_tab, 0)
        self.tab_group.addButton(self.btn_local_tab, 1)
        self.tab_group.buttonToggled.connect(self.switch_tab)

        # 탭 선택 스타일링 지정
        self.update_tab_styles()
        
        tab_layout.addWidget(self.btn_gdocs_tab)
        tab_layout.addWidget(self.btn_local_tab)
        main_layout.addLayout(tab_layout)

        # 3. Stacked Content Widget
        self.stacked_widget = QStackedWidget()
        
        # 3a. Tab 1: Google Docs Input Screen
        gdocs_widget = QWidget()
        gdocs_layout = QVBoxLayout(gdocs_widget)
        gdocs_layout.setContentsMargins(0, 5, 0, 5)
        
        url_label = QLabel("Google Docs 또는 Drive URL 입력:")
        url_label.setStyleSheet("font-weight: bold; color: #a9b1d6;")
        self.url_input = QLineEdit()
        self.url_input.setPlaceholderText("https://docs.google.com/... 또는 https://drive.google.com/... 입력하세요.")
        
        info_label = QLabel("💡 링크 공유가 '링크가 있는 모든 사용자에게 뷰어' 이상으로 공개된 문서여야 접근 가능합니다.")
        info_label.setStyleSheet("color: #ff9e64; font-size: 9pt;")
        
        gdocs_layout.addWidget(url_label)
        gdocs_layout.addWidget(self.url_input)
        gdocs_layout.addWidget(info_label)
        gdocs_layout.addStretch()
        self.stacked_widget.addWidget(gdocs_widget)

        # 3b. Tab 2: Local File Input Screen
        local_widget = QWidget()
        local_layout = QVBoxLayout(local_widget)
        local_layout.setContentsMargins(0, 5, 0, 5)
        
        self.drop_area = DropArea()
        self.drop_area.file_dropped.connect(self.handle_file_input)
        
        self.lbl_selected_file = QLabel("선택된 파일: 없음")
        self.lbl_selected_file.setStyleSheet("color: #a9b1d6; font-weight: bold;")
        self.lbl_selected_file.setWordWrap(True)
        
        local_layout.addWidget(self.lbl_selected_file)
        local_layout.addWidget(self.drop_area)
        self.stacked_widget.addWidget(local_widget)

        main_layout.addWidget(self.stacked_widget)

        # Separator Line
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setFrameShadow(QFrame.Sunken)
        sep.setStyleSheet("background-color: #242533;")
        main_layout.addWidget(sep)

        # 4. Conversion Options Panel
        options_layout = QVBoxLayout()
        options_title = QLabel("변환 옵션 설정")
        options_title.setFont(QFont("Malgun Gothic", 11, QFont.Bold))
        options_title.setStyleSheet("color: #00ADB5;")
        options_layout.addWidget(options_title)

        # 4a. Format Option removed (HWPX default)

        # 4a-1. Translation Option
        trans_layout = QHBoxLayout()
        trans_lbl = QLabel("텍스트 번역:  ")
        trans_lbl.setStyleSheet("font-weight: bold;")
        self.rad_trans_none = QRadioButton("원본 유지 (그대로 추출)")
        self.rad_trans_ko = QRadioButton("영어 -> 한국어로 번역 후 추출")
        self.rad_trans_none.setChecked(True)
        
        self.trans_group = QButtonGroup(self)
        self.trans_group.addButton(self.rad_trans_none)
        self.trans_group.addButton(self.rad_trans_ko)

        trans_layout.addWidget(trans_lbl)
        trans_layout.addWidget(self.rad_trans_none)
        trans_layout.addWidget(self.rad_trans_ko)
        trans_layout.addStretch()
        options_layout.addLayout(trans_layout)
        # 4b. Engine Option
        engine_layout = QHBoxLayout()
        engine_lbl = QLabel("변환 엔진:  ")
        engine_lbl.setStyleSheet("font-weight: bold;")
        self.rad_engine_auto = QRadioButton("한글 연동 (Windows 전용, 서식 완벽 보존)")
        self.rad_engine_lib = QRadioButton("오픈소스 라이브러리 (한글 불필요, 텍스트 위주)")
        
        # 시스템 사양에 맞게 초기 엔진 제안
        if HwpAutomationEngine.is_available():
            self.rad_engine_auto.setChecked(True)
        else:
            self.rad_engine_lib.setChecked(True)
            self.rad_engine_auto.setEnabled(False)
            self.rad_engine_auto.setText("한글 연동 (미지원: 한글 미설치 또는 윈도우가 아님)")

        self.engine_group = QButtonGroup(self)
        self.engine_group.addButton(self.rad_engine_auto)
        self.engine_group.addButton(self.rad_engine_lib)

        engine_layout.addWidget(engine_lbl)
        engine_layout.addWidget(self.rad_engine_auto)
        engine_layout.addWidget(self.rad_engine_lib)
        engine_layout.addStretch()
        options_layout.addLayout(engine_layout)

        # 4c. Save Directory Path
        save_dir_layout = QHBoxLayout()
        save_lbl = QLabel("저장 경로:  ")
        save_lbl.setStyleSheet("font-weight: bold;")
        self.txt_save_path = QLineEdit()
        self.txt_save_path.setPlaceholderText("기본 다운로드 또는 입력 파일 위치에 저장됩니다.")
        self.btn_browse = QPushButton("폴더 선택")
        self.btn_browse.clicked.connect(self.browse_save_directory)
        
        save_dir_layout.addWidget(save_lbl)
        save_dir_layout.addWidget(self.txt_save_path)
        save_dir_layout.addWidget(self.btn_browse)
        options_layout.addLayout(save_dir_layout)

        main_layout.addLayout(options_layout)

        # 5. Conversion Start Button & Progress Bar
        self.btn_convert = QPushButton("🚀 문서 변환 시작")
        self.btn_convert.setFont(QFont("Malgun Gothic", 12, QFont.Bold))
        # 그라데이션 및 현대적인 버튼 호버 스타일 직접 오버라이딩
        self.btn_convert.setStyleSheet("""
            QPushButton {
                background: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 #00ADB5, stop:1 #393E46);
                color: #ffffff;
                border: none;
                border-radius: 8px;
                padding: 12px;
                min-height: 25px;
            }
            QPushButton:hover {
                background: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 #00f0ff, stop:1 #4f5366);
            }
            QPushButton:disabled {
                background: #2b2c3a;
                color: #62647a;
            }
        """)
        self.btn_convert.clicked.connect(self.start_conversion)
        main_layout.addWidget(self.btn_convert)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)

        self.lbl_status = QLabel("")
        self.lbl_status.setStyleSheet("color: #00ADB5; font-weight: bold;")
        self.lbl_status.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.lbl_status)

        # 6. Log Console
        log_layout = QVBoxLayout()
        log_lbl = QLabel("실행 로그")
        log_lbl.setStyleSheet("font-weight: bold; color: #a9b1d6;")
        self.txt_log = QPlainTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setMaximumHeight(150)
        
        log_layout.addWidget(log_lbl)
        log_layout.addWidget(self.txt_log)
        main_layout.addLayout(log_layout)

        if HwpAutomationEngine.is_available():
            self.append_log("시스템 감지: 로컬에 설치된 아래아한글 한컴오피스 엔진을 사용 가능합니다.")
        else:
            self.append_log("시스템 감지: 아래아한글 OLE API를 감지할 수 없습니다. 오픈소스 엔진만 선택할 수 있습니다.")

        # Footer Info
        footer = QLabel("© 2026 Antigravity IDE. All Rights Reserved.")
        footer.setFont(QFont("Segoe UI", 8))
        footer.setStyleSheet("color: #4f5372;")
        footer.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(footer)

    def update_tab_styles(self):
        active_style = """
            QPushButton {
                background-color: #00ADB5;
                color: #1a1b26;
                font-weight: bold;
                border-radius: 6px;
                padding: 10px;
            }
        """
        inactive_style = """
            QPushButton {
                background-color: #242533;
                color: #a0a5c0;
                font-weight: normal;
                border-radius: 6px;
                padding: 10px;
                border: 1px solid #3c3f58;
            }
            QPushButton:hover {
                background-color: #2c2d3f;
                color: #ffffff;
            }
        """
        self.btn_gdocs_tab.setStyleSheet(active_style if self.btn_gdocs_tab.isChecked() else inactive_style)
        self.btn_local_tab.setStyleSheet(active_style if self.btn_local_tab.isChecked() else inactive_style)

    @Slot(bool)
    def switch_tab(self, checked):
        if not checked:
            return
        active_id = self.tab_group.checkedId()
        self.stacked_widget.setCurrentIndex(active_id)
        self.update_tab_styles()
        
        if active_id == 0:
            self.append_log("모드 변경: 구글 Docs 주소 변환")
        else:
            self.append_log("모드 변경: 로컬 파일 변환 (PDF / DOCX)")

    def handle_file_input(self, file_path):
        if file_path == "":
            # 파일 선택 창 열기
            file_name, _ = QFileDialog.getOpenFileName(
                self, "변환할 파일 선택", "", "문서 파일 (*.pdf *.docx)"
            )
            if file_name:
                self.selected_file_path = file_name
                self.lbl_selected_file.setText(f"선택된 파일: {os.path.basename(file_name)}")
                self.append_log(f"파일이 로드되었습니다: {file_name}")
                # 기본 저장 경로 제안
                self.txt_save_path.setText(os.path.dirname(file_name))
        else:
            self.selected_file_path = file_path
            self.lbl_selected_file.setText(f"선택된 파일: {os.path.basename(file_path)}")
            self.append_log(f"파일이 드롭되었습니다: {file_path}")
            # 기본 저장 경로 제안
            self.txt_save_path.setText(os.path.dirname(file_path))

    def browse_save_directory(self):
        dir_path = QFileDialog.getExistingDirectory(self, "저장 폴더 선택")
        if dir_path:
            self.txt_save_path.setText(dir_path)
            self.append_log(f"저장 폴더가 변경되었습니다: {dir_path}")

    def append_log(self, text):
        self.txt_log.appendPlainText(text)

    def start_conversion(self):
        # 탭 확인
        active_tab = self.tab_group.checkedId()
        
        input_source = ""
        default_out_name = "converted_document"
        
        if active_tab == 0:
            # Google Docs URL 모드
            url = self.url_input.text().strip()
            if not url:
                self.append_log("경고: Google Docs URL을 입력해 주세요.")
                return
            input_source = url
            # 구글 Docs는 url에서 문서명 매핑하기가 곤란하므로 기본 변환 문서명 사용
            doc_id = GoogleDocsDownloader.extract_document_id(url)
            if doc_id:
                default_out_name = f"gdocs_{doc_id[:8]}"
        else:
            # 로컬 파일 모드
            if not self.selected_file_path:
                self.append_log("경고: 변환할 PDF 또는 DOCX 파일을 선택해 주세요.")
                return
            input_source = self.selected_file_path
            default_out_name = os.path.splitext(os.path.basename(input_source))[0]

        # 저장 폴더 정의
        save_dir = self.txt_save_path.text().strip()
        if not save_dir:
            # 입력 파일 경로의 디렉토리 혹은 다운로드 경로
            if active_tab == 1 and self.selected_file_path:
                save_dir = os.path.dirname(self.selected_file_path)
            else:
                save_dir = os.path.join(os.path.expanduser("~"), "Downloads")
                if not os.path.exists(save_dir):
                    save_dir = os.getcwd()

        # 출력 포맷 (HWPX 고정)
        fmt = "HWPX"
        out_file_name = f"{default_out_name}.{fmt.lower()}"
        output_path = os.path.join(save_dir, out_file_name)

        # 엔진 선택
        use_auto = self.rad_engine_auto.isChecked()
        
        # 번역 선택
        translate_to_ko = self.rad_trans_ko.isChecked()

        # UI 비활성화
        self.btn_convert.setEnabled(False)
        self.btn_gdocs_tab.setEnabled(False)
        self.btn_local_tab.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.lbl_status.setText("변환 작업 시작 중...")

        self.append_log(f"--- 변환 작업 시작 ---")
        self.append_log(f"입력 소스: {input_source}")
        self.append_log(f"출력 경로: {output_path}")
        self.append_log(f"사용 엔진: {'한글 API 연동' if use_auto else '오픈소스 라이브러리'}")
        self.append_log(f"번역 옵션: {'적용' if translate_to_ko else '미적용'}")
        self.append_log(f"출력 형식: {fmt}")

        # 백그라운드 스레드 시작
        self.worker = ConversionWorker(
            input_source=input_source,
            output_path=output_path,
            use_automation=use_auto,
            translate_to_ko=translate_to_ko
        )
        self.worker.progress.connect(self.on_conversion_progress)
        self.worker.finished.connect(self.on_conversion_finished)
        self.worker.start()

    @Slot(str, int)
    def on_conversion_progress(self, status_text, percent):
        self.lbl_status.setText(status_text)
        self.progress_bar.setValue(percent)
        self.append_log(f"[{percent}%] {status_text}")

    @Slot(bool, str)
    def on_conversion_finished(self, success, message):
        self.btn_convert.setEnabled(True)
        self.btn_gdocs_tab.setEnabled(True)
        self.btn_local_tab.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.lbl_status.setText("")

        if success:
            self.lbl_status.setStyleSheet("color: #50fa7b; font-weight: bold;")
            self.lbl_status.setText("변환 완료!")
            self.append_log(f"[완료] {message}")
        else:
            self.lbl_status.setStyleSheet("color: #ff5555; font-weight: bold;")
            self.lbl_status.setText("변환 실패")
            self.append_log(f"[실패] {message}")

        self.append_log(f"----------------------\n")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
