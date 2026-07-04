document.addEventListener("DOMContentLoaded", () => {
    // 1. DOM Elements
    const tabGdocs = document.getElementById("tab-gdocs");
    const tabLocal = document.getElementById("tab-local");
    const panelGdocs = document.getElementById("panel-gdocs");
    const panelLocal = document.getElementById("panel-local");

    const gdocsUrlInput = document.getElementById("gdocs-url");

    const dropZone = document.getElementById("drop-zone");
    const fileInput = document.getElementById("file-input");
    const fileBanner = document.getElementById("file-banner");
    const selectedFileName = document.getElementById("selected-file-name");
    const btnRemoveFile = document.getElementById("btn-remove-file");
    const btnSelectFile = document.getElementById("btn-select-file");

    const engineAuto = document.getElementById("engine-auto");

    const progressSection = document.getElementById("progress-section");
    const statusMessage = document.getElementById("status-message");
    const timeEstimate = document.getElementById("time-estimate");
    const progressBarFill = document.getElementById("progress-bar-fill");
    const progressPercent = document.getElementById("progress-percent");

    const btnConvert = document.getElementById("btn-convert");
    const logConsole = document.getElementById("log-console");

    // Modal Elements
    const successModal = document.getElementById("success-modal");
    const modalFileTitle = document.getElementById("modal-file-title");
    const btnDownload = document.getElementById("btn-download");
    const btnCloseModal = document.getElementById("btn-close-modal");

    // 2. Global State Variables
    let activeTab = "gdocs"; // "gdocs" or "local"
    let selectedFile = null;
    let currentTaskId = null;
    let pollInterval = null;

    // Detect Engine Support
    async function detectEngines() {
        try {
            const res = await fetch("/api/engines");
            if (res.ok) {
                const data = await res.json();
                const hasAuto = data.hwp_automation;
                const hasLib = data.pure_lib;

                appendLog(`엔진 감지: [한글 OLE 자동화: ${hasAuto ? "사용가능" : "미설치"}], [오픈소스 라이브러리: ${hasLib ? "사용가능" : "미설치"}]`, "sys");

                if (!hasAuto) {
                    engineAuto.disabled = true;
                    engineAuto.parentElement.style.color = "#ff5555";
                    engineAuto.parentElement.style.textDecoration = "line-through";
                    engineAuto.parentElement.title = "로컬 윈도우 환경에 아래아한글 프로그램이 설치되어 있지 않습니다.";
                    
                    const engineLib = document.getElementById("engine-lib");
                    if (engineLib) {
                        engineLib.checked = true;
                        appendLog("로컬 한글 미감지로 인해 오픈소스 라이브러리 엔진이 기본 활성화되었습니다.", "sys");
                    }
                }
                
                const engineLib = document.getElementById("engine-lib");
                if (engineLib && !hasLib) {
                    engineLib.disabled = true;
                    engineLib.parentElement.style.color = "#ff5555";
                    engineLib.parentElement.style.textDecoration = "line-through";
                }
            }
        } catch (err) {
            appendLog(`엔진 감지 실패: ${err.message}`, "error");
        }
    }
    
    // 즉시 감지 실행
    detectEngines();

    // 3. Tab Switching Logic
    tabGdocs.addEventListener("click", () => {
        if (activeTab === "gdocs") return;
        activeTab = "gdocs";
        tabGdocs.classList.add("active");
        tabLocal.classList.remove("active");
        panelGdocs.classList.add("active");
        panelLocal.classList.remove("active");
        appendLog("모드 전환: 구글 Docs 링크 변환", "sys");
    });

    tabLocal.addEventListener("click", () => {
        if (activeTab === "local") return;
        activeTab = "local";
        tabLocal.classList.add("active");
        tabGdocs.classList.remove("active");
        panelLocal.classList.add("active");
        panelGdocs.classList.remove("active");
        appendLog("모드 전환: 로컬 파일 업로드 변환 (PDF/DOCX)", "sys");
    });

    // 4. File Selection / Drag & Drop Logic
    // Select File Button
    btnSelectFile.addEventListener("click", (e) => {
        e.stopPropagation(); // Drop zone 클릭 이벤트 전파 차단
        fileInput.click();
    });

    dropZone.addEventListener("click", () => {
        fileInput.click();
    });

    fileInput.addEventListener("change", (e) => {
        if (e.target.files.length > 0) {
            handleFileSelect(e.target.files[0]);
        }
    });

    // Drag and Drop Events
    ["dragenter", "dragover"].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropZone.classList.add("dragover");
        }, false);
    });

    ["dragleave", "drop"].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropZone.classList.remove("dragover");
        }, false);
    });

    dropZone.addEventListener("drop", (e) => {
        const dt = e.dataTransfer;
        const files = dt.files;
        if (files.length > 0) {
            handleFileSelect(files[0]);
        }
    });

    // Remove selected file
    btnRemoveFile.addEventListener("click", (e) => {
        e.stopPropagation();
        resetFileSelection();
        appendLog("파일 선택 해제 완료", "sys");
    });

    function handleFileSelect(file) {
        const ext = file.name.substring(file.name.lastIndexOf(".")).toLowerCase();
        if (ext !== ".pdf" && ext !== ".docx") {
            appendLog(`오류: 지원하지 않는 파일 형식입니다 (${ext}). PDF 또는 DOCX 파일을 선택하세요.`, "error");
            return;
        }
        selectedFile = file;
        selectedFileName.textContent = file.name;
        dropZone.style.display = "none";
        fileBanner.style.display = "flex";
        appendLog(`파일 로드됨: ${file.name} (${formatBytes(file.size)})`, "sys");
        if (file.size > 4 * 1024 * 1024) {
            appendLog(`대용량 파일 감지 (${formatBytes(file.size)}). 변환에 다소 시간이 소요될 수 있습니다.`, "sys");
        }
    }

    function resetFileSelection() {
        selectedFile = null;
        fileInput.value = "";
        selectedFileName.textContent = "";
        fileBanner.style.display = "none";
        dropZone.style.display = "flex";
    }

    // 5. App Logging Helper
    function appendLog(text, type = "info") {
        const entry = document.createElement("p");
        entry.className = `log-entry ${type}`;
        
        const now = new Date();
        const timeStr = `[${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}:${now.getSeconds().toString().padStart(2, '0')}]`;
        
        entry.textContent = `${timeStr} ${text}`;
        logConsole.appendChild(entry);
        logConsole.scrollTop = logConsole.scrollHeight;
    }

    // Size formatting helper
    function formatBytes(bytes, decimals = 2) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const dm = decimals < 0 ? 0 : decimals;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
    }

    // 6. Conversion Trigger
    btnConvert.addEventListener("click", async () => {
        // Validation check
        const useAutomation = document.querySelector('input[name="engine-type"]:checked').value;
        const translateToKo = document.querySelector('input[name="trans-type"]:checked').value;
        
        const formData = new FormData();
        formData.append("use_automation", useAutomation);
        formData.append("translate_to_ko", translateToKo);

        if (activeTab === "gdocs") {
            const url = gdocsUrlInput.value.trim();
            if (!url) {
                appendLog("오류: Google Docs URL을 입력하세요.", "error");
                alert("구글 Docs 공유 링크를 입력해 주세요.");
                return;
            }
            formData.append("url", url);
            appendLog("Google Docs 변환 요청을 준비 중...", "sys");
        } else {
            if (!selectedFile) {
                appendLog("오류: 변환 대상 파일(PDF/DOCX)을 업로드해 주세요.", "error");
                alert("변환할 파일을 업로드하거나 공유 링크를 입력해 주세요.");
                return;
            }
            formData.append("file", selectedFile);
            appendLog(`로컬 파일 변환 요청 준비 중 (${selectedFile.name})...`, "sys");
        }

        // UI State: Disable Inputs
        setUIStateLoading(true);
        appendLog("서버에 변환 태스크 전송 중...", "process");

        try {
            const response = await fetch("/api/convert", {
                method: "POST",
                body: formData
            });

            if (!response.ok) {
                let errMsg = "변환 요청이 실패했습니다.";
                try {
                    const errDetail = await response.json();
                    errMsg = errDetail.detail || errMsg;
                } catch (_) {}
                throw new Error(errMsg);
            }

            const data = await response.json();
            currentTaskId = data.task_id;
            const estTime = data.estimated_seconds;

            appendLog(`변환 작업이 백그라운드로 성공적으로 등록되었습니다. (ID: ${currentTaskId.substring(0, 8)}...)`, "process");
            appendLog(`작업 예상 소요 시간: 약 ${estTime}초`, "process");

            // Start Polling Status
            startStatusPolling(currentTaskId, estTime);

        } catch (error) {
            appendLog(`오류 발생: ${error.message}`, "error");
            setUIStateLoading(false);
        }
    });

    // 7. Status Polling Control
    function startStatusPolling(taskId, estTime) {
        progressSection.style.display = "block";
        progressBarFill.style.width = "0%";
        progressPercent.textContent = "0%";
        statusMessage.textContent = "변환 서버 대기 중...";
        timeEstimate.textContent = `남은 예상 시간: 약 ${estTime}초`;

        let counter = 0;

        let consecutiveErrors = 0;

        pollInterval = setInterval(async () => {
            try {
                const res = await fetch(`/api/status/${taskId}`);

                if (res.status === 502 || res.status === 503) {
                    consecutiveErrors++;
                    if (consecutiveErrors >= 3) {
                        clearInterval(pollInterval);
                        const oomMsg = "서버가 메모리 부족(OOM)으로 재시작되었습니다. 잠시 후 다시 시도해 주세요.";
                        appendLog(`오류: ${oomMsg}`, "error");
                        alert(`변환 실패: ${oomMsg}`);
                        setUIStateLoading(false);
                    }
                    return;
                }

                if (res.status === 404) {
                    clearInterval(pollInterval);
                    const lostMsg = "서버가 재시작되어 작업 정보가 소실되었습니다. 다시 변환을 시도해 주세요.";
                    appendLog(`오류: ${lostMsg}`, "error");
                    alert(`변환 실패: ${lostMsg}`);
                    setUIStateLoading(false);
                    return;
                }

                if (!res.ok) throw new Error("상태 조회를 실패했습니다.");

                consecutiveErrors = 0;
                const data = await res.json();

                const progress = data.progress || 0;
                const message = data.message || "변환 작업 처리 중...";
                const status = data.status;
                const remaining = data.remaining_seconds ?? estTime;

                // UI 업데이트
                progressBarFill.style.width = `${progress}%`;
                progressPercent.textContent = `${progress}%`;
                statusMessage.textContent = message;
                timeEstimate.textContent = remaining > 0 ? `남은 예상 시간: 약 ${remaining}초` : "거의 완료됨...";

                if (counter % 3 === 0) {
                    appendLog(`[진행률 ${progress}%] ${message}`, "process");
                }
                counter++;

                if (status === "completed") {
                    clearInterval(pollInterval);
                    appendLog("축하합니다! 변환 작업이 성공적으로 끝났습니다.", "success");
                    showSuccessModal(taskId, data.original_name);
                    setUIStateLoading(false);
                } else if (status === "failed") {
                    clearInterval(pollInterval);
                    appendLog(`오류: ${message}`, "error");
                    alert(`변환 실패: ${message}`);
                    setUIStateLoading(false);
                }

            } catch (error) {
                consecutiveErrors++;
                if (consecutiveErrors >= 5) {
                    clearInterval(pollInterval);
                    appendLog(`서버 연결 실패가 반복됩니다. 서버가 과부하 상태일 수 있습니다.`, "error");
                    setUIStateLoading(false);
                } else {
                    appendLog(`상태 모니터링 에러: ${error.message}`, "error");
                }
            }
        }, 1000);
    }

    // UI Loading states toggle
    function setUIStateLoading(isLoading) {
        btnConvert.disabled = isLoading;
        tabGdocs.disabled = isLoading;
        tabLocal.disabled = isLoading;
        gdocsUrlInput.disabled = isLoading;
        btnSelectFile.disabled = isLoading;
        btnRemoveFile.disabled = isLoading;
        
        if (isLoading) {
            btnConvert.textContent = "⚙️ 백그라운드 변환 작업 처리 중...";
            btnConvert.style.opacity = "0.7";
        } else {
            btnConvert.textContent = "🚀 HWPX로 변환하기";
            btnConvert.style.opacity = "1";
            progressSection.style.display = "none";
        }
    }

    // 8. Modal Handling
    function showSuccessModal(taskId, fileName) {
        modalFileTitle.textContent = fileName;
        // 다운로드 버튼 액션 설정
        btnDownload.onclick = () => {
            window.location.href = `/api/download/${taskId}`;
            appendLog(`파일 다운로드 시작: ${fileName}`, "success");
        };
        successModal.style.display = "flex";
    }

    btnCloseModal.addEventListener("click", () => {
        successModal.style.display = "none";
    });

    // 배경 클릭 시 닫기 방지(실수로 닫힘 방지)
    successModal.addEventListener("click", (e) => {
        if (e.target === successModal) {
            successModal.style.display = "none";
        }
    });
});
