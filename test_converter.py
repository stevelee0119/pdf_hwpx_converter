import sys
import os
from core_converter import UniversalConverter, HwpAutomationEngine, LibBasedEngine

def run_test():
    print("=== HwpxConverter CLI Test Utility ===")
    
    # 1. 환경 정보 출력
    auto_avail = HwpAutomationEngine.is_available()
    lib_avail = LibBasedEngine.is_available()
    print(f"Windows HWP Automation Engine Available: {auto_avail}")
    print(f"Pure Library (Libre) Engine Available: {lib_avail}")
    print("-" * 40)

    # 2. 인자 확인
    if len(sys.argv) < 3:
        print("Usage:")
        print("  python test_converter.py <input_source> <output_path> [engine_type] [format_type]")
        print("\nArguments:")
        print("  input_source : Google Docs URL or local PDF/DOCX file path")
        print("  output_path  : Target HWP/HWPX file path")
        print("  engine_type  : 'auto' (Windows HWP OLE) or 'lib' (Pure Python), default: 'auto' if available, else 'lib'")
        print("  format_type  : 'HWPX' or 'HWP', default: 'HWPX'")
        print("\nExamples:")
        print("  python test_converter.py sample.pdf output.hwpx")
        print("  python test_converter.py \"https://docs.google.com/document/d/1Xxxx/edit\" output.hwpx auto")
        sys.exit(0)

    input_source = sys.argv[1]
    output_path = sys.argv[2]
    
    # engine_type 파싱
    engine_type = "auto"
    if len(sys.argv) >= 4:
        engine_type = sys.argv[3].lower()
    
    use_automation = True
    if engine_type == "lib":
        use_automation = False
    elif engine_type == "auto":
        use_automation = True
    else:
        # 감지된 값에 따라 자동 할당
        use_automation = auto_avail
        
    # format_type 파싱
    format_type = "HWPX"
    if len(sys.argv) >= 5:
        format_type = sys.argv[4].upper()

    print(f"[*] Starting conversion...")
    print(f"    Input  : {input_source}")
    print(f"    Output : {output_path}")
    print(f"    Engine : {'Automation' if use_automation else 'Pure Library'}")
    print(f"    Format : {format_type}")
    print("-" * 40)

    # 프로그레스 콜백
    def progress_callback(status, percent):
        print(f"[{percent}%] {status}")

    try:
        res_path = UniversalConverter.convert(
            input_source=input_source,
            output_path=output_path,
            use_automation=use_automation,
            format_type=format_type,
            progress_callback=progress_callback
        )
        print("-" * 40)
        print(f"[+] SUCCESS: Converted file saved to -> {res_path}")
    except Exception as e:
        print("-" * 40)
        print(f"[-] ERROR: Conversion failed: {e}")
        sys.exit(1)

if __name__ == "__main__":
    run_test()
