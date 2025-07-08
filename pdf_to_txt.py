import os
import sys
import PyPDF2
import pdfplumber
from pathlib import Path

def extract_text_with_pypdf2(pdf_path):
    """PyPDF2를 사용하여 PDF에서 텍스트 추출"""
    text = ""
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                text += page.extract_text() + "\n"
    except Exception as e:
        print(f"PyPDF2로 텍스트 추출 중 오류 발생: {e}")
        return None
    return text

def extract_text_with_pdfplumber(pdf_path):
    """pdfplumber를 사용하여 PDF에서 텍스트 추출 (더 정확함)"""
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
    except Exception as e:
        print(f"pdfplumber로 텍스트 추출 중 오류 발생: {e}")
        return None
    return text

def pdf_to_txt(pdf_path, output_path=None, method="pdfplumber"):
    """
    PDF 파일을 TXT 파일로 변환
    
    Args:
        pdf_path (str): 입력 PDF 파일 경로
        output_path (str, optional): 출력 TXT 파일 경로. None이면 자동 생성
        method (str): 추출 방법 ("pdfplumber" 또는 "pypdf2")
    
    Returns:
        bool: 변환 성공 여부
    """
    
    # PDF 파일 존재 확인
    if not os.path.exists(pdf_path):
        print(f"오류: PDF 파일을 찾을 수 없습니다: {pdf_path}")
        return False
    
    # 출력 파일 경로 설정
    if output_path is None:
        pdf_name = Path(pdf_path).stem
        output_path = f"{pdf_name}.txt"
    
    # 텍스트 추출
    print(f"PDF 파일에서 텍스트 추출 중: {pdf_path}")
    
    if method == "pdfplumber":
        extracted_text = extract_text_with_pdfplumber(pdf_path)
    elif method == "pypdf2":
        extracted_text = extract_text_with_pypdf2(pdf_path)
    else:
        print("오류: 지원하지 않는 추출 방법입니다. 'pdfplumber' 또는 'pypdf2'를 사용하세요.")
        return False
    
    if extracted_text is None:
        print("텍스트 추출에 실패했습니다.")
        return False
    
    if not extracted_text.strip():
        print("경고: 추출된 텍스트가 비어있습니다. PDF가 이미지 기반이거나 보호되어 있을 수 있습니다.")
    
    # TXT 파일로 저장
    try:
        with open(output_path, 'w', encoding='utf-8') as txt_file:
            txt_file.write(extracted_text)
        print(f"변환 완료: {output_path}")
        print(f"추출된 텍스트 길이: {len(extracted_text)} 문자")
        return True
    except Exception as e:
        print(f"파일 저장 중 오류 발생: {e}")
        return False

def batch_convert(input_folder, output_folder=None, method="pdfplumber"):
    """
    폴더 내 모든 PDF 파일을 일괄 변환
    
    Args:
        input_folder (str): PDF 파일들이 있는 폴더
        output_folder (str, optional): 출력 폴더. None이면 입력 폴더와 동일
        method (str): 추출 방법
    """
    if not os.path.exists(input_folder):
        print(f"오류: 입력 폴더를 찾을 수 없습니다: {input_folder}")
        return
    
    if output_folder and not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    pdf_files = [f for f in os.listdir(input_folder) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        print("PDF 파일을 찾을 수 없습니다.")
        return
    
    print(f"{len(pdf_files)}개의 PDF 파일을 발견했습니다.")
    
    success_count = 0
    for pdf_file in pdf_files:
        pdf_path = os.path.join(input_folder, pdf_file)
        txt_filename = Path(pdf_file).stem + ".txt"
        
        if output_folder:
            output_path = os.path.join(output_folder, txt_filename)
        else:
            output_path = os.path.join(input_folder, txt_filename)
        
        if pdf_to_txt(pdf_path, output_path, method):
            success_count += 1
        print("-" * 50)
    
    print(f"변환 완료: {success_count}/{len(pdf_files)} 파일 성공")

def main():
    """메인 함수 - 명령행 인터페이스"""
    if len(sys.argv) < 2:
        print("=" * 60)
        print("               PDF to TXT 변환기")
        print("=" * 60)
        print("\n사용법:")
        print("  단일 파일 변환: python pdf_to_txt.py <PDF파일경로> [출력파일경로] [방법]")
        print("  일괄 변환: python pdf_to_txt.py --batch <입력폴더> [출력폴더] [방법]")
        print("  GUI 실행: python pdf_to_txt.py --gui")
        print("\n방법:")
        print("  pdfplumber (기본값) - 더 정확한 텍스트 추출")
        print("  pypdf2 - 빠른 텍스트 추출")
        print("\n예시:")
        print("  python pdf_to_txt.py document.pdf")
        print("  python pdf_to_txt.py document.pdf output.txt pdfplumber")
        print("  python pdf_to_txt.py --batch ./pdfs ./texts")
        print("  python pdf_to_txt.py --gui")
        print("\n" + "=" * 60)
        return
    
    if sys.argv[1] == "--gui":
        # GUI 모드 실행
        try:
            import subprocess
            subprocess.run(["python", "pdf_to_txt_gui.py"], check=True)
        except ImportError:
            print("GUI 실행을 위해 tkinter가 필요합니다.")
        except FileNotFoundError:
            print("pdf_to_txt_gui.py 파일을 찾을 수 없습니다.")
        return
    
    if sys.argv[1] == "--batch":
        # 일괄 변환 모드
        if len(sys.argv) < 3:
            print("오류: 입력 폴더를 지정해주세요.")
            return
        
        input_folder = sys.argv[2]
        output_folder = sys.argv[3] if len(sys.argv) > 3 else None
        method = sys.argv[4] if len(sys.argv) > 4 else "pdfplumber"
        
        batch_convert(input_folder, output_folder, method)
    else:
        # 단일 파일 변환 모드
        pdf_path = sys.argv[1]
        output_path = sys.argv[2] if len(sys.argv) > 2 else None
        method = sys.argv[3] if len(sys.argv) > 3 else "pdfplumber"
        
        pdf_to_txt(pdf_path, output_path, method)

if __name__ == "__main__":
    main()