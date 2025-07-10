import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import PyPDF2
import pdfplumber
from pathlib import Path

try:
    import fitz  # PyMuPDF
    FITZ_AVAILABLE = True
except ImportError:
    FITZ_AVAILABLE = False

try:
    import pytesseract
    from pdf2image import convert_from_path
    from PIL import Image
    from PIL import ImageEnhance, ImageFilter, ImageOps
    
    # Tesseract 경로 설정 (Windows의 일반적인 설치 경로들을 시도)
    import os
    tesseract_paths = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        r"C:\Users\User\AppData\Local\Tesseract-OCR\tesseract.exe"
    ]
    
    for path in tesseract_paths:
        if os.path.exists(path):
            pytesseract.pytesseract.tesseract_cmd = path
            break
    
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

# from pykospacing import Spacing
# spacing = Spacing()

def correct_korean_spacing(text):
    return text

class PDFToTxtGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to TXT 변환기")
        # OCR 기능이 있으면 더 큰 창 크기로 설정
        window_height = 800 if OCR_AVAILABLE else 650
        self.root.geometry(f"700x{window_height}")
        self.root.resizable(True, True)
        
        # 최소 크기 설정
        self.root.minsize(600, 500)
        
        # 변수 초기화
        self.pdf_files = []
        self.output_folder = ""
        self.method = tk.StringVar(value="pdfplumber")
        if OCR_AVAILABLE:
            self.ocr_lang = tk.StringVar(value="kor+eng")
            self.ocr_quality = tk.StringVar(value="고품질")
        
        self.create_widgets()
        
    def create_widgets(self):
        # 스크롤 가능한 캔버스와 프레임 생성
        canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # 그리드 설정
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # 메인 프레임 (스크롤 가능한 프레임 내부)
        main_frame = ttk.Frame(scrollable_frame, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 메인 프레임의 컬럼 가중치 설정
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)
        main_frame.grid_columnconfigure(2, weight=1)
        
        # 마우스 휠 스크롤 바인딩
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # 제목
        title_label = ttk.Label(main_frame, text="PDF to TXT 변환기", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # PDF 파일 선택 섹션
        pdf_frame = ttk.LabelFrame(main_frame, text="PDF 파일 선택", padding="10")
        pdf_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        pdf_frame.grid_columnconfigure(0, weight=1)
        pdf_frame.grid_columnconfigure(1, weight=1)
        pdf_frame.grid_columnconfigure(2, weight=1)
        
        # 단일 파일 선택
        ttk.Button(pdf_frame, text="PDF 파일 선택", command=self.select_single_file).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(pdf_frame, text="여러 PDF 파일 선택", command=self.select_multiple_files).grid(row=0, column=1, padx=(0, 10))
        ttk.Button(pdf_frame, text="폴더 선택", command=self.select_folder).grid(row=0, column=2)
        
        # 선택된 파일 목록
        self.file_listbox = tk.Listbox(pdf_frame, height=6)
        self.file_listbox.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 스크롤바
        scrollbar = ttk.Scrollbar(pdf_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        scrollbar.grid(row=1, column=3, sticky=(tk.N, tk.S), pady=(10, 0))
        self.file_listbox.config(yscrollcommand=scrollbar.set)
        
        # 파일 목록 삭제 버튼
        ttk.Button(pdf_frame, text="선택 해제", command=self.clear_files).grid(row=2, column=0, pady=(5, 0))
        
        # 출력 폴더 선택 섹션
        output_frame = ttk.LabelFrame(main_frame, text="출력 설정", padding="10")
        output_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        output_frame.grid_columnconfigure(1, weight=1)
        
        ttk.Label(output_frame, text="저장 폴더:").grid(row=0, column=0, sticky=tk.W)
        self.output_label = ttk.Label(output_frame, text="선택되지 않음 (PDF와 같은 폴더에 저장)", foreground="gray")
        self.output_label.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 0))
        
        ttk.Button(output_frame, text="폴더 선택", command=self.select_output_folder).grid(row=1, column=0, pady=(5, 0))
        ttk.Button(output_frame, text="기본값으로 설정", command=self.reset_output_folder).grid(row=1, column=1, sticky=tk.W, padx=(10, 0))
        
        # 추출 방법 선택
        method_frame = ttk.LabelFrame(main_frame, text="추출 방법", padding="10")
        method_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Radiobutton(method_frame, text="pdfplumber (권장 - 표 지원)", 
                       variable=self.method, value="pdfplumber").grid(row=0, column=0, sticky=tk.W)
        ttk.Radiobutton(method_frame, text="PyPDF2 (빠름)", 
                       variable=self.method, value="pypdf2").grid(row=1, column=0, sticky=tk.W)
        
        if FITZ_AVAILABLE:
            ttk.Radiobutton(method_frame, text="PyMuPDF (가장 강력 - 권장)", 
                           variable=self.method, value="pymupdf").grid(row=2, column=0, sticky=tk.W)
            self.method.set("pymupdf")  # PyMuPDF가 있으면 기본값으로 설정
        
        if OCR_AVAILABLE:
            ttk.Radiobutton(method_frame, text="OCR (이미지 기반 PDF용 - 느림)", 
                           variable=self.method, value="ocr").grid(row=3, column=0, sticky=tk.W)
            
        # OCR 설정 프레임 (별도 행에 배치)
        if OCR_AVAILABLE:
            ocr_frame = ttk.LabelFrame(main_frame, text="OCR 설정 (이미지 기반 PDF용)", padding="10")
            ocr_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
            
            ttk.Label(ocr_frame, text="언어:").grid(row=0, column=0, sticky=tk.W)
            lang_combo = ttk.Combobox(ocr_frame, textvariable=self.ocr_lang, width=15)
            lang_combo['values'] = ('kor+eng', 'eng', 'kor', 'jpn', 'chi_sim', 'chi_tra')
            lang_combo.grid(row=0, column=1, padx=(5, 0), sticky=tk.W)
            
            # OCR 품질 옵션 추가
            ttk.Label(ocr_frame, text="품질:").grid(row=1, column=0, sticky=tk.W)
            self.ocr_quality = tk.StringVar(value="고품질")
            quality_combo = ttk.Combobox(ocr_frame, textvariable=self.ocr_quality, width=15)
            quality_combo['values'] = ('고품질 (느림)', '표준', '빠름')
            quality_combo.grid(row=1, column=1, padx=(5, 0), sticky=tk.W)
            
            # 변환 버튼을 다음 행으로 이동
            button_row = 5
        else:
            # OCR이 없으면 기존 위치 유지
            button_row = 4
        
        # 변환 버튼
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=button_row, column=0, columnspan=3, pady=(20, 10))
        
        self.convert_button = ttk.Button(button_frame, text="🚀 변환 시작", command=self.start_conversion)
        self.convert_button.grid(row=0, column=0, padx=(0, 20))
        
        ttk.Button(button_frame, text="❌ 종료", command=self.root.quit).grid(row=0, column=1)
        
        # 진행 상황 표시
        self.status_label = ttk.Label(main_frame, text="", font=("Arial", 9))
        self.status_label.grid(row=button_row+1, column=0, columnspan=3, pady=(5, 0))
        
        self.progress = ttk.Progressbar(main_frame, mode='determinate')
        self.progress.grid(row=button_row+2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 20))
        
        # 상태 표시
        self.status_label = ttk.Label(main_frame, text="📁 PDF 파일을 선택하고 🚀 변환 시작 버튼을 클릭하세요.", font=("Arial", 9))
        self.status_label.grid(row=button_row+2, column=0, columnspan=3, pady=(10, 0))
        
        # 그리드 가중치 설정
        main_frame.columnconfigure(1, weight=1)
        pdf_frame.columnconfigure(2, weight=1)
        output_frame.columnconfigure(1, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
    def select_single_file(self):
        """단일 PDF 파일 선택"""
        file_path = filedialog.askopenfilename(
            title="PDF 파일 선택",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file_path:
            self.pdf_files = [file_path]
            self.update_file_list()
    
    def select_multiple_files(self):
        """여러 PDF 파일 선택"""
        file_paths = filedialog.askopenfilenames(
            title="PDF 파일들 선택",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file_paths:
            self.pdf_files = list(file_paths)
            self.update_file_list()
    
    def select_folder(self):
        """폴더 내 모든 PDF 파일 선택"""
        folder_path = filedialog.askdirectory(title="PDF 파일이 있는 폴더 선택")
        if folder_path:
            pdf_files = []
            for file in os.listdir(folder_path):
                if file.lower().endswith('.pdf'):
                    pdf_files.append(os.path.join(folder_path, file))
            
            if pdf_files:
                self.pdf_files = pdf_files
                self.update_file_list()
            else:
                messagebox.showwarning("경고", "선택한 폴더에 PDF 파일이 없습니다.")
    
    def update_file_list(self):
        """파일 목록 업데이트"""
        self.file_listbox.delete(0, tk.END)
        for file_path in self.pdf_files:
            self.file_listbox.insert(tk.END, os.path.basename(file_path))
        
        self.status_label.config(text=f"{len(self.pdf_files)}개의 PDF 파일이 선택되었습니다.")
    
    def clear_files(self):
        """선택된 파일 목록 지우기"""
        self.pdf_files = []
        self.file_listbox.delete(0, tk.END)
        self.status_label.config(text="📁 PDF 파일을 선택하고 🚀 변환 시작 버튼을 클릭하세요.")
    
    def select_output_folder(self):
        """출력 폴더 선택"""
        folder_path = filedialog.askdirectory(title="TXT 파일을 저장할 폴더 선택")
        if folder_path:
            self.output_folder = folder_path
            self.output_label.config(text=folder_path, foreground="black")
    
    def reset_output_folder(self):
        """출력 폴더를 기본값으로 리셋"""
        self.output_folder = ""
        self.output_label.config(text="선택되지 않음 (PDF와 같은 폴더에 저장)", foreground="gray")
    
    def start_conversion(self):
        """변환 시작"""
        if not self.pdf_files:
            messagebox.showwarning("경고", "변환할 PDF 파일을 선택해주세요.")
            return
        
        # 변환 중 버튼 비활성화 및 상태 표시
        self.convert_button.config(state='disabled', text='🔄 변환 중...')
        self.status_label.config(text="변환을 시작합니다...")
        
        # 별도 스레드에서 변환 실행
        threading.Thread(target=self.convert_files, daemon=True).start()
    
    def convert_files(self):
        """파일 변환 실행"""
        total_files = len(self.pdf_files)
        self.progress.config(maximum=total_files)
        success_count = 0
        
        for i, pdf_path in enumerate(self.pdf_files):
            # 상태 업데이트
            filename = os.path.basename(pdf_path)
            self.root.after(0, lambda f=filename: self.status_label.config(text=f"변환 중: {f}"))
            
            # 출력 파일 경로 결정
            if self.output_folder:
                output_path = os.path.join(self.output_folder, Path(pdf_path).stem + ".txt")
            else:
                output_path = os.path.join(os.path.dirname(pdf_path), Path(pdf_path).stem + ".txt")
            
            # 변환 실행
            if self.convert_single_file(pdf_path, output_path):
                success_count += 1
            
            # 진행률 업데이트
            self.root.after(0, lambda v=i+1: self.progress.config(value=v))
        
        # 완료 메시지
        self.root.after(0, lambda: self.conversion_complete(success_count, total_files))
    
    def convert_single_file(self, pdf_path, output_path):
        """단일 파일 변환"""
        try:
            method = self.method.get()
            
            if method == "pymupdf":
                extracted_text = self.extract_text_with_pymupdf(pdf_path)
            elif method == "pdfplumber":
                extracted_text = self.extract_text_with_pdfplumber(pdf_path)
            elif method == "ocr":
                extracted_text = self.extract_text_with_ocr(pdf_path)
            else:
                extracted_text = self.extract_text_with_pypdf2(pdf_path)
            
            # 텍스트 추출 실패 시 다른 방법들을 순차적으로 시도
            if extracted_text is None or extracted_text.startswith("오류:"):
                print(f"첫 번째 방법({method}) 실패, 다른 방법으로 재시도: {pdf_path}")
                
                # 모든 방법을 시도해보기
                methods_to_try = []
                if method != "pymupdf" and FITZ_AVAILABLE:
                    methods_to_try.append(("pymupdf", self.extract_text_with_pymupdf))
                if method != "pdfplumber":
                    methods_to_try.append(("pdfplumber", self.extract_text_with_pdfplumber))
                if method != "pypdf2":
                    methods_to_try.append(("pypdf2", self.extract_text_with_pypdf2))
                if method != "ocr" and OCR_AVAILABLE:
                    methods_to_try.append(("ocr", self.extract_text_with_ocr))
                
                for method_name, extract_func in methods_to_try:
                    print(f"시도 중: {method_name}")
                    extracted_text = extract_func(pdf_path)
                    if extracted_text and not extracted_text.startswith("오류:"):
                        print(f"{method_name}로 성공!")
                        break
            
            # 여전히 실패한 경우
            if extracted_text is None or extracted_text.startswith("오류:"):
                error_msg = extracted_text if extracted_text else "알 수 없는 오류"
                with open(output_path, 'w', encoding='utf-8') as txt_file:
                    txt_file.write(f"PDF 텍스트 추출 실패\n")
                    txt_file.write(f"파일: {pdf_path}\n")
                    txt_file.write(f"오류: {error_msg}\n")
                    txt_file.write(f"\n이 파일은 다음 중 하나일 수 있습니다:\n")
                    txt_file.write(f"- 이미지 기반 PDF (OCR 방법을 시도해보세요)\n")
                    txt_file.write(f"- 암호화된 PDF (비밀번호 필요)\n")
                    txt_file.write(f"- 손상된 PDF 파일\n")
                    txt_file.write(f"- 특수 포맷이나 복잡한 레이아웃\n")
                    txt_file.write(f"- 폰트가 임베드되지 않은 PDF\n")
                    if OCR_AVAILABLE:
                        txt_file.write(f"\n* 이미지 기반 PDF인 경우 'OCR' 방법을 선택하여 다시 시도해보세요.\n")
                    else:
                        txt_file.write(f"\n* OCR 기능을 사용하려면 다음 패키지를 설치하세요:\n")
                        txt_file.write(f"  pip install pytesseract pdf2image pillow\n")
                        txt_file.write(f"  그리고 Tesseract OCR 엔진을 설치하세요.\n")
                return False
            
            # 성공적으로 텍스트를 추출한 경우
            with open(output_path, 'w', encoding='utf-8') as txt_file:
                txt_file.write(extracted_text)
            
            return True
        except Exception as e:
            print(f"변환 오류: {e}")
            # 오류 정보를 파일로 저장
            try:
                with open(output_path, 'w', encoding='utf-8') as txt_file:
                    txt_file.write(f"PDF 변환 중 오류 발생\n")
                    txt_file.write(f"파일: {pdf_path}\n")
                    txt_file.write(f"오류: {str(e)}\n")
            except:
                pass
            return False
    
    def extract_text_with_pypdf2(self, pdf_path):
        """PyPDF2를 사용하여 PDF에서 텍스트 추출"""
        text = ""
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                
                # 암호화 체크
                if pdf_reader.is_encrypted:
                    print(f"암호화된 PDF 파일입니다: {pdf_path}")
                    return "오류: 암호화된 PDF 파일입니다. 비밀번호가 필요합니다."
                
                # 페이지 수 체크
                if len(pdf_reader.pages) == 0:
                    print(f"페이지가 없는 PDF 파일입니다: {pdf_path}")
                    return "오류: 페이지가 없는 PDF 파일입니다."
                
                for page_num in range(len(pdf_reader.pages)):
                    page = pdf_reader.pages[page_num]
                    page_text = page.extract_text()
                    if page_text.strip():  # 빈 텍스트가 아닌 경우만 추가
                        text += page_text + "\n"
                    else:
                        text += f"[페이지 {page_num + 1}: 텍스트 추출 불가 - 이미지 기반일 수 있음]\n"
                        
        except Exception as e:
            print(f"PyPDF2로 텍스트 추출 중 오류 발생: {e}")
            return f"오류: {str(e)}"
        
        if not text.strip():
            return "오류: 추출된 텍스트가 없습니다. 이미지 기반 PDF이거나 텍스트가 없는 파일일 수 있습니다."
        
        return text

    def extract_text_with_pdfplumber(self, pdf_path):
        """pdfplumber를 사용하여 PDF에서 텍스트 추출"""
        text = ""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                # 페이지 수 체크
                if len(pdf.pages) == 0:
                    return "오류: 페이지가 없는 PDF 파일입니다."
                
                for page_num, page in enumerate(pdf.pages):
                    page_text = page.extract_text()
                    if page_text and page_text.strip():
                        text += page_text + "\n"
                    else:
                        # 표나 다른 요소 추출 시도
                        tables = page.extract_tables()
                        if tables:
                            for table in tables:
                                for row in table:
                                    if row:
                                        text += " ".join([cell if cell else "" for cell in row]) + "\n"
                        else:
                            text += f"[페이지 {page_num + 1}: 텍스트 추출 불가 - 이미지 기반일 수 있음]\n"
                            
        except Exception as e:
            print(f"pdfplumber로 텍스트 추출 중 오류 발생: {e}")
            return f"오류: {str(e)}"
        
        if not text.strip():
            return "오류: 추출된 텍스트가 없습니다. 이미지 기반 PDF이거나 텍스트가 없는 파일일 수 있습니다."
        
        return text
    
    def extract_text_with_pymupdf(self, pdf_path):
        """PyMuPDF(fitz)를 사용하여 PDF에서 텍스트 추출 - 가장 강력한 방법"""
        if not FITZ_AVAILABLE:
            return "오류: PyMuPDF가 설치되지 않았습니다."
        
        text = ""
        try:
            doc = fitz.open(pdf_path)
            
            # 암호화 체크
            if doc.needs_pass:
                doc.close()
                return "오류: 암호화된 PDF 파일입니다. 비밀번호가 필요합니다."
            
            # 페이지 수 체크
            if doc.page_count == 0:
                doc.close()
                return "오류: 페이지가 없는 PDF 파일입니다."
            
            for page_num in range(doc.page_count):
                page = doc[page_num]
                page_text = page.get_text()
                
                if page_text.strip():
                    text += page_text + "\n"
                else:
                    # 이미지에서 텍스트 추출 시도 (OCR 없이)
                    blocks = page.get_text("dict")
                    if blocks.get("blocks"):
                        text += f"[페이지 {page_num + 1}: 구조화된 내용 감지됨]\n"
                        for block in blocks["blocks"]:
                            if "lines" in block:
                                for line in block["lines"]:
                                    for span in line["spans"]:
                                        if span.get("text", "").strip():
                                            text += span["text"] + " "
                                text += "\n"
                    else:
                        text += f"[페이지 {page_num + 1}: 텍스트 추출 불가 - 이미지 기반일 수 있음]\n"
            
            doc.close()
            
        except Exception as e:
            print(f"PyMuPDF로 텍스트 추출 중 오류 발생: {e}")
            return f"오류: {str(e)}"
        
        if not text.strip():
            return "오류: 추출된 텍스트가 없습니다. 이미지 기반 PDF이거나 텍스트가 없는 파일일 수 있습니다."
        
        return text
    
    def extract_text_with_ocr(self, pdf_path):
        """OCR을 사용하여 이미지 기반 PDF에서 텍스트 추출"""
        if not OCR_AVAILABLE:
            return "오류: OCR 라이브러리(pytesseract, pdf2image)가 설치되지 않았습니다."
        
        text = ""
        try:
            # PDF를 이미지로 변환 (여러 방법 시도)
            print(f"OCR 처리 시작: {pdf_path}")
            
            # 방법 1: 기본 설정으로 시도
            images = None
            try:
                images = convert_from_path(pdf_path, dpi=300, fmt='jpeg')
            except Exception as e1:
                print(f"방법 1 실패: {e1}")
                
                # 방법 2: 낮은 DPI로 시도
                try:
                    images = convert_from_path(pdf_path, dpi=200)
                except Exception as e2:
                    print(f"방법 2 실패: {e2}")
                    
                    # 방법 3: PyMuPDF로 이미지 추출 후 OCR
                    try:
                        if FITZ_AVAILABLE:
                            return self.extract_text_with_fitz_ocr(pdf_path)
                        else:
                            return f"오류: PDF를 이미지로 변환할 수 없습니다. Poppler가 설치되지 않았을 수 있습니다.\n원본 오류: {str(e1)}"
                    except Exception as e3:
                        return f"오류: 모든 OCR 방법이 실패했습니다.\n오류들: {str(e1)}, {str(e2)}, {str(e3)}"
            
            if not images:
                return "오류: PDF를 이미지로 변환할 수 없습니다."
            
            # 언어 설정 가져오기
            lang = self.ocr_lang.get() if hasattr(self, 'ocr_lang') else 'kor+eng'
            quality = self.ocr_quality.get() if hasattr(self, 'ocr_quality') else '고품질'
            
            for page_num, image in enumerate(images):
                print(f"OCR 처리 중: 페이지 {page_num + 1}/{len(images)}")
                
                # 품질 설정에 따른 이미지 전처리
                if quality.startswith('고품질'):
                    # 고품질: 최대한 정확한 OCR
                    image = image.convert('L')  # 그레이스케일 변환
                    
                    # 이미지 크기 조정 (너무 작으면 확대)
                    width, height = image.size
                    if width < 2000:  # 더 높은 해상도로 확대
                        scale_factor = 2000 / width
                        new_width = int(width * scale_factor)
                        new_height = int(height * scale_factor)
                        image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
                    
                    # 이미지 품질 향상
                    try:
                        from PIL import ImageEnhance, ImageFilter
                        
                        # 대비 향상
                        enhancer = ImageEnhance.Contrast(image)
                        image = enhancer.enhance(1.5)  # 대비를 더 강하게
                        
                        # 선명도 향상
                        enhancer = ImageEnhance.Sharpness(image)
                        image = enhancer.enhance(1.3)  # 선명도를 더 강하게
                        
                        # 밝기 조정
                        enhancer = ImageEnhance.Brightness(image)
                        image = enhancer.enhance(1.1)  # 약간 밝게
                    except ImportError:
                        pass  # PIL의 ImageEnhance가 없으면 건너뛰기
                        
                elif quality == '표준':
                    # 표준: 균형잡힌 처리
                    image = image.convert('L')  # 그레이스케일 변환
                    
                    width, height = image.size
                    if width < 1500:  # 표준 해상도로 확대
                        scale_factor = 1500 / width
                        new_width = int(width * scale_factor)
                        new_height = int(height * scale_factor)
                        image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
                    
                    try:
                        from PIL import ImageEnhance
                        enhancer = ImageEnhance.Contrast(image)
                        image = enhancer.enhance(1.2)
                        enhancer = ImageEnhance.Sharpness(image)
                        image = enhancer.enhance(1.1)
                    except ImportError:
                        pass
                        
                else:  # 빠름
                    # 빠름: 최소한의 처리
                    image = image.convert('L')  # 그레이스케일 변환만
                
                # Tesseract OCR 설정을 품질에 따라 조정
                if quality.startswith('고품질'):
                    ocr_configs = [
                        '--psm 6 -c preserve_interword_spaces=1 -c textord_really_old_xheight=1',  # 단일 텍스트 블록
                        '--psm 4 -c preserve_interword_spaces=1 -c textord_really_old_xheight=1',  # 단일 열, 공백 보존
                        '--psm 1 -c preserve_interword_spaces=1',  # 자동 페이지 분할
                        '--psm 3 -c textord_really_old_xheight=1',  # 완전 자동
                        '--psm 8 -c preserve_interword_spaces=1',  # 단일 단어
                    ]
                elif quality == '표준':
                    ocr_configs = [
                        '--psm 6 -c preserve_interword_spaces=1',  # 단일 텍스트 블록
                        '--psm 4 -c preserve_interword_spaces=1',  # 단일 열
                        '--psm 3',  # 완전 자동
                    ]
                else:  # 빠름
                    ocr_configs = [
                        '--psm 6',  # 단일 텍스트 블록만
                    ]
                
                page_text = ""
                for config in ocr_configs:
                    try:
                        page_text = pytesseract.image_to_string(image, lang=lang, config=config)
                        if page_text.strip():
                            break
                    except Exception as e:
                        print(f"OCR 설정 {config} 실패: {e}")
                        continue
                
                if page_text.strip():
                    # 텍스트 후처리 - 불필요한 공백 제거 및 정리
                    lines = page_text.strip().split('\n')
                    cleaned_lines = []
                    for line in lines:
                        line = line.strip()
                        # 너무 짧거나 특수 문자만 있는 라인 제거
                        if line and len(line) > 1 and not line.replace(' ', '').replace('.', '').replace('_', '').replace('-', '') == '':
                            # 일반적인 OCR 오류 패턴 수정
                            line = line.replace('|', 'I')  # 세로선을 I로
                            line = line.replace('０', '0')  # 전각 숫자를 반각으로
                            line = line.replace('１', '1')
                            line = line.replace('２', '2')
                            line = line.replace('３', '3')
                            line = line.replace('４', '4')
                            line = line.replace('５', '5')
                            line = line.replace('６', '6')
                            line = line.replace('７', '7')
                            line = line.replace('８', '8')
                            line = line.replace('９', '9')
                            cleaned_lines.append(line)
                    
                    if cleaned_lines:
                        text += f"[페이지 {page_num + 1}]\n"
                        text += '\n'.join(cleaned_lines) + "\n\n"
                    else:
                        text += f"[페이지 {page_num + 1}: OCR로 텍스트를 추출할 수 없음]\n\n"
                else:
                    text += f"[페이지 {page_num + 1}: OCR로 텍스트를 추출할 수 없음]\n\n"
            
        except Exception as e:
            print(f"OCR로 텍스트 추출 중 오류 발생: {e}")
            return f"오류: OCR 처리 실패 - {str(e)}"
        
        if not text.strip():
            return "오류: OCR로 추출된 텍스트가 없습니다. 이미지 품질이 낮거나 텍스트가 없는 파일일 수 있습니다."
        
        return text
    
    def extract_text_with_fitz_ocr(self, pdf_path):
        """PyMuPDF로 이미지를 추출한 후 OCR 처리"""
        if not FITZ_AVAILABLE:
            return "오류: PyMuPDF가 설치되지 않았습니다."
        
        text = ""
        try:
            import fitz
            doc = fitz.open(pdf_path)
            
            # 언어 설정 가져오기
            lang = self.ocr_lang.get() if hasattr(self, 'ocr_lang') else 'kor+eng'
            quality = self.ocr_quality.get() if hasattr(self, 'ocr_quality') else '고품질'
            
            for page_num in range(doc.page_count):
                print(f"PyMuPDF OCR 처리 중: 페이지 {page_num + 1}/{doc.page_count}")
                page = doc[page_num]
                
                # 페이지를 고해상도 이미지로 변환 (3x 확대로 품질 향상)
                mat = fitz.Matrix(3.0, 3.0)  # 3x 확대로 더 좋은 품질
                pix = page.get_pixmap(matrix=mat)
                img_data = pix.tobytes("png")
                
                # PIL Image로 변환
                from io import BytesIO
                image = Image.open(BytesIO(img_data))
                
                # 이미지 전처리
                # 1. 그레이스케일 변환
                image = image.convert('L')
                
                # 2. 이미지 개선 (선명도 향상)
                from PIL import ImageEnhance, ImageFilter
                
                # 대비 향상
                enhancer = ImageEnhance.Contrast(image)
                image = enhancer.enhance(1.5)  # 대비를 더 강하게
                
                # 선명도 향상
                enhancer = ImageEnhance.Sharpness(image)
                image = enhancer.enhance(1.3)  # 선명도를 더 강하게
                
                # 밝기 조정
                enhancer = ImageEnhance.Brightness(image)
                image = enhancer.enhance(1.1)  # 약간 밝게
                
                # OCR 처리 - 품질에 따른 설정
                page_text = ""
                if quality.startswith('고품질'):
                    ocr_configs = [
                        '--psm 6 -c preserve_interword_spaces=1 -c textord_really_old_xheight=1',  # 단일 텍스트 블록
                        '--psm 4 -c preserve_interword_spaces=1 -c textord_really_old_xheight=1',  # 단일 열, 공백 보존
                        '--psm 1 -c preserve_interword_spaces=1',  # 자동 페이지 분할
                        '--psm 3 -c textord_really_old_xheight=1',  # 완전 자동
                        '--psm 8 -c preserve_interword_spaces=1',  # 단일 단어
                    ]
                elif quality == '표준':
                    ocr_configs = [
                        '--psm 6 -c preserve_interword_spaces=1',  # 단일 텍스트 블록
                        '--psm 4 -c preserve_interword_spaces=1',  # 단일 열
                        '--psm 3',  # 완전 자동
                    ]
                else:  # 빠름
                    ocr_configs = [
                        '--psm 6',  # 단일 텍스트 블록만
                    ]
                
                for config in ocr_configs:
                    try:
                        page_text = pytesseract.image_to_string(image, lang=lang, config=config)
                        if page_text.strip():
                            break
                    except Exception as e:
                        print(f"OCR 설정 {config} 실패: {e}")
                        continue
                
                if page_text.strip():
                    # 텍스트 후처리 - 불필요한 공백 제거 및 정리
                    lines = page_text.strip().split('\n')
                    cleaned_lines = []
                    for line in lines:
                        line = line.strip()
                        # 너무 짧거나 특수 문자만 있는 라인 제거
                        if line and len(line) > 1 and not line.replace(' ', '').replace('.', '').replace('_', '').replace('-', '') == '':
                            # 일반적인 OCR 오류 패턴 수정
                            line = line.replace('|', 'I')  # 세로선을 I로
                            line = line.replace('０', '0')  # 전각 숫자를 반각으로
                            line = line.replace('１', '1')
                            line = line.replace('２', '2')
                            line = line.replace('３', '3')
                            line = line.replace('４', '4')
                            line = line.replace('５', '5')
                            line = line.replace('６', '6')
                            line = line.replace('７', '7')
                            line = line.replace('８', '8')
                            line = line.replace('９', '9')
                            cleaned_lines.append(line)
                    
                    if cleaned_lines:
                        text += f"[페이지 {page_num + 1}]\n"
                        text += '\n'.join(cleaned_lines) + "\n\n"
                    else:
                        text += f"[페이지 {page_num + 1}: OCR로 텍스트를 추출할 수 없음]\n\n"
                else:
                    text += f"[페이지 {page_num + 1}: OCR로 텍스트를 추출할 수 없음]\n\n"
            
            doc.close()
            
        except Exception as e:
            print(f"PyMuPDF OCR로 텍스트 추출 중 오류 발생: {e}")
            return f"오류: PyMuPDF OCR 처리 실패 - {str(e)}"
        
        text = correct_korean_spacing(text)
        return text
    
    def conversion_complete(self, success_count, total_files):
        """변환 완료 처리"""
        self.convert_button.config(state='normal', text='🚀 변환 시작')
        self.progress.config(value=0)
        
        if success_count == total_files:
            self.status_label.config(text=f"✅ 모든 변환 완료! ({success_count}/{total_files})")
            messagebox.showinfo("완료", f"모든 파일 변환이 완료되었습니다!\n성공: {success_count}/{total_files}")
        else:
            self.status_label.config(text=f"⚠️ 변환 완료: {success_count}/{total_files} (일부 실패)")
            messagebox.showwarning("완료", f"변환이 완료되었습니다.\n성공: {success_count}/{total_files}\n실패한 파일이 있습니다.")

def main():
    root = tk.Tk()
    app = PDFToTxtGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
