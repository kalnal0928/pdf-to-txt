import pdfplumber
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def convert_pdf_to_txt(pdf_path, txt_path):
    """
    PDF 파일의 텍스트를 추출하여 TXT 파일로 저장합니다.
    """
    if not pdf_path or not txt_path:
        messagebox.showerror("오류", "입력 및 출력 파일을 모두 지정해야 합니다.")
        return

    try:
        with pdfplumber.open(pdf_path) as pdf:
            with open(txt_path, 'w', encoding='utf-8') as txt_file:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        txt_file.write(text)
                        txt_file.write('\n\n--- 다음 페이지 ---\n\n')
        messagebox.showinfo("성공", f"파일이 성공적으로 변환되었습니다!\n저장 위치: {txt_path}")
    except Exception as e:
        messagebox.showerror("오류", f"오류가 발생했습니다: {e}")

def select_pdf_file():
    """PDF 파일을 선택하는 대화상자를 엽니다."""
    file_path = filedialog.askopenfilename(
        title="변환할 PDF 파일을 선택하세요",
        filetypes=(("PDF 파일", "*.pdf"), ("모든 파일", "*.*"))
    )
    if file_path:
        pdf_path_var.set(file_path)

def select_save_path():
    """TXT 파일을 저장할 경로를 선택하는 대화상자를 엽니다."""
    file_path = filedialog.asksaveasfilename(
        title="TXT 파일을 저장할 위치를 선택하세요",
        defaultextension=".txt",
        filetypes=(("텍스트 파일", "*.txt"), ("모든 파일", "*.*"))
    )
    if file_path:
        txt_path_var.set(file_path)

def start_conversion():
    """변환을 시작합니다."""
    pdf_path = pdf_path_var.get()
    txt_path = txt_path_var.get()
    convert_pdf_to_txt(pdf_path, txt_path)

# --- GUI 설정 ---
window = tk.Tk()
window.title("PDF to TXT 변환기")
window.geometry("500x250")

pdf_path_var = tk.StringVar()
txt_path_var = tk.StringVar()

# --- 위젯 생성 ---
# 프레임
main_frame = tk.Frame(window, padx=10, pady=10)
main_frame.pack(expand=True, fill=tk.BOTH)

# PDF 선택
pdf_frame = tk.Frame(main_frame)
pdf_frame.pack(fill=tk.X, pady=5)
pdf_label = tk.Label(pdf_frame, text="PDF 파일:", width=10, anchor='w')
pdf_label.pack(side=tk.LEFT)
pdf_entry = tk.Entry(pdf_frame, textvariable=pdf_path_var, state='readonly')
pdf_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
pdf_button = tk.Button(pdf_frame, text="찾아보기", command=select_pdf_file)
pdf_button.pack(side=tk.RIGHT)

# TXT 저장 경로 선택
txt_frame = tk.Frame(main_frame)
txt_frame.pack(fill=tk.X, pady=5)
txt_label = tk.Label(txt_frame, text="저장 위치:", width=10, anchor='w')
txt_label.pack(side=tk.LEFT)
txt_entry = tk.Entry(txt_frame, textvariable=txt_path_var, state='readonly')
txt_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
txt_button = tk.Button(txt_frame, text="찾아보기", command=select_save_path)
txt_button.pack(side=tk.RIGHT)

# 변환 버튼
convert_button = tk.Button(main_frame, text="변환 시작", command=start_conversion, height=2, bg="#4CAF50", fg="white")
convert_button.pack(fill=tk.X, pady=20)

# --- 메인 루프 시작 ---
window.mainloop()
