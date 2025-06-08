# file_utils.py
from tkinter import filedialog
import tkinter as tk
import pandas as pd
import xlwings as xw

def create_button(parent, text, command, pady=5):
    """버튼을 생성하는 유틸리티 함수"""
    button = tk.Button(parent, text=text, command=command)
    button.pack(pady=pady)
    return button

def select_folder_path():
    """사용자가 폴더를 선택하면 절대 경로를 반환한다."""
    folder_selected = filedialog.askdirectory()
    return folder_selected

def select_file_path(filetypes=(("모든 파일", "*.*"),)):
    """
    사용자가 파일을 선택하면 절대 경로를 반환한다.
    
    filetypes: 튜플 목록으로 파일 타입 필터 설정 가능.
               예: (("텍스트 파일", "*.txt"), ("모든 파일", "*.*"))
    """
    file_selected = filedialog.askopenfilename(filetypes=filetypes)
    return file_selected

# 실험용 run_folder 함수 / 이후 .py 파일로 하나씩 넘겨서 실행할 예정
def excel_to_dataframe(selected_path):
    """
    선택된 엑셀 파일 경로를 xlwings로 읽어 DataFrame으로 반환합니다.
    """
    if selected_path:
        if selected_path.endswith('.xlsx'):
            try:
                with xw.App(visible=False) as app:
                    wb = app.books.open(selected_path)
                    sht = wb.sheets[0]
                    df = sht.used_range.options(pd.DataFrame, header=1, index=False).value
                    wb.close()
                return df
            except Exception as e:
                print(f"Excel 파일을 읽는 중 오류 발생: {e}")
        else:
            print("선택된 파일이 .xlsx 형식이 아닙니다.")
    else:
        print("파일이 선택되지 않았습니다.")