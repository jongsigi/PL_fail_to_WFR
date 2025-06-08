# FolderSelectorApp.py
import tkinter as tk
from utils.file_utils import *
from utils.extract_wfr import *

class FolderSelectorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("파일 선택기")
        
        # 선택된 파일 경로를 저장할 인스턴스 변수
        self.selected_file_path_1 = None
        self.selected_file_path_2 = None
        self.df = None
        
        # 라벨 생성
        self.label_var = tk.StringVar()
        self.label_var.set("파일을 선택해주세요.")
        self.label = tk.Label(root, textvariable=self.label_var, wraplength=400, height=4)
        self.label.pack(pady=10) 
        
        # 첫 번째 파일 선택 버튼
        self.file_button_1 = create_button(root, "첫 번째 파일 선택", self.on_select_file_1)
        
        # 실행 버튼
        self.run_button = create_button(root, "Run", self.run_files)

    def on_select_file_1(self):
        # 첫 번째 파일 선택 대화상자 열기
        path = select_file_path()
        if path:
            self.selected_file_path_1 = path
            self.label_var.set(f"첫 번째 파일 선택됨: {path}")
        else:
            self.label_var.set("첫 번째 파일이 선택되지 않았습니다.")

    def run_files(self):
        # 두 파일 경로를 확인하고 처리
        if self.selected_file_path_1:
            print(f"첫 번째 파일: {self.selected_file_path_1}")
            df = excel_to_dataframe(self.selected_file_path_1)
            result_df = process_elec_items(df)
            save_to_excel(self.selected_file_path_1, result_df, 'Parsed_Results')
        else:
            print("두 파일이 모두 선택되지 않았습니다.")
            self.label_var.set("두 파일이 모두 선택되지 않았습니다.")