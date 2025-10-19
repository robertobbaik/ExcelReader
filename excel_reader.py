import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
import pandas as pd
from collections import Counter
import json
import os

class ExcelReaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to C# Class & JSON Generator")
        self.root.geometry("900x750")
        
        self.current_df = None
        self.current_file = None
        self.output_directory = None
        
        # 상단 프레임
        top_frame = tk.Frame(root)
        top_frame.pack(pady=20, padx=10, fill=tk.X)
        
        # 파일 선택 버튼
        self.btn_select = tk.Button(
            top_frame, 
            text="Excel 파일 선택", 
            command=self.select_file,
            font=("Arial", 12),
            bg="#4CAF50",
            fg="white",
            padx=20,
            pady=10
        )
        self.btn_select.pack(side=tk.LEFT, padx=5)
        
        # 저장 디렉토리 선택 버튼
        self.btn_select_dir = tk.Button(
            top_frame,
            text="저장 폴더 선택",
            command=self.select_output_directory,
            font=("Arial", 12),
            bg="#FF9800",
            fg="white",
            padx=20,
            pady=10
        )
        self.btn_select_dir.pack(side=tk.LEFT, padx=5)
        
        # C# 클래스 & JSON 생성 버튼
        self.btn_generate = tk.Button(
            top_frame,
            text="Convert (C# + JSON)",
            command=self.generate_files,
            font=("Arial", 12),
            bg="#2196F3",
            fg="white",
            padx=20,
            pady=10,
            state=tk.DISABLED
        )
        self.btn_generate.pack(side=tk.LEFT, padx=5)
        
        # 파일 경로 표시
        self.label_path = tk.Label(root, text="선택된 파일: 없음", font=("Arial", 10))
        self.label_path.pack(pady=5)
        
        # 저장 디렉토리 표시
        self.label_output_dir = tk.Label(root, text="저장 폴더: 미지정 (현재 폴더에 저장됨)", font=("Arial", 10), fg="gray")
        self.label_output_dir.pack(pady=5)
        
        # 클래스명 입력
        class_frame = tk.Frame(root)
        class_frame.pack(pady=10)
        
        tk.Label(class_frame, text="클래스명:", font=("Arial", 10)).pack(side=tk.LEFT, padx=5)
        self.entry_class_name = tk.Entry(class_frame, font=("Arial", 10), width=30)
        self.entry_class_name.pack(side=tk.LEFT, padx=5)
        self.entry_class_name.insert(0, "MyClass")
        
        # 안내 메시지
        info_label = tk.Label(root, text="💡 팁: 컬럼명이 ~로 시작하면 해당 컬럼은 무시됩니다", font=("Arial", 9), fg="gray")
        info_label.pack(pady=5)
        
        # 결과 표시 영역
        frame = tk.Frame(root)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 스크롤바
        scrollbar_y = tk.Scrollbar(frame)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        scrollbar_x = tk.Scrollbar(frame, orient=tk.HORIZONTAL)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 텍스트 영역
        self.text_result = tk.Text(
            frame, 
            wrap=tk.NONE,
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set,
            font=("Courier", 10)
        )
        self.text_result.pack(fill=tk.BOTH, expand=True)
        
        scrollbar_y.config(command=self.text_result.yview)
        scrollbar_x.config(command=self.text_result.xview)
    
    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Excel 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            self.current_file = file_path
            self.label_path.config(text=f"선택된 파일: {file_path}")
            self.read_excel(file_path)
    
    def select_output_directory(self):
        directory = filedialog.askdirectory(title="저장 폴더 선택")
        
        if directory:
            self.output_directory = directory
            self.label_output_dir.config(text=f"저장 폴더: {directory}", fg="blue")
    
    def read_excel(self, file_path):
        try:
            # Excel 파일 읽기 (헤더 없이)
            self.current_df = pd.read_excel(file_path, header=None)
            
            # 결과 표시
            self.text_result.delete(1.0, tk.END)
            self.text_result.insert(tk.END, "=== Excel 데이터 미리보기 ===\n\n")
            
            if len(self.current_df) >= 2:
                self.text_result.insert(tk.END, "첫 번째 행 (변수명):\n")
                self.text_result.insert(tk.END, str(self.current_df.iloc[0].values) + "\n\n")
                self.text_result.insert(tk.END, "두 번째 행 (자료형):\n")
                self.text_result.insert(tk.END, str(self.current_df.iloc[1].values) + "\n\n")
                
                # ~ 로 시작하는 컬럼 확인
                ignored_columns = [str(col) for col in self.current_df.iloc[0].values if str(col).startswith('~')]
                if ignored_columns:
                    self.text_result.insert(tk.END, f"무시되는 컬럼: {', '.join(ignored_columns)}\n\n")
            
            self.text_result.insert(tk.END, self.current_df.head(10).to_string())
            self.text_result.insert(tk.END, f"\n\n총 행 수: {len(self.current_df)}")
            self.text_result.insert(tk.END, f"\n총 열 수: {len(self.current_df.columns)}")
            
            # Convert 버튼 활성화
            self.btn_generate.config(state=tk.NORMAL)
            
            messagebox.showinfo("성공", f"파일을 성공적으로 읽었습니다!\n행 수: {len(self.current_df)}")
            
        except Exception as e:
            messagebox.showerror("오류", f"파일을 읽는 중 오류가 발생했습니다:\n{str(e)}")
    
    def generate_files(self):
        if self.current_df is None:
            messagebox.showwarning("경고", "먼저 Excel 파일을 선택해주세요!")
            return
        
        if len(self.current_df) < 2:
            messagebox.showwarning("경고", "Excel 파일에 최소 2개의 행(변수명, 자료형)이 필요합니다!")
            return
        
        class_name = self.entry_class_name.get().strip()
        if not class_name:
            messagebox.showwarning("경고", "클래스명을 입력해주세요!")
            return
        
        # 저장 디렉토리 결정
        if self.output_directory:
            output_dir = self.output_directory
        else:
            output_dir = os.path.dirname(self.current_file) if self.current_file else os.getcwd()
        
        # C# 클래스 코드 생성
        cs_code = self.create_csharp_class_code(class_name)
        
        # JSON 데이터 생성
        json_data = self.create_json_data()
        
        # 결과 표시
        self.text_result.delete(1.0, tk.END)
        self.text_result.insert(tk.END, "=== 생성된 C# 클래스 ===\n\n")
        self.text_result.insert(tk.END, cs_code)
        self.text_result.insert(tk.END, "\n\n=== 생성된 JSON 데이터 ===\n\n")
        self.text_result.insert(tk.END, json.dumps(json_data, indent=2, ensure_ascii=False))
        
        # 파일 저장
        try:
            # C# 파일 저장
            cs_file_path = os.path.join(output_dir, f"{class_name}.cs")
            with open(cs_file_path, 'w', encoding='utf-8') as f:
                f.write(cs_code)
            
            # JSON 파일 저장
            json_file_path = os.path.join(output_dir, f"{class_name}.json")
            with open(json_file_path, 'w', encoding='utf-8') as f:
                json.dump(json_data, f, indent=2, ensure_ascii=False)
            
            messagebox.showinfo("성공", 
                f"파일이 생성되었습니다!\n\n"
                f"C# 클래스: {cs_file_path}\n"
                f"JSON 데이터: {json_file_path}")
            
        except Exception as e:
            messagebox.showerror("오류", f"파일 저장 중 오류가 발생했습니다:\n{str(e)}")
    
    def create_csharp_class_code(self, class_name):
        """C# 클래스 코드 생성"""
        # 첫 번째 행: 변수명
        variable_names = self.current_df.iloc[0].values
        # 두 번째 행: 자료형
        data_types = self.current_df.iloc[1].values
        
        # 변수명 카운트 (중복 체크) - ~ 로 시작하는 것 제외
        valid_names = [name for name in variable_names if not str(name).startswith('~')]
        name_counts = Counter(valid_names)
        
        # 이미 처리한 변수명 추적
        processed_names = set()
        
        lines = []
        lines.append("using System;")
        lines.append("using System.Collections.Generic;")
        lines.append("")
        lines.append(f"public class {class_name}")
        lines.append("{")
        
        # 각 컬럼에 대한 속성 생성
        for i in range(len(variable_names)):
            var_name = str(variable_names[i]).strip()
            var_type = str(data_types[i]).strip()
            
            # ~ 로 시작하는 컬럼 무시
            if var_name.startswith('~'):
                continue
            
            # 빈 값 체크
            if not var_name or var_name == 'nan':
                continue
            
            # PascalCase로 변환
            property_name = self.to_pascal_case(var_name)
            
            # 중복된 변수명 처리
            if name_counts[var_name] > 1:
                # 아직 처리하지 않은 경우에만 List로 추가
                if var_name not in processed_names:
                    lines.append(f"    public List<{var_type}> {property_name} {{ get; set; }}")
                    processed_names.add(var_name)
            else:
                # 중복 없으면 일반 속성
                lines.append(f"    public {var_type} {property_name} {{ get; set; }}")
        
        lines.append("}")
        
        return "\n".join(lines)
    
    def create_json_data(self):
        """Excel 데이터를 JSON으로 변환"""
        # 첫 번째 행: 변수명
        variable_names = self.current_df.iloc[0].values
        
        # 변수명 카운트 (중복 체크) - ~ 로 시작하는 것 제외
        valid_names = [name for name in variable_names if not str(name).startswith('~')]
        name_counts = Counter(valid_names)
        
        # 3행부터 실제 데이터
        data_rows = self.current_df.iloc[2:] if len(self.current_df) > 2 else []
        
        result = []
        
        for row_idx, row in data_rows.iterrows():
            row_data = {}
            processed_names = {}
            
            for col_idx, var_name in enumerate(variable_names):
                var_name = str(var_name).strip()
                
                # ~ 로 시작하는 컬럼 무시
                if var_name.startswith('~'):
                    continue
                
                # 빈 값 체크
                if not var_name or var_name == 'nan':
                    continue
                
                property_name = self.to_pascal_case(var_name)
                cell_value = row.iloc[col_idx]
                
                # NaN 처리
                if pd.isna(cell_value):
                    cell_value = None
                
                # 중복된 변수명 처리
                if name_counts[var_name] > 1:
                    # List로 수집
                    if property_name not in processed_names:
                        row_data[property_name] = []
                        processed_names[property_name] = True
                    row_data[property_name].append(cell_value)
                else:
                    # 일반 값
                    row_data[property_name] = cell_value
            
            result.append(row_data)
        
        return result
    
    def to_pascal_case(self, text):
        """문자열을 PascalCase로 변환"""
        # 공백, 언더스코어, 하이픈을 기준으로 분리
        words = text.replace('_', ' ').replace('-', ' ').split()
        # 각 단어의 첫 글자를 대문자로
        return ''.join(word.capitalize() for word in words)

# 프로그램 실행
if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelReaderApp(root)
    root.mainloop()
# ```

# ## 주요 변경사항

# ✅ **`~`로 시작하는 컬럼 무시**
# - C# 클래스 생성 시 제외
# - JSON 데이터 생성 시 제외
# - 미리보기에서 무시되는 컬럼 표시

# ## Excel 예시
# ```
# | Name   | Age | ~Comment  | Salary | ~Internal |
# | string | int | string    | double | string    |
# | John   | 30  | 테스트     | 50000  | 비밀      |
# | Alice  | 25  | 노트      | 45000  | 메모      |