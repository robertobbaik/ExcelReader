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
        self.sheet_names = []
        self.current_sheet = None
        
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
        
        # 현재 시트만 변환 버튼
        self.btn_generate = tk.Button(
            top_frame,
            text="현재 시트 변환",
            command=self.generate_files,
            font=("Arial", 12),
            bg="#2196F3",
            fg="white",
            padx=20,
            pady=10,
            state=tk.DISABLED
        )
        self.btn_generate.pack(side=tk.LEFT, padx=5)
        
        # 모든 시트 변환 버튼
        self.btn_generate_all = tk.Button(
            top_frame,
            text="모든 시트 변환",
            command=self.generate_all_sheets,
            font=("Arial", 12),
            bg="#9C27B0",
            fg="white",
            padx=20,
            pady=10,
            state=tk.DISABLED
        )
        self.btn_generate_all.pack(side=tk.LEFT, padx=5)
        
        # 파일 경로 표시
        self.label_path = tk.Label(root, text="선택된 파일: 없음", font=("Arial", 10))
        self.label_path.pack(pady=5)
        
        # 저장 디렉토리 표시
        self.label_output_dir = tk.Label(root, text="저장 폴더: 미지정 (현재 폴더에 저장됨)", font=("Arial", 10), fg="gray")
        self.label_output_dir.pack(pady=5)
        
        # 시트 선택 프레임
        sheet_frame = tk.Frame(root)
        sheet_frame.pack(pady=10)
        
        tk.Label(sheet_frame, text="시트 선택:", font=("Arial", 10)).pack(side=tk.LEFT, padx=5)
        self.combo_sheet = ttk.Combobox(sheet_frame, font=("Arial", 10), width=28, state="readonly")
        self.combo_sheet.pack(side=tk.LEFT, padx=5)
        self.combo_sheet.bind("<<ComboboxSelected>>", self.on_sheet_selected)
        
        # 클래스명 표시 (읽기 전용)
        class_frame = tk.Frame(root)
        class_frame.pack(pady=10)
        
        tk.Label(class_frame, text="클래스명:", font=("Arial", 10)).pack(side=tk.LEFT, padx=5)
        self.label_class_name = tk.Label(class_frame, text="(시트 이름)", font=("Arial", 10, "bold"), fg="blue")
        self.label_class_name.pack(side=tk.LEFT, padx=5)
        
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
            self.load_sheet_names(file_path)
    
    def load_sheet_names(self, file_path):
        """Excel 파일의 시트 이름 목록 로드"""
        try:
            # openpyxl로 시트 이름 가져오기
            workbook = openpyxl.load_workbook(file_path, read_only=True)
            self.sheet_names = workbook.sheetnames
            workbook.close()
            
            # 콤보박스에 시트 이름 설정
            self.combo_sheet['values'] = self.sheet_names
            
            # 첫 번째 시트 자동 선택
            if self.sheet_names:
                self.combo_sheet.current(0)
                self.current_sheet = self.sheet_names[0]
                self.label_class_name.config(text=self.sanitize_class_name(self.current_sheet))
                self.read_excel(file_path, self.current_sheet)
                
                # 버튼 활성화
                self.btn_generate_all.config(state=tk.NORMAL)
            
        except Exception as e:
            messagebox.showerror("오류", f"시트 이름을 불러오는 중 오류가 발생했습니다:\n{str(e)}")
    
    def on_sheet_selected(self, event):
        """시트 선택 시 호출"""
        selected_sheet = self.combo_sheet.get()
        if selected_sheet and self.current_file:
            self.current_sheet = selected_sheet
            self.label_class_name.config(text=self.sanitize_class_name(selected_sheet))
            self.read_excel(self.current_file, selected_sheet)
    
    def select_output_directory(self):
        directory = filedialog.askdirectory(title="저장 폴더 선택")
        
        if directory:
            self.output_directory = directory
            self.label_output_dir.config(text=f"저장 폴더: {directory}", fg="blue")
    
    def read_excel(self, file_path, sheet_name):
        try:
            # Excel 파일의 특정 시트 읽기 (헤더 없이)
            self.current_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            
            # 결과 표시
            self.text_result.delete(1.0, tk.END)
            self.text_result.insert(tk.END, f"=== Excel 데이터 미리보기 (시트: {sheet_name}) ===\n\n")
            
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
            
            messagebox.showinfo("성공", f"시트 '{sheet_name}'를 성공적으로 읽었습니다!\n행 수: {len(self.current_df)}")
            
        except Exception as e:
            messagebox.showerror("오류", f"파일을 읽는 중 오류가 발생했습니다:\n{str(e)}")
    
    def sanitize_class_name(self, sheet_name):
        """시트 이름을 C# 클래스명으로 변환"""
        # 공백, 특수문자 제거 및 PascalCase 변환
        name = sheet_name.strip()
        # 특수문자를 공백으로 변환
        for char in ['-', '_', '.', '(', ')', '[', ']', '{', '}', '!', '@', '#', '$', '%', '^', '&', '*']:
            name = name.replace(char, ' ')
        # PascalCase로 변환
        words = name.split()
        return ''.join(word.capitalize() for word in words if word)
    
    def generate_all_sheets(self):
        """모든 시트를 각각 변환"""
        if not self.current_file:
            messagebox.showwarning("경고", "먼저 Excel 파일을 선택해주세요!")
            return
        
        if not self.sheet_names:
            messagebox.showwarning("경고", "시트가 없습니다!")
            return
        
        # 저장 디렉토리 결정
        if self.output_directory:
            output_dir = self.output_directory
        else:
            output_dir = os.path.dirname(self.current_file) if self.current_file else os.getcwd()
        
        # 결과 표시 초기화
        self.text_result.delete(1.0, tk.END)
        self.text_result.insert(tk.END, "=== 모든 시트 변환 시작 ===\n\n")
        
        success_count = 0
        fail_count = 0
        generated_files = []
        
        # 각 시트를 순회하며 변환
        for sheet_name in self.sheet_names:
            try:
                self.text_result.insert(tk.END, f"처리 중: {sheet_name}...\n")
                self.text_result.update()
                
                # 시트 읽기
                df = pd.read_excel(self.current_file, sheet_name=sheet_name, header=None)
                
                # 최소 2행 체크
                if len(df) < 2:
                    self.text_result.insert(tk.END, f"  ⚠️  건너뜀: 데이터 부족\n\n")
                    fail_count += 1
                    continue
                
                # 클래스명 생성
                class_name = self.sanitize_class_name(sheet_name)
                
                # C# 클래스 코드 생성
                cs_code = self.create_csharp_class_code_from_df(df, class_name)
                
                # JSON 데이터 생성
                json_data = self.create_json_data_from_df(df)
                
                # 파일 저장
                cs_file_path = os.path.join(output_dir, f"{class_name}.cs")
                json_file_path = os.path.join(output_dir, f"{class_name}.json")
                
                with open(cs_file_path, 'w', encoding='utf-8') as f:
                    f.write(cs_code)
                
                with open(json_file_path, 'w', encoding='utf-8') as f:
                    json.dump(json_data, f, indent=2, ensure_ascii=False)
                
                self.text_result.insert(tk.END, f"  ✅ 성공: {class_name}.cs, {class_name}.json\n\n")
                generated_files.append(f"{class_name}.cs")
                generated_files.append(f"{class_name}.json")
                success_count += 1
                
            except Exception as e:
                self.text_result.insert(tk.END, f"  ❌ 실패: {str(e)}\n\n")
                fail_count += 1
        
        # 완료 메시지
        self.text_result.insert(tk.END, "=== 변환 완료 ===\n")
        self.text_result.insert(tk.END, f"성공: {success_count}개 시트\n")
        self.text_result.insert(tk.END, f"실패: {fail_count}개 시트\n")
        self.text_result.insert(tk.END, f"저장 위치: {output_dir}\n\n")
        self.text_result.insert(tk.END, "생성된 파일:\n")
        for file in generated_files:
            self.text_result.insert(tk.END, f"  - {file}\n")
        
        messagebox.showinfo("완료", 
            f"모든 시트 변환 완료!\n\n"
            f"성공: {success_count}개\n"
            f"실패: {fail_count}개\n\n"
            f"저장 위치: {output_dir}")
    
    def generate_files(self):
        """현재 시트만 변환"""
        if self.current_df is None:
            messagebox.showwarning("경고", "먼저 Excel 파일을 선택해주세요!")
            return
        
        if len(self.current_df) < 2:
            messagebox.showwarning("경고", "Excel 파일에 최소 2개의 행(변수명, 자료형)이 필요합니다!")
            return
        
        if not self.current_sheet:
            messagebox.showwarning("경고", "시트가 선택되지 않았습니다!")
            return
        
        # 시트 이름을 클래스명으로 사용
        class_name = self.sanitize_class_name(self.current_sheet)
        
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
        self.text_result.insert(tk.END, f"=== 생성된 C# 클래스 (시트: {self.current_sheet}) ===\n\n")
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
                f"시트 '{self.current_sheet}'에서 파일이 생성되었습니다!\n\n"
                f"C# 클래스: {cs_file_path}\n"
                f"JSON 데이터: {json_file_path}")
            
        except Exception as e:
            messagebox.showerror("오류", f"파일 저장 중 오류가 발생했습니다:\n{str(e)}")
    
    def create_csharp_class_code_from_df(self, df, class_name):
        """DataFrame에서 C# 클래스 코드 생성"""
        # 첫 번째 행: 변수명
        variable_names = df.iloc[0].values
        # 두 번째 행: 자료형
        data_types = df.iloc[1].values
        
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
    
    def create_json_data_from_df(self, df):
        """DataFrame에서 JSON 데이터 생성"""
        # 첫 번째 행: 변수명
        variable_names = df.iloc[0].values
        
        # 변수명 카운트 (중복 체크) - ~ 로 시작하는 것 제외
        valid_names = [name for name in variable_names if not str(name).startswith('~')]
        name_counts = Counter(valid_names)
        
        # 3행부터 실제 데이터
        data_rows = df.iloc[2:] if len(df) > 2 else []
        
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
    
    def create_csharp_class_code(self, class_name):
        """C# 클래스 코드 생성 (현재 DataFrame 사용)"""
        return self.create_csharp_class_code_from_df(self.current_df, class_name)
    
    def create_json_data(self):
        """JSON 데이터 생성 (현재 DataFrame 사용)"""
        return self.create_json_data_from_df(self.current_df)
    
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