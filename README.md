# Excel to C# Class & JSON Converter

Excel 파일을 C# 클래스와 JSON 데이터로 자동 변환하는 프로그램입니다.

## 📋 주요 기능

- ✅ Excel 파일(.xlsx, .xls)을 C# 클래스 파일(.cs)로 변환
- ✅ Excel 데이터를 JSON 파일(.json)로 자동 생성
- ✅ 단일 시트, 전체 시트, 폴더 내 모든 파일 일괄 변환 지원
- ✅ 시트 이름이 자동으로 클래스명으로 설정
- ✅ 중복된 컬럼명은 자동으로 `List<T>` 타입으로 변환
- ✅ `~`로 시작하는 컬럼은 자동으로 무시
- ✅ Windows와 macOS 모두 지원

---

## 🚀 실행 방법

### 방법 1: 실행파일 사용 (권장)

1. [Releases](https://github.com/your-repo/ExcelReader/releases)에서 운영체제에 맞는 실행파일 다운로드
   - **Windows**: `ExcelConverter.exe`
   - **macOS**: `ExcelConverter-macOS.zip` (압축 해제 후 실행)

2. 실행파일을 더블클릭하여 실행

**macOS 보안 경고 해결:**
```bash
# 터미널에서 실행 (보안 경고 시)
xattr -cr ExcelConverter.app
```

### 방법 2: Python으로 직접 실행

```bash
# 1. 필요한 라이브러리 설치
pip3 install openpyxl pandas

# 2. 프로그램 실행
python3 excel_reader.py
```

---

## 📝 Excel 파일 형식

### 필수 구조

Excel 파일은 다음과 같은 형식이어야 합니다:

| 1행 (변수명) | Name   | Age | Salary | Email          |
|-------------|--------|-----|--------|----------------|
| 2행 (자료형) | string | int | double | string         |
| 3행 (데이터) | John   | 30  | 50000  | john@test.com  |
| 4행 (데이터) | Alice  | 25  | 45000  | alice@test.com |

**규칙:**
- **1행**: 변수명 (C# 속성명으로 변환됨)
- **2행**: 자료형 (int, string, double, bool, DateTime 등)
- **3행 이후**: 실제 데이터

### 예시

**입력 (Excel):**
```
| Id  | Name   | Level |
| int | string | int   |
| 1   | Sword  | 5     |
| 2   | Shield | 3     |
```

**출력 (C# 클래스):**
```csharp
using System;
using System.Collections.Generic;

public class Item
{
    public int Id { get; set; }
    public string Name { get; set; }
    public int Level { get; set; }
}
```

**출력 (JSON):**
```json
[
  {
    "Id": 1,
    "Name": "Sword",
    "Level": 5
  },
  {
    "Id": 2,
    "Name": "Shield",
    "Level": 3
  }
]
```

---

## 🎯 사용 방법

### 1. 단일 파일의 특정 시트 변환

1. **"Excel 파일 선택"** 버튼 클릭
2. 변환할 Excel 파일 선택
3. 시트 선택 드롭다운에서 원하는 시트 선택
4. **"현재 시트 변환"** 버튼 클릭

### 2. 단일 파일의 모든 시트 변환

1. **"Excel 파일 선택"** 버튼 클릭
2. 변환할 Excel 파일 선택
3. **"모든 시트 변환"** 버튼 클릭

### 3. 폴더 내 모든 Excel 파일 일괄 변환 (추천!)

1. **"📁 폴더 전체 변환"** 버튼 클릭
2. Excel 파일들이 있는 폴더 선택
3. 자동으로 모든 `.xlsx`, `.xls` 파일의 모든 시트가 변환됨

### 4. 저장 폴더 지정 (선택사항)

1. **"저장 폴더 선택"** 버튼 클릭
2. 원하는 저장 위치 선택
3. 미지정 시 Excel 파일과 같은 폴더에 저장됨

---

## ⚙️ 특수 기능

### 1. 컬럼 무시하기

컬럼명이 `~`로 시작하면 해당 컬럼은 무시됩니다.

**예시:**
```
| Name   | Age | ~Comment  | Salary |
| string | int | string    | double |
| John   | 30  | 테스트     | 50000  |
```

→ `~Comment` 컬럼은 C# 클래스와 JSON에서 제외됨

### 2. 중복 컬럼명 처리

같은 이름의 컬럼이 여러 개 있으면 자동으로 `List<T>`로 변환됩니다.

**예시:**
```
| Name   | Age | Name   |
| string | int | string |
| John   | 30  | Jane   |
```

**생성되는 C# 클래스:**
```csharp
public class MyClass
{
    public List<string> Name { get; set; }
    public int Age { get; set; }
}
```

**생성되는 JSON:**
```json
[
  {
    "Name": ["John", "Jane"],
    "Age": 30
  }
]
```

### 3. 클래스명 자동 설정

- 시트 이름이 자동으로 클래스명이 됩니다
- 특수문자는 제거되고 PascalCase로 변환됩니다

**예시:**
- 시트명: `Employee Data` → 클래스명: `EmployeeData`
- 시트명: `item-list` → 클래스명: `ItemList`
- 시트명: `quest_info` → 클래스명: `QuestInfo`

---

## ⚠️ 주의사항

### 필수 사항

1. **Excel 파일은 반드시 최소 2행 이상**이어야 합니다
   - 1행: 변수명
   - 2행: 자료형
   - 3행 이후: 데이터

2. **변수명(1행)은 비어있으면 안 됩니다**

3. **자료형(2행)은 유효한 C# 타입**이어야 합니다
   - 권장: `int`, `string`, `double`, `bool`, `DateTime`, `float`, `long`

### 권장 사항

1. **시트 이름을 명확하게 작성**하세요
   - 시트 이름이 클래스명이 됩니다
   - 예: `UserData`, `ItemInfo`, `QuestList`

2. **변수명은 영어로 작성**하는 것을 권장합니다
   - 한글도 가능하지만 C#에서 사용하기 불편할 수 있습니다

3. **임시 컬럼은 `~`로 시작**하세요
   - 개발 중 메모나 임시 데이터는 `~`를 붙이면 무시됩니다

4. **대용량 파일 처리 시**
   - 폴더 전체 변환 시 시간이 걸릴 수 있습니다
   - 진행 상황이 화면에 표시되니 기다려주세요

### 에러 발생 시

**"데이터 부족" 오류:**
- Excel 파일에 최소 2행(변수명, 자료형)이 있는지 확인하세요

**"파일 읽기 실패" 오류:**
- Excel 파일이 다른 프로그램에서 열려있지 않은지 확인하세요
- 파일 형식이 `.xlsx` 또는 `.xls`인지 확인하세요

**"시트가 없습니다" 오류:**
- Excel 파일에 시트가 1개 이상 있는지 확인하세요

---

## 📂 출력 파일

### 생성되는 파일

각 시트당 2개의 파일이 생성됩니다:

1. **`{클래스명}.cs`**: C# 클래스 파일
2. **`{클래스명}.json`**: JSON 데이터 파일

### 예시

**입력:**
- `GameData.xlsx` 파일
  - `Item` 시트
  - `Quest` 시트
  - `Enemy` 시트

**출력:**
```
📁 저장 폴더/
  ├── Item.cs
  ├── Item.json
  ├── Quest.cs
  ├── Quest.json
  ├── Enemy.cs
  └── Enemy.json
```

---

## 🔧 개발자 정보

### 시스템 요구사항

- **Windows**: Windows 10 이상
- **macOS**: macOS 10.13 이상
- **Python** (소스코드 실행 시): Python 3.9 이상

### 의존성 라이브러리

```
openpyxl>=3.0.0
pandas>=1.3.0
```

### 빌드 방법

```bash
# PyInstaller 설치
pip install pyinstaller

# Windows용 빌드
pyinstaller --onefile --windowed --name "ExcelConverter" --clean -y excel_reader.py

# macOS용 빌드
pyinstaller --onedir --windowed --name "ExcelConverter" --clean -y excel_reader.py
```

---

## 📄 라이선스

MIT License

---

## 🙋 FAQ

**Q: 한 번에 몇 개의 파일을 처리할 수 있나요?**  
A: 제한 없습니다. 폴더 전체 변환 기능으로 수백 개의 파일도 한 번에 처리할 수 있습니다.

**Q: Excel 파일을 열어둔 상태에서 변환할 수 있나요?**  
A: 아니요. Excel 파일은 닫혀 있어야 합니다.

**Q: 생성된 C# 파일을 바로 Unity에서 사용할 수 있나요?**  
A: 네, 생성된 `.cs` 파일을 Unity 프로젝트의 Scripts 폴더에 복사하면 바로 사용할 수 있습니다.

**Q: JSON 파일의 인코딩은 무엇인가요?**  
A: UTF-8로 저장되며, 한글도 정상적으로 표시됩니다.

**Q: 맥에서 "확인되지 않은 개발자" 경고가 뜨면?**  
A: 시스템 환경설정 → 보안 및 개인 정보 보호 → "확인 없이 열기" 클릭하거나, 터미널에서 `xattr -cr ExcelConverter.app` 실행

**Q: 윈도우와 맥에서 만든 파일을 서로 호환할 수 있나요?**  
A: 네, 생성된 `.cs`와 `.json` 파일은 모든 운영체제에서 동일하게 사용할 수 있습니다.

---

## 📞 문의 및 버그 리포트

이슈가 있거나 개선 제안이 있으시면 [GitHub Issues](https://github.com/your-repo/ExcelReader/issues)에 등록해주세요.

---

**Happy Coding! 🎉**
