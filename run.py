Python 3.12.6 (tags/v3.12.6:a4a2d2b, Sep  6 2024, 20:11:23) [MSC v.1940 64 bit (AMD64)] on win32
Type "help", "copyright", "credits" or "license()" for more information.
>>> import pandas as pd
... 
... # 파일 경로 지정
... file_path = r"C:\Users\uih17186\Documents\TMPS_4inchPF.xlsm"
... 
... # 엑셀 파일 읽기
... data = pd.read_excel(file_path, sheet_name=2, engine='openpyxl', header=0)
... 
... # 열 이름 정리 (공백 제거)
... data.columns = data.columns.str.strip()
... 
... # 데이터 확인
... print("데이터 프레임의 첫 5개 행:")
... print(data.head())
... print("\n데이터 프레임의 열 이름:")
... print(data.columns)
... 
... # 사용자 입력
... filter_type = input("필터 유형을 선택하세요 (모듈 -> SSTS: 1, SSTS -> 모듈: 2): ")
... 
... # 결과 리스트 초기화
... result = []
... 
... if filter_type == "1":
...     # 모듈 -> SSTS
...     selected_modules = input("모듈을 선택하세요 (예: ABGS,ADAS,AFC): ").split(',')
... 
...     for module in selected_modules:
...         module = module.strip()  # 공백 제거
...         if module in data.columns:
...             # 해당 모듈과 관련
