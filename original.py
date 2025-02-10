import pandas as pd

# 파일 경로 지정
file_path = r"C:\Users\uih17186\Documents\TMPS_4inchPF.xlsm"

# 엑셀 파일 읽기
data = pd.read_excel(file_path, sheet_name=2, engine='openpyxl', header=0)

# 열 이름 정리 (공백 제거)
data.columns = data.columns.str.strip()

# 데이터 확인
print("데이터 프레임의 첫 5개 행:")
print(data.head())
print("\n데이터 프레임의 열 이름:")
print(data.columns)

# 사용자 입력
filter_type = input("필터 유형을 선택하세요 (모듈 -> SSTS: 1, SSTS -> 모듈: 2): ").strip()

# 결과 리스트 초기화
result = []

if filter_type == "1":
    # 모듈 -> SSTS
    selected_modules = input("모듈을 선택하세요 (예: ABGS,ADAS,AFC): ").split(',')
    for module in selected_modules:
        module = module.strip()  # 공백 제거
        if module in data.columns:
            # 해당 모듈과 관련된 처리
            print(f"모듈 {module}과 관련된 데이터를 처리합니다.")
            result.append(data[module])
        else:
            print(f"모듈 {module}은 데이터에 없습니다.")

elif filter_type == "2":
    # SSTS -> 모듈
    selected_ssts = input("SSTS를 선택하세요 (예: SSTS1,SSTS2): ").split(',')
    for ssts in selected_ssts:
        ssts = ssts.strip()  # 공백 제거
        if ssts in data.columns:
            # 해당 SSTS와 관련된 처리
            print(f"SSTS {ssts}와 관련된 데이터를 처리합니다.")
            result.append(data[ssts])
        else:
            print(f"SSTS {ssts}은 데이터에 없습니다.")
else:
    print("올바르지 않은 입력입니다. 프로그램을 종료합니다.")

# 결과 출력
if result:
    print("\n선택된 데이터:")
    print(pd.concat(result, axis=1))
else:
    print("\n결과가 없습니다.")
