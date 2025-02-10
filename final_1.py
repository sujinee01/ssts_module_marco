import pandas as pd

# 파일 경로
file_path = r"C:\Users\uih17186\Documents\TMPS_4inchPF.xlsm"

# 올바른 시트 이름과 헤더 설정
data = pd.read_excel(file_path, sheet_name=2, engine='openpyxl', header=0)

# 열 이름에서 공백 제거
data.columns = data.columns.str.strip()

# 데이터 확인
print("데이터 프레임의 첫 5개 행:")
print(data.head())
print("\n데이터 프레임의 열 이름:")
print(data.columns)

# 결과 저장을 위한 리스트
result = []

# 모듈 선택
selected_modules = input("모듈을 선택하세요 (예: A,B,C): ").split(',')

# 모듈 필터링
for module in selected_modules:
    module = module.strip()  # 공백 제거
    if module in data.columns:
        # 해당 모듈의 SSTS 필터링
        module_ssts = data[data[module] == "x"]["Module name"]
        for ssts in module_ssts:
            result.append({"SSTS": ssts, "Module": module})
    else:
        print(f"모듈 '{module}'은 데이터에 존재하지 않습니다.")

# 결과를 데이터프레임으로 변환
result_df = pd.DataFrame(result).drop_duplicates()

# 결과 정렬 (SSTS와 Module을 기준으로 오름차순 정렬)
result_df = result_df.sort_values(by=["Module", "SSTS"]).reset_index(drop=True)

# 결과 출력
if not result_df.empty:
    print("\n필터링된 결과 (오름차순 정렬):")
    print(result_df)
else:
    print("\n일치하는 데이터가 없습니다.")
