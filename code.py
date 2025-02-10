import pandas as pd


st.title("SSTS-Module Filter Application")

# 엑셀 파일 업로드
uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요:", type=["xlsm", "xlsx"])

if uploaded_file:
    # 엑셀 파일 읽기
    sheet_names = pd.ExcelFile(uploaded_file, engine="openpyxl").sheet_names
    sheet_name = st.selectbox("시트를 선택하세요:", options=sheet_names)
    
    # 데이터 읽기
    data = pd.read_excel(uploaded_file, sheet_name=sheet_name, engine="openpyxl", header=0)
    data.columns = data.columns.str.strip()  # 열 이름에서 공백 제거

    st.write("업로드된 데이터:")
    st.dataframe(data.head())  # 데이터 미리 보기

    # 필터 유형 선택
    filter_type = st.selectbox("필터 유형을 선택하세요", options=["모듈 -> SSTS", "SSTS -> 모듈"])

    result = []

    if filter_type == "모듈 -> SSTS":
        # 모듈 선택
        modules = st.multiselect("모듈을 선택하세요:", options=data.columns[2:])  # 첫 2열은 제외
        if modules:
            for module in modules:
                if module in data.columns:
                    module_ssts = data[data[module] == "x"]["Module name"]
                    for ssts in module_ssts:
                        result.append({"SSTS": ssts, "Module": module})
            # 결과 출력
            result_df = pd.DataFrame(result).drop_duplicates()
            result_df = result_df.sort_values(by=["Module", "SSTS"]).reset_index(drop=True)
            st.write("필터링된 결과:")
            st.dataframe(result_df)

    elif filter_type == "SSTS -> 모듈":
        # SSTS 선택
        ssts_list = st.multiselect("SSTS를 선택하세요:", options=data["Module name"].dropna())
        if ssts_list:
            for ssts in ssts_list:
                if ssts in data["Module name"].values:
                    ssts_row = data[data["Module name"] == ssts]
                    for module in data.columns[2:]:  # 실제 모듈 열만 탐색
                        if ssts_row[module].values[0] == "x":
                            result.append({"SSTS": ssts, "Module": module})
            # 결과 출력
            result_df = pd.DataFrame(result).drop_duplicates()
            result_df = result_df.sort_values(by=["Module", "SSTS"]).reset_index(drop=True)
            st.write("필터링된 결과:")
            st.dataframe(result_df)

        

