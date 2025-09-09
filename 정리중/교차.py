import pandas as pd

# Excel 파일 로드 (첫 행을 헤더로 명시적 지정)
file_path = '거래내역조회_20250408.xls'
sheet_name = 'Sheet1 (2)'
df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)

# 데이터 구조 확인
print("데이터 열 이름:", df.columns.tolist())

# 3번째와 4번째 열 선택 (인덱스 2, 3)
column_c = df.columns[2]  # 3번째 열 이름 (적요)
column_d = df.columns[3]  # 4번째 열 이름 (출금액)

# 적요별 출금액 합계 계산
grouped_data = df.groupby(column_c)[column_d].sum().reset_index()
grouped_data.columns = ['적요', '출금액 합계']

# 결과 출력
print(grouped_data)

# 결과를 새 Excel 파일로 저장
output_file = '적요별_출금액_합계.xlsx'
grouped_data.to_excel(output_file, index=False)
print(f"'{output_file}' 파일로 결과가 저장되었습니다.")
