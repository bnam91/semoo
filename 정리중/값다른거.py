import pandas as pd

# 엑셀 파일의 "후처리" 시트 읽기
file_path = '구글시트_2503_20250408_210743.xlsx'
df = pd.read_excel(file_path, sheet_name="후처리")

# 데이터프레임의 열 이름 확인
print("데이터프레임 열 이름:", df.columns.tolist())
print("데이터프레임 크기:", df.shape)

# G열과 J열의 값이 다른 행 찾기
different_rows = []

# 엑셀의 G열과 J열에 해당하는 인덱스 찾기 (A열이 인덱스 0이면 G열은 6, J열은 9)
g_index = 6  # G열 (0부터 시작, A=0, B=1, ..., G=6)
j_index = 9  # J열 (0부터 시작, A=0, B=1, ..., J=9)

# 열 인덱스가 데이터프레임의 범위를 벗어나지 않는지 확인
if len(df.columns) > max(g_index, j_index):
    for idx, row in df.iterrows():
        # iloc으로 직접 위치 인덱스 접근
        g_value = row.iloc[g_index]  # G열
        j_value = row.iloc[j_index]  # J열
        
        if pd.notna(g_value) and pd.notna(j_value) and g_value != j_value:
            # 실제 엑셀 행 번호는 인덱스 + 2 (헤더 행과 0부터 시작하는 인덱스 고려)
            different_rows.append(idx + 2)
else:
    print(f"오류: 데이터프레임에 G열(인덱스 {g_index}) 또는 J열(인덱스 {j_index})이 없습니다.")
    print(f"데이터프레임에는 {len(df.columns)}개의 열만 있습니다.")

# 결과 출력
print("G열과 J열의 값이 다른 행 번호:")
for row_num in different_rows:
    print(f"행 번호: {row_num}")

print(f"총 {len(different_rows)}개의 행이 발견되었습니다.")