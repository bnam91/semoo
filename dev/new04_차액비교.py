import pandas as pd

# 엑셀 파일의 '후처리' 시트 읽기
df = pd.read_excel('(전처리)구글시트_2509_20251002_161633.xlsx', sheet_name='후처리')

# 인덱스 1열과 6열을 함께 출력
print("이름과 값:")
for i in range(len(df)):
    name = df.iloc[i, 1]  # 인덱스 1열 (이름)
    value = df.iloc[i, 6]  # 인덱스 6열 (값)
    print(f"{name} {value}")
