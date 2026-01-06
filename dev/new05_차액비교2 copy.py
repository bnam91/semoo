import pandas as pd
import os
import glob

# 루트 디렉토리에서 '거래내역조회_'가 포함된 파일 찾기
root_dir = '.'
pattern = os.path.join(root_dir, '*거래내역조회_*.xls*')
matching_files = glob.glob(pattern)

if not matching_files:
    print("'거래내역조회_'가 포함된 파일을 찾을 수 없습니다.")
    exit()

# 파일이 여러 개인 경우 선택할 수 있게 함
if len(matching_files) == 1:
    selected_file = matching_files[0]
    print(f"자동 선택된 파일: {os.path.basename(selected_file)}")
else:
    print("📁 하나은행 거래내역 조회 파일을 선택하세요: ")
    print()
    for i, file in enumerate(matching_files, 1):
        print(f"{i}. {os.path.basename(file)}")
    print()
    
    while True:
        try:
            choice = int(input("번호를 입력하세요: ")) - 1
            if 0 <= choice < len(matching_files):
                selected_file = matching_files[choice]
                break
            else:
                print("올바른 번호를 입력하세요.")
        except ValueError:
            print("숫자를 입력하세요.")

print(f"선택된 파일: {os.path.basename(selected_file)}")
print()

# '(전처리)구글시트_'가 포함된 파일 찾기
preprocessing_pattern = os.path.join(root_dir, '*(전처리)구글시트_*.xlsx')
preprocessing_files = glob.glob(preprocessing_pattern)

if not preprocessing_files:
    print("'(전처리)구글시트_'가 포함된 파일을 찾을 수 없습니다.")
    exit()

# 전처리 파일이 여러 개인 경우 선택할 수 있게 함
if len(preprocessing_files) == 1:
    selected_preprocessing_file = preprocessing_files[0]
    print(f"자동 선택된 전처리 파일: {os.path.basename(selected_preprocessing_file)}")
else:
    print("📊 다음 전처리 파일들 중에서 선택하세요: ")
    print()
    for i, file in enumerate(preprocessing_files, 1):
        print(f"{i}. {os.path.basename(file)}")
    print()
    
    while True:
        try:
            choice = int(input("전처리 파일 번호를 입력하세요: ")) - 1
            if 0 <= choice < len(preprocessing_files):
                selected_preprocessing_file = preprocessing_files[choice]
                break
            else:
                print("올바른 번호를 입력하세요.")
        except ValueError:
            print("숫자를 입력하세요.")

print(f"선택된 전처리 파일: {os.path.basename(selected_preprocessing_file)}")
print()

# 전처리 파일의 '후처리' 시트 A열에서 진행상품 추출 (중복 제거)
try:
    df_preprocessing = pd.read_excel(selected_preprocessing_file, sheet_name='후처리')
    print("\n진행상품 목록:")
    
    # A열의 모든 값들을 가져와서 쉼표로 분리
    진행상품_set = set()
    for value in df_preprocessing.iloc[:, 0].dropna():
        if pd.notna(value):
            # 쉼표로 분리하고 공백 제거
            items = [item.strip() for item in str(value).split(',')]
            진행상품_set.update(items)
    
    # 정렬해서 출력
    진행상품_list = []
    for item in sorted(진행상품_set):
        if item:  # 빈 문자열이 아닌 경우만
            진행상품_list.append(item)
            print(f"{item}")
        
except FileNotFoundError:
    print(f"\n{selected_preprocessing_file} 파일을 찾을 수 없습니다.")
    진행상품_list = []
except Exception as e:
    print(f"\n전처리 파일 읽기 오류: {e}")
    진행상품_list = []

print("\n" + "="*50)
print()

# 엑셀 파일의 'Sheet1' 시트 읽기
df = pd.read_excel(selected_file, sheet_name='Sheet1')

# 진행상품이 포함된 항목과 포함되지 않은 항목 분리
진행상품_포함 = []
진행상품_미포함 = []

for i in range(len(df)):
    c_value = df.iloc[i, 2]  # C열 (인덱스 2)
    d_value = df.iloc[i, 3]  # D열 (인덱스 3)
    
    # C열 값이 진행상품 목록에 포함되는지 확인
    is_진행상품 = False
    if pd.notna(c_value):
        c_str = str(c_value)
        for 진행상품 in 진행상품_list:
            if 진행상품 in c_str:
                is_진행상품 = True
                break
    
    if is_진행상품:
        진행상품_포함.append((c_value, d_value))
    else:
        진행상품_미포함.append((c_value, d_value))

# 진행상품이 포함된 항목 출력
print("✅ 진행상품이 포함된 거래내역:")
print()
for c_value, d_value in 진행상품_포함:
    print(f"{c_value} {d_value}")

# 진행상품이 포함되지 않은 항목을 txt 파일로 저장 (덮어쓰기)
output_file = "진행상품_미포함_거래내역.txt"
with open(output_file, 'w', encoding='utf-8') as f:
    f.write("진행상품이 포함되지 않은 거래내역:\n")
    f.write("="*50 + "\n")
    for c_value, d_value in 진행상품_미포함:
        f.write(f"{c_value} {d_value}\n")

print(f"\n진행상품 미포함 거래내역이 '{output_file}' 파일로 저장되었습니다.")
print(f"진행상품 포함: {len(진행상품_포함)}건")
print(f"진행상품 미포함: {len(진행상품_미포함)}건")
