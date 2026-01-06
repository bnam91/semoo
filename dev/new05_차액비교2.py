import pandas as pd
import numpy as np
import os
import glob
import time

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

# 전처리 파일의 '25년3월데이터' 시트에서 진행상품 정보 추출
try:
    df_preprocessing = pd.read_excel(selected_preprocessing_file, sheet_name='25년3월데이터')
    print("\n✅ 전처리 진행상품 정보:")
    print()
    
    # A열(진행상품명), B열(진행자이름), G열(상품가), H열(입금일) 추출
    진행상품_정보 = []
    전처리_총액 = 0
    
    for i in range(len(df_preprocessing)):
        진행상품명 = df_preprocessing.iloc[i, 0] if pd.notna(df_preprocessing.iloc[i, 0]) else ""
        진행자이름 = df_preprocessing.iloc[i, 1] if pd.notna(df_preprocessing.iloc[i, 1]) else ""
        상품가 = df_preprocessing.iloc[i, 6] if pd.notna(df_preprocessing.iloc[i, 6]) else 0
        입금일 = df_preprocessing.iloc[i, 7] if pd.notna(df_preprocessing.iloc[i, 7]) else ""
        
        # 입금일에서 '입금완료_' 제거
        if isinstance(입금일, str) and '입금완료_' in 입금일:
            입금일 = 입금일.replace('입금완료_', '')
        
        # 상품가가 숫자가 아닌 경우 처리
        if not isinstance(상품가, (int, float, np.integer, np.floating)):
            try:
                if isinstance(상품가, str):
                    상품가 = float(상품가.replace(',', '').replace(' ', ''))
                else:
                    상품가 = float(상품가) if 상품가 else 0
            except (ValueError, TypeError):
                상품가 = 0
        
        # 상품가를 정수로 변환하여 .0 제거
        상품가_정수 = int(상품가) if 상품가 == int(상품가) else 상품가
        
        if 진행상품명 and 진행자이름:
            진행상품_정보.append((진행상품명, 진행자이름, 상품가, 입금일))
            전처리_총액 += 상품가
            print(f"{진행자이름}{진행상품명} {상품가_정수}, {입금일}")
    
    print(f"\n💰 전처리 진행상품 총액: {전처리_총액:,.0f}원")
    
    # 진행상품 목록 생성 (중복 제거)
    진행상품_list = list(set([정보[0] for 정보 in 진행상품_정보]))
        
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
    a_value = df.iloc[i, 0]  # A열 (인덱스 0) - 날짜
    c_value = df.iloc[i, 2]  # C열 (인덱스 2)
    d_value = df.iloc[i, 3]  # D열 (인덱스 3)
    
    # A열 날짜를 YYYYMMDD 형식으로 변환
    날짜_문자열 = ""
    if pd.notna(a_value):
        try:
            # pandas datetime으로 변환
            if isinstance(a_value, str):
                날짜_객체 = pd.to_datetime(a_value)
            else:
                날짜_객체 = a_value
            
            # YYMMDD 형식으로 변환 (25년이면 25)
            년도 = 날짜_객체.year % 100  # 2025 -> 25
            월 = 날짜_객체.month
            일 = 날짜_객체.day
            날짜_문자열 = f"{년도:02d}{월:02d}{일:02d}"
        except:
            날짜_문자열 = ""
    
    # C열 값이 진행상품 목록에 포함되는지 확인
    is_진행상품 = False
    if pd.notna(c_value):
        c_str = str(c_value)
        for 진행상품 in 진행상품_list:
            if 진행상품 in c_str:
                is_진행상품 = True
                break
    
    if is_진행상품:
        진행상품_포함.append((c_value, d_value, 날짜_문자열))
    else:
        진행상품_미포함.append((c_value, d_value))

# 진행상품이 포함된 항목 출력
print("✅ 진행상품이 포함된 거래내역:")
print()

# 총액 계산
총액 = 0
for c_value, d_value, 날짜_문자열 in 진행상품_포함:
    print(f"{c_value} {d_value}, {날짜_문자열}")
    
    # D열 값이 숫자인 경우 총액에 더하기
    if pd.notna(d_value) and isinstance(d_value, (int, float, np.integer, np.floating)):
        총액 += float(d_value)
    elif pd.notna(d_value) and isinstance(d_value, str):
        try:
            # 쉼표 제거 후 숫자로 변환
            숫자값 = float(d_value.replace(',', ''))
            총액 += 숫자값
        except (ValueError, AttributeError):
            pass

print(f"\n💰 진행상품 포함 거래내역 총액: {총액:,.0f}원")

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

# 총액 비교
print("\n" + "="*50)
print("📊 총액 비교 결과:")
print(f"전처리 진행상품 총액: {전처리_총액:,.0f}원")
print(f"거래내역 진행상품 총액: {총액:,.0f}원")

차액 = 전처리_총액 - 총액
if 차액 == 0:
    print("✅ 두 총액이 일치합니다!")
else:
    print(f"❌ 두 총액이 다릅니다. 차액: {차액:,.0f}원")
    if 차액 > 0:
        print(f"   → 전처리 총액이 {차액:,.0f}원 더 많습니다.")
    else:
        print(f"   → 거래내역 총액이 {abs(차액):,.0f}원 더 많습니다.")
    
    # 차집합 분석
    print("\n🔍 차집합 분석:")
    
    # 전처리 항목들을 거래내역에서 찾기
    전처리_미매칭 = []
    거래내역_미매칭 = []
    
    # 전처리 항목들을 거래내역에서 찾기
    for 진행상품명, 진행자이름, 상품가, 입금일 in 진행상품_정보:
        매칭됨 = False
        
        for c_value, d_value, 날짜_문자열 in 진행상품_포함:
            # 거래금액 계산
            if pd.notna(d_value) and isinstance(d_value, (int, float, np.integer, np.floating)):
                거래금액 = float(d_value)
            elif pd.notna(d_value) and isinstance(d_value, str):
                try:
                    거래금액 = float(d_value.replace(',', ''))
                except:
                    continue
            else:
                continue
                
            # 상품가, 진행자이름, 진행상품명, 날짜가 모두 일치하는지 확인
            # 공백과 괄호 정규화
            c_value_정규화 = str(c_value).replace(' ', '').replace('（', '(').replace('）', ')')
            진행자이름_정규화 = 진행자이름.replace(' ', '').replace('（', '(').replace('）', ')')
            진행상품명_정규화 = 진행상품명.replace(' ', '').replace('（', '(').replace('）', ')')
            
            if (abs(상품가 - 거래금액) < 0.01 and 
                진행자이름_정규화 in c_value_정규화 and 
                진행상품명_정규화 in c_value_정규화 and 
                입금일 == 날짜_문자열):
                매칭됨 = True
                break
        
        if not 매칭됨:
            전처리_미매칭.append((진행자이름, 진행상품명, 상품가, 입금일))
    
    # 거래내역 항목들을 전처리에서 찾기
    for c_value, d_value, 날짜_문자열 in 진행상품_포함:
        # 거래금액 계산
        if pd.notna(d_value) and isinstance(d_value, (int, float, np.integer, np.floating)):
            거래금액 = float(d_value)
        elif pd.notna(d_value) and isinstance(d_value, str):
            try:
                거래금액 = float(d_value.replace(',', ''))
            except:
                continue
        else:
            continue
            
        매칭됨 = False
        for 진행상품명, 진행자이름, 상품가, 입금일 in 진행상품_정보:
            # 상품가, 진행자이름, 진행상품명, 날짜가 모두 일치하는지 확인
            if (abs(상품가 - 거래금액) < 0.01 and 
                진행자이름 in str(c_value) and 
                진행상품명 in str(c_value) and 
                입금일 == 날짜_문자열):
                매칭됨 = True
                break
        
        if not 매칭됨:
            거래내역_미매칭.append((c_value, 거래금액, 날짜_문자열))
    
    # 결과 출력
    if 전처리_미매칭:
        print(f"\n📋 전처리에만 있는 항목 ({len(전처리_미매칭)}건):")
        for 진행자이름, 진행상품명, 상품가, 입금일 in 전처리_미매칭:
            print(f"   {진행자이름}{진행상품명} {상품가:,.0f}원, {입금일}")
    
    if 거래내역_미매칭:
        print(f"\n📋 거래내역에만 있는 항목 ({len(거래내역_미매칭)}건):")
        for c_value, 거래금액, 날짜_문자열 in 거래내역_미매칭:
            print(f"   {c_value} {거래금액:,.0f}원, {날짜_문자열}")
    
    if not 전처리_미매칭 and not 거래내역_미매칭:
        print("   → 모든 항목이 매칭되었지만 총액이 다릅니다. 중복 또는 부분 매칭 가능성이 있습니다.")
    
    # 상세 매칭 확인
    print("\n🔍 상세 매칭 확인:")
    print("전처리 진행상품 정보의 각 항목을 거래내역에서 찾아보겠습니다.")
    print("자동으로 다음 항목으로 넘어갑니다.\n")
    
    매칭되지_않은_항목들 = []
    
    for i, (진행상품명, 진행자이름, 상품가, 입금일) in enumerate(진행상품_정보, 1):
        print(f"[{i}/{len(진행상품_정보)}] 전처리 항목: {진행자이름}{진행상품명} {상품가:,.0f}원, {입금일}")
        
        # 거래내역에서 매칭되는 항목 찾기
        매칭된_항목들 = []
        for c_value, d_value, 날짜_문자열 in 진행상품_포함:
            # 거래금액 계산
            if pd.notna(d_value) and isinstance(d_value, (int, float, np.integer, np.floating)):
                거래금액 = float(d_value)
            elif pd.notna(d_value) and isinstance(d_value, str):
                try:
                    거래금액 = float(d_value.replace(',', ''))
                except:
                    continue
            else:
                continue
                
            # 상품가, 진행자이름, 진행상품명, 날짜가 모두 일치하는지 확인
            # 공백과 괄호 정규화
            c_value_정규화 = str(c_value).replace(' ', '').replace('（', '(').replace('）', ')')
            진행자이름_정규화 = 진행자이름.replace(' ', '').replace('（', '(').replace('）', ')')
            진행상품명_정규화 = 진행상품명.replace(' ', '').replace('（', '(').replace('）', ')')
            
            if (abs(상품가 - 거래금액) < 0.01 and 
                진행자이름_정규화 in c_value_정규화 and 
                진행상품명_정규화 in c_value_정규화 and 
                입금일 == 날짜_문자열):
                매칭된_항목들.append((c_value, 거래금액, 날짜_문자열))
        
        if 매칭된_항목들:
            print(f"✅ 거래내역에서 {len(매칭된_항목들)}개 항목이 매칭되었습니다:")
            for c_value, 거래금액, 날짜_문자열 in 매칭된_항목들:
                print(f"   - {c_value} {거래금액:,.0f}원, {날짜_문자열}")
        else:
            print("❌ 거래내역에서 매칭되는 항목이 없습니다.")
            매칭되지_않은_항목들.append((진행자이름, 진행상품명, 상품가, 입금일))
        
        time.sleep(0.05)  # 0.05초 대기
        print()
    
    # 매칭되지 않은 거래내역 항목들 찾기
    거래내역_미매칭_항목들 = []
    for c_value, d_value, 날짜_문자열 in 진행상품_포함:
        # 거래금액 계산
        if pd.notna(d_value) and isinstance(d_value, (int, float, np.integer, np.floating)):
            거래금액 = float(d_value)
        elif pd.notna(d_value) and isinstance(d_value, str):
            try:
                거래금액 = float(d_value.replace(',', ''))
            except:
                continue
        else:
            continue
            
        매칭됨 = False
        for 진행상품명, 진행자이름, 상품가, 입금일 in 진행상품_정보:
            # 상품가, 진행자이름, 진행상품명, 날짜가 모두 일치하는지 확인
            # 공백과 괄호 정규화
            c_value_정규화 = str(c_value).replace(' ', '').replace('（', '(').replace('）', ')')
            진행자이름_정규화 = 진행자이름.replace(' ', '').replace('（', '(').replace('）', ')')
            진행상품명_정규화 = 진행상품명.replace(' ', '').replace('（', '(').replace('）', ')')
            
            if (abs(상품가 - 거래금액) < 0.01 and 
                진행자이름_정규화 in c_value_정규화 and 
                진행상품명_정규화 in c_value_정규화 and 
                입금일 == 날짜_문자열):
                매칭됨 = True
                break
        
        if not 매칭됨:
            거래내역_미매칭_항목들.append((c_value, 거래금액, 날짜_문자열))
    
    # 매칭되지 않은 항목들 최종 출력
    if 매칭되지_않은_항목들 or 거래내역_미매칭_항목들:
        if 매칭되지_않은_항목들:
            print("\n" + "="*60)
            print("❌ 매칭되지 않은 전처리 항목들 (입금완료했지만 실제론 미입금 가능성 있음):")
            print("="*60)
            for 진행자이름, 진행상품명, 상품가, 입금일 in 매칭되지_않은_항목들:
                print(f"   {진행자이름}{진행상품명} {상품가:,.0f}원, {입금일}")
            print(f"\n총 {len(매칭되지_않은_항목들)}개 전처리 항목이 매칭되지 않았습니다.")
        
        if 거래내역_미매칭_항목들:
            print("\n" + "="*60)
            print("❌ 매칭되지 않은 거래내역 항목들:")
            print("="*60)
            for c_value, 거래금액, 날짜_문자열 in 거래내역_미매칭_항목들:
                print(f"   {c_value} {거래금액:,.0f}원, {날짜_문자열}")
            print(f"\n총 {len(거래내역_미매칭_항목들)}개 거래내역 항목이 매칭되지 않았습니다.")
    else:
        print("\n✅ 모든 항목이 매칭되었습니다!")
