# <<누르시오 //// #1번코드 (processed_transactions.xlsx 생성)
import pandas as pd
from collections import defaultdict
import tkinter as tk
from tkinter import filedialog
import re
from datetime import datetime

def process_transaction_data(file_path):
    df = pd.read_excel(file_path)
    results = defaultdict(lambda: {'products': set(), 'total': 0, 'last_date': '', 'category': ''})
    
    for _, row in df.iterrows():
        date = row.get('거래일시', '')
        category = row.get('구분', '')
        description = row.get('적요', '')
        amount = row.get('출금액', 0)
        
        if pd.isna(date) or pd.isna(description) or pd.isna(amount) or pd.isna(category):
            continue
        
        # 적요에서 앞뒤 공백 제거 및 중복 공백 제거
        description = ' '.join(str(description).strip().split())
        
        name, product = extract_name_and_product(description)
        key = f"{name}_{category}"  # 이름과 구분을 조합하여 고유 키 생성
        
        results[key]['products'].add(product)
        results[key]['total'] += amount
        results[key]['last_date'] = max(results[key]['last_date'], date)
        results[key]['category'] = category
    
    output = []
    for key, data in results.items():
        name = key.split('_')[0]  # 키에서 이름 추출
        output.append({
            '날짜': data['last_date'],
            '구분': data['category'],
            '이름': name,
            '총액': data['total'],
            '제품': ', '.join(sorted(data['products']))
        })
    
    return pd.DataFrame(output).sort_values('날짜', ascending=False)

def extract_name_and_product(description):
    # 적요에서 앞뒤 공백 제거 및 중복 공백 제거
    description = ' '.join(description.strip().split())
    
    products = ['원고', '카페원고', '신현빈_원고 5건', '신현빈_원고 10건', '로도프', '로도프 가구매', '_빅뽀로로파워페이지', '왕스프', 
                '몰랑젤리', '밀크티', '바닐라빈', '빅뽀로로(딸기)', '빅뽀로로(밀크)', '빅뽀로로', 
                '유다모','유다모리뷰', '두장군', '컵오트밀(미역국)','비락식혜', '마파두부',
                '허브랜치', '치킨스프', '칠리/머스타드', '허브랜치스콜쳐', '컵오트', '샷시손잡이', '허브랜치', '스콜쳐', '머스타드', '스위트칠리', '따스파족욕기',
                '배송비','목베개_네이버',
                '녹제거제_쿠팡', 
                '마라스톡', '헤로잼', '헛개차', '비락식혜', '멸치육수', '헤로잼(복숭아)', '헤로잼(딸기)',
                ]
    
    # 제품 리스트를 길이 기준 내림차순으로 정렬 (길이가 긴 것부터)
    products = sorted(products, key=len, reverse=True)
    
    # 특수한 경우 처리
    if '당근_촬영' in description:
        return description.split('_')[0], '촬영'
    
    # 제품명 찾기
    found_product = None
    for product in products:
        if product in description:
            found_product = product
            # 제품명을 제외한 나머지 부분이 이름
            name_part = description.replace(product, '')
            break
    
    if found_product:
        # 괄호 안의 이름 처리 (박경민(유지태) 같은 경우)
        bracket_match = re.search(r'\((.*?)\)', name_part)
        if bracket_match:
            name = bracket_match.group(1)  # 괄호 안의 이름 추출
        else:
            name = name_part.strip()  # 괄호 없으면 그대로 사용
            
        return name, found_product
    else:
        # 제품명을 찾지 못한 경우
        parts = description.split()
        if len(parts) > 1:
            return ' '.join(parts[:-1]), parts[-1]
        else:
            return 'Unknown', description

# 파일 선택 대화상자
root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])

if file_path:
    result_df = process_transaction_data(file_path)
    
    # 날짜 형식 통일
    result_df['날짜'] = pd.to_datetime(result_df['날짜']).dt.strftime('%Y-%m-%d %H:%M:%S')
    
    # 열 순서 조정
    result_df = result_df[['날짜', '구분', '이름', '총액', '제품']]
    
    # 결과 출력 형식 조정
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    pd.set_option('display.max_colwidth', None)
    print(result_df.to_string(index=False))
    
    output_file = 'processed_transactions.xlsx'
    result_df.to_excel(output_file, index=False)
    print(f"\n결과가 {output_file}에 저장되었습니다.")
else:
    print("파일이 선택되지 않았습니다.")