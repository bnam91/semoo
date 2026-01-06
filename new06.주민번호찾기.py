from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from auth import get_credentials
import time
import re

SPREADSHEET_ID = '1CK2UXTy7HKjBe2T0ovm5hfzAAKZxZAR_ev3cbTPOMPs'

def is_exact_match(search_name, cell_str):
    """
    정확한 이름 매칭 함수
    - 정확히 일치하는 경우: '정민' → '정민' (O)
    - 괄호 안의 이름: '배진수(정민)' → '정민' (O)
    - 부분 일치: '정민지' → '정민' (X)
    """
    # 1. 정확히 일치하는 경우
    if search_name == cell_str:
        return True
    
    # 2. 괄호 안에 있는 경우 (예: '배진수(정민)')
    if f'({search_name})' in cell_str:
        return True
    
    # 3. 단어 경계를 고려한 정확한 매칭 (공백이나 특수문자로 구분된 경우)
    # 예: '정민 ' 또는 ' 정민' 또는 '정민,' 등
    pattern = r'(^|[^\w])' + re.escape(search_name) + r'($|[^\w])'
    if re.search(pattern, cell_str):
        return True
    
    return False

def find_persons(search_names):
    try:
        creds = get_credentials()
        service = build('sheets', 'v4', credentials=creds)

        sheet_metadata = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
        sheets = sheet_metadata.get('sheets', '')

        sheet_map = {sheet['properties']['title']: sheet['properties']['sheetId'] for sheet in sheets}

        # 각 이름별로 결과를 저장할 딕셔너리
        all_results = {name: [] for name in search_names}
        
        for sheet_name, sheet_id in sheet_map.items():
            # 모든 시트에서 검색
            print(f"검색 중: {sheet_name}")
            
            # 각 시트에서 모든 이름 검색
            sheet_results = search_in_sheet(service, sheet_name, search_names)
            
            # 결과를 각 이름별로 분류
            for name, results in sheet_results.items():
                if results:
                    all_results[name].extend(results)
                    print(f"  🚩 '{name}' 발견! ({len(results)}개)")
            
            # API 제한을 피하기 위해 1초 대기
            time.sleep(1)

        # 최종 결과 출력
        print(f"\n🚩 검색 결과")
        for name in search_names:
            if all_results[name]:
                print(f"\n{name} 발견결과:")
                for i, result in enumerate(all_results[name], 1):
                    print(f"{i}. {result['sheet_name']} / {result['i_value']} / {result['j_value']}")
            else:
                print(f"\n{name}: 찾을 수 없습니다.")

    except HttpError as error:
        print(f'오류가 발생했습니다: {error}')

def search_in_sheet(service, sheet_name, search_names):
    try:
        range_name = f'{sheet_name}!A1:P1000'  # P열까지만 범위 지정
        
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range=range_name).execute()
        values = result.get('values', [])

        if not values:
            return {name: [] for name in search_names}

        # 각 이름별로 결과를 저장할 딕셔너리
        sheet_results = {name: [] for name in search_names}

        # 모든 셀에서 각 이름 검색
        for row_index, row in enumerate(values, start=1):
            for col_index, cell_value in enumerate(row):
                cell_str = str(cell_value)
                for search_name in search_names:
                    # 정확한 이름 매칭 (괄호 안의 이름도 허용)
                    if is_exact_match(search_name, cell_str):
                        # I열과 J열 값 저장
                        i_value = row[8] if len(row) > 8 else ''
                        j_value = row[9] if len(row) > 9 else ''
                        location = f"{chr(65 + col_index)}{row_index}"
                        
                        # J열 값 검증 (하이픈, 공백 제외 13자리 확인)
                        j_value_clean = str(j_value).replace('-', '').replace(' ', '')
                        if len(j_value_clean) != 13:
                            j_value = f"{j_value} (오류)"
                        
                        # 결과 저장
                        sheet_results[search_name].append({
                            'sheet_name': sheet_name,
                            'location': location,
                            'i_value': i_value,
                            'j_value': j_value
                        })
        
        return sheet_results
        
    except HttpError as error:
        if error.resp.status == 429:  # Rate limit exceeded
            print(f'{sheet_name} 시트 검색 중 API 제한 도달. 30초 대기 후 재시도...')
            time.sleep(30)
            return search_in_sheet(service, sheet_name, search_names)  # 재시도
        else:
            print(f'{sheet_name} 시트 검색 중 오류: {error}')
            return False

if __name__ == '__main__':
    search_input = input("검색할 이름을 입력하세요 (쉼표로 구분): ")
    search_names = [name.strip() for name in search_input.split(',')]
    find_persons(search_names)
