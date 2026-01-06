import pandas as pd
from googleapiclient.discovery import build
import re
from auth import get_credentials
from datetime import datetime
import os
from openpyxl import load_workbook

def main():
    # 사용자로부터 년월 입력 받기 (예: 2504, 2512)
    user_input = input("년월을 입력하세요 (예: 2504, 2512): ")
    payment_pattern = f'입금완료_{user_input}'
    
    # 구글 인증 자격 증명 가져오기
    creds = get_credentials()
    
    # 구글 시트 API 클라이언트 생성
    sheets_service = build('sheets', 'v4', credentials=creds)
    
    # 스프레드시트 ID
    SPREADSHEET_ID = '1CK2UXTy7HKjBe2T0ovm5hfzAAKZxZAR_ev3cbTPOMPs'
    
    # 스프레드시트의 모든 시트 정보 가져오기
    spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheets = spreadsheet.get('sheets', [])
    
    # 모든 시트의 필터링된 데이터를 저장할 리스트
    all_filtered_dfs = []
    total_items = 0
    
    # 각 시트 순회
    for sheet in sheets:
        sheet_title = sheet.get('properties', {}).get('title', '')
        
        # '(가라)'가 포함된 시트는 건너뛰기
        if '(가라)' in sheet_title:
            continue
        
        print(f'{sheet_title} 시트 처리 중...')
        
        # 시트 데이터 범위 설정 (9행부터 모든 데이터)
        RANGE_NAME = f'{sheet_title}!A9:Z'
        
        try:
            # 시트 데이터 가져오기
            result = sheets_service.spreadsheets().values().get(
                spreadsheetId=SPREADSHEET_ID,
                range=RANGE_NAME
            ).execute()
            
            # 데이터 처리
            rows = result.get('values', [])
            if not rows:
                print(f'{sheet_title} 시트에 데이터가 없습니다.')
                continue
            
            # 데이터프레임으로 변환
            df = pd.DataFrame(rows)
            
            # 충분한 열이 없는 행 처리 (P열이 없는 경우)
            df = df.dropna(axis=0, how='all')  # 완전히 빈 행 제거
            
            # B열(인덱스 1)에서 괄호 안의 텍스트만 추출
            for i, row in enumerate(df.values):
                if len(row) > 1:  # B열이 있는지 확인
                    if pd.notna(row[1]):  # null 체크
                        # 괄호 패턴 찾기 (예: '조수갑(유지현)' -> '유지현')
                        match = re.search(r'.*\((.*)\).*', str(row[1]))
                        if match:
                            row[1] = match.group(1).strip()  # 괄호 안의 내용으로 대체
            
            # 특정 열의 공백 및 특수문자 제거
            for i, row in enumerate(df.values):
                # 인덱스 5(F열)에서 괄호 안의 텍스트만 추출
                if len(row) > 5:  # F열이 있는지 확인
                    if pd.notna(row[5]):  # null 체크
                        # 괄호 패턴 찾기
                        match = re.search(r'.*\((.*)\).*', str(row[5]))
                        if match:
                            row[5] = match.group(1).strip()  # 괄호 안의 내용으로 대체
                
                # 인덱스 5, 6, 9의 공백 제거
                for idx in [5, 6, 9]:
                    if len(row) > idx:
                        if pd.notna(row[idx]):  # null 체크
                            row[idx] = str(row[idx]).replace(' ', '')
                
                # 인덱스 6, 8, 9의 '.', '-' 제거
                for idx in [6, 8, 9]:
                    if len(row) > idx:
                        if pd.notna(row[idx]):  # null 체크
                            row[idx] = str(row[idx]).replace('.', '').replace('-', '')
                
                # 인덱스 9(F열)의 숫자만 남기기
                if len(row) > 9:
                    if pd.notna(row[9]):  # null 체크
                        # 숫자만 남기기 위해 정규식 사용
                        row[9] = ''.join(re.findall(r'\d+', str(row[9])))
            
            # P열 (인덱스 15) 필터링 - 사용자 입력에 따른 데이터만 선택
            filtered_rows = []
            for i, row in enumerate(df.values):
                if len(row) > 15:  # P열이 있는지 확인
                    p_value = str(row[15])
                    
                    # F열 값 확인 (인덱스 5)
                    f_value = "" if len(row) <= 5 else str(row[5]).strip()
                    
                    # 사용자 입력에 따른 패턴 확인 및 F열이 비어있지 않은지 확인
                    if re.search(f'{payment_pattern}\\d{{2}}', p_value) and f_value:
                        filtered_rows.append(row)
            
            # 필터링된 데이터로 새 데이터프레임 생성
            filtered_df = pd.DataFrame(filtered_rows)
            
            if not filtered_df.empty:
                print(f'{sheet_title} 시트에서 {len(filtered_df)}개 항목 추출')
                
                # 리스트에 데이터프레임 추가
                all_filtered_dfs.append(filtered_df)
                total_items += len(filtered_df)
            else:
                print(f'{sheet_title} 시트에서 {user_input} 데이터가 없습니다.')
                
        except Exception as e:
            print(f'{sheet_title} 시트 처리 중 오류 발생: {e}')
    
    if not all_filtered_dfs:
        print(f'모든 시트에서 {user_input}에 해당하는 데이터가 없습니다.')
        return
    
    # 모든 데이터프레임을 하나로 합치기
    combined_df = pd.concat(all_filtered_dfs, ignore_index=True)
    
    # 필요한 열만 선택 (A~D열(0~3), L~O열(11~14) 제외)
    # 데이터프레임에 충분한 열이 있는지 확인
    if combined_df.shape[1] >= 16:  # P열(인덱스 15)까지는 최소한 있어야 함
        # 열 선택
        needed_columns = list(range(4, 11)) + list(range(15, combined_df.shape[1]))
        selected_df = combined_df.iloc[:, needed_columns]
    else:
        # 열이 부족한 경우 원본 데이터프레임 사용
        selected_df = combined_df
        print('경고: 일부 필요한 열이 없어 모든 열을 유지합니다.')
    
    # 결과를 엑셀 파일로 저장
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = f'(전처리)구글시트_{user_input}_{timestamp}.xlsx'
    
    # 인덱스 5(엑셀의 F열)를 문자열로 변환하여 숫자 형식이 아닌 텍스트로 저장되도록 함
    if selected_df.shape[1] > 5:  # F열이 있는지 확인
        selected_df.iloc[:, 5] = selected_df.iloc[:, 5].astype(str)
    
    # 하나의 시트에 모든 데이터 저장
    selected_df.to_excel(output_file, sheet_name='25년3월데이터', index=False)
    
    # F열을 텍스트 형식으로 변환 (엑셀 파일 직접 수정)
    try:
        # 엑셀 파일 로드
        wb = load_workbook(output_file)
        ws = wb['25년3월데이터']
        
        # F열 인덱스 (엑셀에서는 F열은 6번째 열)
        f_col_idx = 6
        
        # 첫 번째 행은 헤더이므로 2번째 행부터 처리
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=f_col_idx)
            # 셀 형식을 텍스트로 설정
            cell.number_format = '@'
        
        # 변경사항 저장
        wb.save(output_file)
        print(f'F열을 텍스트 형식으로 저장했습니다.')
    except Exception as e:
        print(f'F열 텍스트 형식 변환 중 오류 발생: {e}')
    
    print(f'데이터가 {output_file} 파일로 저장되었습니다.')
    print(f'총 {total_items}개의 항목이 추출되었습니다.')

if __name__ == '__main__':
    main()
