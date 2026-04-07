#!/usr/bin/env python3
"""
가상환경 자동 감지 및 활성화
"""
import sys
import os
from pathlib import Path

# 현재 스크립트의 디렉토리
SCRIPT_DIR = Path(__file__).parent.absolute()
VENV_PYTHON = SCRIPT_DIR / 'venv' / 'bin' / 'python3'

# 가상환경이 존재하고 현재 Python이 가상환경이 아닌 경우
if VENV_PYTHON.exists() and 'venv' not in sys.executable:
    # 가상환경의 Python으로 재실행
    os.execv(str(VENV_PYTHON), [str(VENV_PYTHON)] + sys.argv)

import pandas as pd
from googleapiclient.discovery import build
import re
from auth import get_credentials
from datetime import datetime
import os
from openpyxl import load_workbook
from google_sheets_config import TRANSACTION_PREPROCESSING_FOLDER_ID

def main():
    # 사용자로부터 년월 입력 받기 (예: 2504, 2512)
    user_input = input("년월을 입력하세요 (예: 2504, 2512): ")
    payment_pattern = f'입금완료_{user_input}'
    
    # 구글 인증 자격 증명 가져오기
    creds = get_credentials()
    
    # 구글 시트 API 및 드라이브 API 클라이언트 생성
    sheets_service = build('sheets', 'v4', credentials=creds)
    drive_service = build('drive', 'v3', credentials=creds)
    
    # 스프레드시트 ID (다운로드용)
    DOWNLOAD_SPREADSHEET_ID = '1CK2UXTy7HKjBe2T0ovm5hfzAAKZxZAR_ev3cbTPOMPs'
    
    # 스프레드시트의 모든 시트 정보 가져오기
    spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=DOWNLOAD_SPREADSHEET_ID).execute()
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
        
        # '완료_'가 포함된 시트는 건너뛰기
        if '완료_' in sheet_title:
            continue
        
        print(f'{sheet_title} 시트 처리 중...')
        
        # 시트 데이터 범위 설정 (9행부터 모든 데이터)
        RANGE_NAME = f'{sheet_title}!A9:Z'
        
        try:
            # 시트 데이터 가져오기
            result = sheets_service.spreadsheets().values().get(
                spreadsheetId=DOWNLOAD_SPREADSHEET_ID,
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
    
    # 구글 드라이브에서 년월 폴더 찾기
    print(f"\n'{user_input}' 폴더 찾는 중...")
    query = f"'{TRANSACTION_PREPROCESSING_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.folder' and name='{user_input}' and trashed=false"
    existing_folders = drive_service.files().list(q=query, fields='files(id, name)').execute()
    
    if not existing_folders.get('files'):
        print(f"❌ '{user_input}' 폴더를 찾을 수 없습니다.")
        return
    
    date_folder_id = existing_folders['files'][0]['id']
    print(f"✅ '{user_input}' 폴더를 찾았습니다.")
    
    # 폴더 안에 스프레드시트가 있는지 확인
    print(f"\n폴더 안에 스프레드시트 확인 중...")
    spreadsheet_query = f"'{date_folder_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
    existing_spreadsheets = drive_service.files().list(q=spreadsheet_query, fields='files(id, name)').execute()
    
    if existing_spreadsheets.get('files'):
        # 기존 스프레드시트가 있으면 첫 번째 것을 사용
        target_spreadsheet_id = existing_spreadsheets['files'][0]['id']
        target_spreadsheet_name = existing_spreadsheets['files'][0]['name']
        print(f"✅ 기존 스프레드시트를 사용합니다: {target_spreadsheet_name}")
        
        # 기존 시트 목록 확인
        spreadsheet_info = sheets_service.spreadsheets().get(spreadsheetId=target_spreadsheet_id).execute()
        existing_sheet_names = {sheet.get('properties', {}).get('title', '') for sheet in spreadsheet_info.get('sheets', [])}
        
        # 새 시트명 생성 (타임스탬프 포함)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        new_sheet_name = f'전처리_구글시트_{timestamp}'
        counter = 2
        while new_sheet_name in existing_sheet_names:
            new_sheet_name = f'전처리_구글시트_{timestamp}_{counter}'
            counter += 1
        
        # 새 시트 추가
        add_sheet_request = {
            'requests': [{
                'addSheet': {
                    'properties': {
                        'title': new_sheet_name
                    }
                }
            }]
        }
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=target_spreadsheet_id,
            body=add_sheet_request
        ).execute()
        print(f"✅ 새 시트 '{new_sheet_name}' 추가 완료")
        
        # 헤더 추가 (A열부터: 항목, 이름, 번호, ...)
        header = ['항목', '이름', '번호', '-', '계좌', '주민번호', '입금액', '상태']
        
        # 데이터 업로드 (헤더 포함)
        values = selected_df.fillna('').astype(str).values.tolist()
        # 헤더를 첫 번째 행으로 추가
        values_with_header = [header] + values
        range_name = f"{new_sheet_name}!A1"
        body = {'values': values_with_header}
        sheets_service.spreadsheets().values().update(
            spreadsheetId=target_spreadsheet_id,
            range=range_name,
            valueInputOption='USER_ENTERED',
            body=body
        ).execute()
        
        # F열 텍스트 형식 설정
        spreadsheet_info = sheets_service.spreadsheets().get(spreadsheetId=target_spreadsheet_id).execute()
        sheet_id = None
        for sheet in spreadsheet_info.get('sheets', []):
            if sheet.get('properties', {}).get('title') == new_sheet_name:
                sheet_id = sheet.get('properties', {}).get('sheetId')
                break
        
        if sheet_id is not None:
            format_requests = [{
                'repeatCell': {
                    'range': {
                        'sheetId': sheet_id,
                        'startRowIndex': 1,  # 헤더 제외
                        'endRowIndex': len(values) + 1,
                        'startColumnIndex': 5,  # F열 (인덱스 5)
                        'endColumnIndex': 6
                    },
                    'cell': {
                        'userEnteredFormat': {
                            'numberFormat': {
                                'type': 'TEXT'
                            }
                        }
                    },
                    'fields': 'userEnteredFormat.numberFormat'
                }
            }]
            
            sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=target_spreadsheet_id,
                body={'requests': format_requests}
            ).execute()
            print(f"✅ F열 텍스트 형식 설정 완료")
        
        spreadsheet_url = f"https://docs.google.com/spreadsheets/d/{target_spreadsheet_id}/edit"
        print(f"\n✅ 데이터가 '{new_sheet_name}' 시트에 업로드되었습니다.")
        print(f"스프레드시트 URL: {spreadsheet_url}")
    else:
        # 스프레드시트가 없으면 새로 생성
        print(f"새 스프레드시트 생성 중...")
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        spreadsheet_title = f'전처리_구글시트_{timestamp}'
        
        spreadsheet = {
            'properties': {
                'title': spreadsheet_title
            }
        }
        
        created_spreadsheet = sheets_service.spreadsheets().create(
            body=spreadsheet,
            fields='spreadsheetId,spreadsheetUrl'
        ).execute()
        
        target_spreadsheet_id = created_spreadsheet.get('spreadsheetId')
        spreadsheet_url = created_spreadsheet.get('spreadsheetUrl')
        
        print(f"✅ 스프레드시트 생성 완료: {spreadsheet_title}")
        
        # 생성된 스프레드시트를 날짜 폴더로 이동
        file = drive_service.files().get(fileId=target_spreadsheet_id, fields='parents').execute()
        previous_parents = ",".join(file.get('parents'))
        
        drive_service.files().update(
            fileId=target_spreadsheet_id,
            addParents=date_folder_id,
            removeParents=previous_parents,
            fields='id, parents'
        ).execute()
        
        print(f"✅ 스프레드시트를 '{user_input}' 폴더로 이동 완료")
        
        # 헤더 추가 (A열부터: 항목, 이름, 번호, ...)
        header = ['항목', '이름', '번호', '-', '계좌', '주민번호', '입금액', '상태']
        
        # 데이터 업로드 (헤더 포함)
        values = selected_df.fillna('').astype(str).values.tolist()
        # 헤더를 첫 번째 행으로 추가
        values_with_header = [header] + values
        range_name = "Sheet1!A1"
        body = {'values': values_with_header}
        sheets_service.spreadsheets().values().update(
            spreadsheetId=target_spreadsheet_id,
            range=range_name,
            valueInputOption='USER_ENTERED',
            body=body
        ).execute()
        
        # F열 텍스트 형식 설정
        format_requests = [{
            'repeatCell': {
                'range': {
                    'sheetId': 0,  # Sheet1의 sheetId는 0
                    'startRowIndex': 1,  # 헤더 제외 (헤더는 0번째 행)
                    'endRowIndex': len(values) + 1,  # 헤더 포함한 전체 행 수
                    'startColumnIndex': 5,  # F열 (인덱스 5)
                    'endColumnIndex': 6
                },
                'cell': {
                    'userEnteredFormat': {
                        'numberFormat': {
                            'type': 'TEXT'
                        }
                    }
                },
                'fields': 'userEnteredFormat.numberFormat'
            }
        }]
        
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=target_spreadsheet_id,
            body={'requests': format_requests}
        ).execute()
        print(f"✅ F열 텍스트 형식 설정 완료")
        print(f"\n✅ 데이터가 '{spreadsheet_title}' 스프레드시트에 업로드되었습니다.")
        print(f"스프레드시트 URL: {spreadsheet_url}")
    
    print(f'\n총 {total_items}개의 항목이 추출되었습니다.')

if __name__ == '__main__':
    main()
