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
from openpyxl import load_workbook
from googleapiclient.discovery import build
from auth import get_credentials
from google_sheets_config import TRANSACTION_PREPROCESSING_FOLDER_ID
from datetime import datetime
import re

# 현재 폴더 경로
current_dir = Path('.')

# '거래내역조회_'가 포함된 엑셀 파일 찾기 (현재 폴더에만)
excel_files = []
for file_path in current_dir.glob('*거래내역조회_*.xls*'):
    if file_path.is_file():  # 파일인지 확인
        excel_files.append(file_path)

# 파일명 기준으로 정렬
excel_files.sort()

# 번호를 붙여서 출력
print(f"총 {len(excel_files)}개의 파일을 찾았습니다:\n")
for idx, file_path in enumerate(excel_files, 1):
    print(f"{idx}. {file_path}")

# 사용자로부터 번호 입력 받기
if excel_files:
    try:
        choice = int(input("\n처리할 파일 번호를 선택하세요: "))
        if 1 <= choice <= len(excel_files):
            selected_file = excel_files[choice - 1]
            print(f"\n선택된 파일: {selected_file}")
            
            # 전처리 파일명 생성 (구글시트 시트명으로 사용)
            original_path = Path(selected_file)
            sheet_name = f"{original_path.stem}_전처리"
            
            # 파일명에서 날짜 추출 (예: 거래내역조회_20260106 -> 2026년 1월 6일)
            # 그리고 전월 계산 (2026년 1월 -> 2025년 12월 -> 2512)
            date_match = re.search(r'(\d{8})', original_path.stem)
            if date_match:
                date_str = date_match.group(1)  # 예: 20260106
                file_year = int(date_str[:4])  # 2026
                file_month = int(date_str[4:6])  # 01
                
                # 전월 계산
                if file_month == 1:
                    prev_month = 12
                    prev_year = file_year - 1
                else:
                    prev_month = file_month - 1
                    prev_year = file_year
                
                # 연월 형식으로 변환 (예: 2512)
                folder_name = f"{str(prev_year)[2:]}{prev_month:02d}"
                print(f"📅 파일 날짜: {file_year}년 {file_month}월")
                print(f"📁 폴더명: {folder_name} ({prev_year}년 {prev_month}월)")
            else:
                # 날짜를 찾을 수 없으면 현재 날짜 기준으로 전월 계산
                today = datetime.now()
                if today.month == 1:
                    prev_month = 12
                    prev_year = today.year - 1
                else:
                    prev_month = today.month - 1
                    prev_year = today.year
                folder_name = f"{str(prev_year)[2:]}{prev_month:02d}"
                print(f"⚠️  파일명에서 날짜를 찾을 수 없어 현재 날짜 기준으로 전월 계산: {folder_name}")
            
            # 구글 인증 자격 증명 가져오기
            print("구글 인증 중...")
            creds = get_credentials()
            
            # 구글 시트 API 및 드라이브 API 클라이언트 생성
            sheets_service = build('sheets', 'v4', credentials=creds)
            drive_service = build('drive', 'v3', credentials=creds)
            
            # 구글 드라이브 폴더 ID (config 파일에서 가져옴)
            FOLDER_ID = TRANSACTION_PREPROCESSING_FOLDER_ID
            
            # 엑셀 파일 읽기
            print("엑셀 파일 읽는 중...")
            df = pd.read_excel(selected_file, header=None)
            
            print(f"원본 파일 크기: {df.shape[0]}행 x {df.shape[1]}열")
            
            # 원본 파일의 D7:D999 합계 계산 (마지막 행은 SUM이므로 제외)
            original_d_sum = 0
            if df.shape[1] > 3 and len(df) > 6:  # D열이 있고 7행 이상인지 확인
                # D7부터 마지막 행-1까지 (마지막 행은 SUM이므로 제외)
                start_idx = 6  # D7 (인덱스 6)
                end_idx = len(df) - 2  # 마지막 행 제외 (인덱스는 len-2)
                if end_idx >= start_idx:
                    d_col_original = pd.to_numeric(df.iloc[start_idx:end_idx+1, 3], errors='coerce')
                    original_d_sum = d_col_original.sum()
                    print(f"\n원본 파일 D7:D{end_idx+1} 합계 (마지막 SUM 행 제외): {original_d_sum:,.0f}")
            
            # a. 1-5행 삭제 (인덱스 0-4)
            if len(df) >= 5:
                df = df.iloc[5:].reset_index(drop=True)
                print("1-5행 삭제 완료")
            
            # b. E, F, G열 삭제 (인덱스 4, 5, 6)
            columns_to_drop = []
            if df.shape[1] > 4:
                columns_to_drop.append(4)  # E열
            if df.shape[1] > 5:
                columns_to_drop.append(5)  # F열
            if df.shape[1] > 6:
                columns_to_drop.append(6)  # G열
            
            if columns_to_drop:
                df = df.drop(df.columns[columns_to_drop], axis=1)
                print(f"E, F, G열 삭제 완료")
            
            # c. 마지막 행 삭제
            if len(df) > 0:
                df = df.iloc[:-1].reset_index(drop=True)
                print("마지막 행 삭제 완료")
            
            print(f"처리 후 파일 크기: {df.shape[0]}행 x {df.shape[1]}열")
            
            # 제외할 키워드 리스트
            exclude_keywords = [
                # 개인명
                '신현빈', '김지수', '임채빈', '정철호',
                # 회사/기관
                '회사', '아이플', '네이버', '디베스트컴퍼니', '애드온비',
                # 보험/연금
                '산재보험', '국민연금', '국민건강',
                # 통신사
                'SKTL',
                # 카드사, 결제
                '삼성카드', '현대카드', '국민카드', 'GSPAY', 'GSPay', '페이',
                # 세금
                '지방세', '소득세', '부가가치세', '관세', '경찰청', '과태료',
                # 쇼핑몰
                '무신사', '쿠팡', '당근',
                # 기타
                '월세', '배송비', '29고6425'
            ]
            
            # 디버깅: 키워드 리스트 확인
            print(f"\n📋 제외 키워드 리스트 ({len(exclude_keywords)}개):")
            print(f"  - 전체 키워드: {exclude_keywords}")
            
            # C열(인덱스 2)에서 키워드가 포함된 행 찾기
            if df.shape[1] > 2:  # C열이 있는지 확인
                # C열의 값이 문자열인지 확인하고 키워드 포함 여부 체크 (대소문자 구분 없이)
                def check_keyword(cell_value):
                    cell_str = str(cell_value)
                    cell_str_upper = cell_str.upper()
                    for keyword in exclude_keywords:
                        keyword_upper = keyword.upper()
                        # 영문은 대소문자 구분 없이, 한글은 그대로 비교
                        if keyword in cell_str or keyword_upper in cell_str_upper:
                            return True
                    return False
                
                exclude_mask = df.iloc[:, 2].astype(str).apply(check_keyword)
                
                # 디버깅: 모든 키워드에 대한 매칭 테스트
                print(f"\n🔍 키워드 매칭 상세 테스트:")
                keyword_match_count = {keyword: 0 for keyword in exclude_keywords}
                
                for idx, row in df.iterrows():
                    if len(row) > 2:
                        c_value = str(row.iloc[2])
                        is_excluded = exclude_mask.iloc[idx] if idx < len(exclude_mask) else False
                        
                        # 각 키워드별로 매칭 확인
                        matched_keywords = []
                        for keyword in exclude_keywords:
                            if keyword in c_value or keyword.upper() in c_value.upper():
                                matched_keywords.append(keyword)
                                keyword_match_count[keyword] += 1
                        
                        if matched_keywords:
                            print(f"  - 행 {idx+1}: C열='{c_value}'")
                            print(f"    → 매칭된 키워드: {matched_keywords}")
                            print(f"    → exclude_mask 결과: {is_excluded}")
                            if not is_excluded:
                                print(f"    → ⚠️  키워드가 매칭되었지만 제외되지 않음!")
                
                # 키워드별 매칭 통계
                print(f"\n📊 키워드별 매칭 통계:")
                for keyword, count in keyword_match_count.items():
                    if count > 0:
                        print(f"  - '{keyword}': {count}개 행 매칭")
                
                # 제외할 행 (sheet2로 이동)
                df_excluded = df[exclude_mask].copy()
                # 남은 행 (sheet1에 유지)
                df_main = df[~exclude_mask].copy()
                
                print(f"\nC열 키워드 검사 결과:")
                print(f"- 제외할 행: {len(df_excluded)}개")
                print(f"- 유지할 행: {len(df_main)}개")
                
                # 구글 드라이브 폴더에 날짜 폴더 생성 및 새로운 스프레드시트 생성
                print(f"\n구글 드라이브 폴더에 날짜 폴더 생성 중...")
                
                # 날짜 폴더가 이미 존재하는지 확인
                query = f"'{FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.folder' and name='{folder_name}' and trashed=false"
                existing_folders = drive_service.files().list(q=query, fields='files(id, name)').execute()
                
                if existing_folders.get('files'):
                    date_folder_id = existing_folders['files'][0]['id']
                    print(f"✅ 날짜 폴더 '{folder_name}'가 이미 존재합니다.")
                else:
                    # 날짜 폴더 생성
                    folder_metadata = {
                        'name': folder_name,
                        'mimeType': 'application/vnd.google-apps.folder',
                        'parents': [FOLDER_ID]
                    }
                    date_folder = drive_service.files().create(
                        body=folder_metadata,
                        fields='id, name'
                    ).execute()
                    date_folder_id = date_folder.get('id')
                    print(f"✅ 날짜 폴더 '{folder_name}' 생성 완료")
                
                # 새로운 스프레드시트 생성
                print(f"\n날짜 폴더에 새로운 스프레드시트 생성 중...")
                spreadsheet_title = sheet_name
                spreadsheet = {
                    'properties': {
                        'title': spreadsheet_title
                    },
                    'sheets': [
                        {'properties': {'title': 'Sheet0'}},
                        {'properties': {'title': 'Sheet1'}},
                        {'properties': {'title': 'Sheet2'}}
                    ]
                }
                
                created_spreadsheet = sheets_service.spreadsheets().create(
                    body=spreadsheet,
                    fields='spreadsheetId,spreadsheetUrl'
                ).execute()
                
                SPREADSHEET_ID = created_spreadsheet.get('spreadsheetId')
                spreadsheet_url = created_spreadsheet.get('spreadsheetUrl')
                
                print(f"✅ 스프레드시트 생성 완료: {spreadsheet_title}")
                print(f"   스프레드시트 ID: {SPREADSHEET_ID}")
                
                # 생성된 스프레드시트를 날짜 폴더로 이동
                # 먼저 기존 부모(내 드라이브)를 가져옴
                file = drive_service.files().get(fileId=SPREADSHEET_ID, fields='parents').execute()
                previous_parents = ",".join(file.get('parents'))
                
                # 날짜 폴더로 이동 (기존 부모 제거하고 새 부모로 이동)
                drive_service.files().update(
                    fileId=SPREADSHEET_ID,
                    addParents=date_folder_id,
                    removeParents=previous_parents,
                    fields='id, parents'
                ).execute()
                
                print(f"✅ 스프레드시트를 '{folder_name}' 폴더로 이동 완료")
                
                # 시트 ID 매핑
                spreadsheet_info = sheets_service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
                sheet_ids = {}
                for sheet in spreadsheet_info.get('sheets', []):
                    sheet_props = sheet.get('properties', {})
                    sheet_title = sheet_props.get('title', '')
                    sheet_id = sheet_props.get('sheetId')
                    
                    if sheet_title == 'Sheet0':
                        sheet_ids['Sheet0'] = {'id': sheet_id, 'title': 'Sheet0'}
                    elif sheet_title == 'Sheet1':
                        sheet_ids['Sheet1'] = {'id': sheet_id, 'title': 'Sheet1'}
                    elif sheet_title == 'Sheet2':
                        sheet_ids['Sheet2'] = {'id': sheet_id, 'title': 'Sheet2'}
                
                print(f"✅ Sheet0, Sheet1, Sheet2 시트 준비 완료")
                
                # Sheet1에 데이터 업로드
                sheet1_title = sheet_ids['Sheet1']['title']
                if len(df_main) > 0:
                    # 데이터프레임을 리스트로 변환
                    values_main = df_main.fillna('').astype(str).values.tolist()
                    range_main = f"{sheet1_title}!A1"
                    body_main = {'values': values_main}
                    sheets_service.spreadsheets().values().update(
                        spreadsheetId=SPREADSHEET_ID,
                        range=range_main,
                        valueInputOption='USER_ENTERED',
                        body=body_main
                    ).execute()
                    print(f"✅ {sheet1_title}에 {len(df_main)}행 데이터 업로드 완료")
                
                # Sheet2에 데이터 업로드
                sheet2_title = sheet_ids['Sheet2']['title']
                if len(df_excluded) > 0:
                    values_excluded = df_excluded.fillna('').astype(str).values.tolist()
                    range_excluded = f"{sheet2_title}!A1"
                    body_excluded = {'values': values_excluded}
                    sheets_service.spreadsheets().values().update(
                        spreadsheetId=SPREADSHEET_ID,
                        range=range_excluded,
                        valueInputOption='USER_ENTERED',
                        body=body_excluded
                    ).execute()
                    print(f"✅ {sheet2_title}에 {len(df_excluded)}행 데이터 업로드 완료")
                
                # Sheet0에 합계 정보 및 수식 작성
                sheet0_title = sheet_ids['Sheet0']['title']
                sheet1_title_for_formula = sheet_ids['Sheet1']['title']
                sheet2_title_for_formula = sheet_ids['Sheet2']['title']
                
                # Sheet0 데이터 준비
                sheet0_values = [
                    ['항목', '금액'],
                    ['Sheet1 D열 합계', f'=SUM({sheet1_title_for_formula}!D2:D999)'],
                    ['Sheet2 D열 합계', f'=SUM({sheet2_title_for_formula}!D2:D999)'],
                    ['총합', '=SUM(B2:B3)']
                ]
                
                range_sheet0 = f"{sheet0_title}!A1"
                body_sheet0 = {'values': sheet0_values}
                sheets_service.spreadsheets().values().update(
                    spreadsheetId=SPREADSHEET_ID,
                    range=range_sheet0,
                    valueInputOption='USER_ENTERED',
                    body=body_sheet0
                ).execute()
                print(f"✅ {sheet0_title}에 합계 정보 업로드 완료")
                
                # 콘솔에 출력할 합계 계산 (참고용)
                sheet1_sum = 0
                sheet2_sum = 0
                
                if df_main.shape[1] > 3:  # D열이 있는지 확인
                    d_col_main = pd.to_numeric(df_main.iloc[:, 3], errors='coerce')
                    sheet1_sum = d_col_main.sum()
                
                if df_excluded.shape[1] > 3:  # D열이 있는지 확인
                    d_col_excluded = pd.to_numeric(df_excluded.iloc[:, 3], errors='coerce')
                    sheet2_sum = d_col_excluded.sum()
                
                total_sum = sheet1_sum + sheet2_sum
                
                print(f"\nD열 합계 (참고용):")
                print(f"- 원본 파일 D7:D999: {original_d_sum:,.0f}")
                print(f"- Sheet1: {sheet1_sum:,.0f}")
                print(f"- Sheet2: {sheet2_sum:,.0f}")
                print(f"- 총합: {total_sum:,.0f}")
                print("(Sheet0에는 수식으로 저장되었습니다)")
                
                print(f"\n처리 완료! 구글 드라이브 폴더에 스프레드시트가 생성되었습니다.")
                print(f"스프레드시트 URL: {spreadsheet_url}")
                print(f"- {sheet_ids['Sheet0']['title']}: 합계 정보")
                print(f"- {sheet_ids['Sheet1']['title']}: {len(df_main)}행 (일반 데이터)")
                print(f"- {sheet_ids['Sheet2']['title']}: {len(df_excluded)}행 (제외된 데이터)")
            else:
                # C열이 없는 경우 구글 드라이브 폴더에 날짜 폴더 생성 및 새로운 스프레드시트 생성
                print(f"\n구글 드라이브 폴더에 날짜 폴더 생성 중...")
                
                # 날짜 폴더가 이미 존재하는지 확인
                query = f"'{FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.folder' and name='{folder_name}' and trashed=false"
                existing_folders = drive_service.files().list(q=query, fields='files(id, name)').execute()
                
                if existing_folders.get('files'):
                    date_folder_id = existing_folders['files'][0]['id']
                    print(f"✅ 날짜 폴더 '{folder_name}'가 이미 존재합니다.")
                else:
                    # 날짜 폴더 생성
                    folder_metadata = {
                        'name': folder_name,
                        'mimeType': 'application/vnd.google-apps.folder',
                        'parents': [FOLDER_ID]
                    }
                    date_folder = drive_service.files().create(
                        body=folder_metadata,
                        fields='id, name'
                    ).execute()
                    date_folder_id = date_folder.get('id')
                    print(f"✅ 날짜 폴더 '{folder_name}' 생성 완료")
                
                # 새로운 스프레드시트 생성
                print(f"\n날짜 폴더에 새로운 스프레드시트 생성 중...")
                spreadsheet_title = sheet_name
                spreadsheet = {
                    'properties': {
                        'title': spreadsheet_title
                    }
                }
                
                created_spreadsheet = sheets_service.spreadsheets().create(
                    body=spreadsheet,
                    fields='spreadsheetId,spreadsheetUrl'
                ).execute()
                
                SPREADSHEET_ID = created_spreadsheet.get('spreadsheetId')
                spreadsheet_url = created_spreadsheet.get('spreadsheetUrl')
                
                print(f"✅ 스프레드시트 생성 완료: {spreadsheet_title}")
                print(f"   스프레드시트 ID: {SPREADSHEET_ID}")
                
                # 생성된 스프레드시트를 날짜 폴더로 이동
                # 먼저 기존 부모(내 드라이브)를 가져옴
                file = drive_service.files().get(fileId=SPREADSHEET_ID, fields='parents').execute()
                previous_parents = ",".join(file.get('parents'))
                
                # 날짜 폴더로 이동
                drive_service.files().update(
                    fileId=SPREADSHEET_ID,
                    addParents=date_folder_id,
                    removeParents=previous_parents,
                    fields='id, parents'
                ).execute()
                
                print(f"✅ 스프레드시트를 '{folder_name}' 폴더로 이동 완료")
                
                # 데이터 업로드
                values = df.fillna('').astype(str).values.tolist()
                range_name = "Sheet1!A1"
                body = {'values': values}
                sheets_service.spreadsheets().values().update(
                    spreadsheetId=SPREADSHEET_ID,
                    range=range_name,
                    valueInputOption='USER_ENTERED',
                    body=body
                ).execute()
                print(f"✅ Sheet1에 {len(df)}행 데이터 업로드 완료")
                
                print(f"\n처리 완료! 구글 드라이브 폴더에 스프레드시트가 생성되었습니다.")
                print(f"스프레드시트 URL: {spreadsheet_url}")
                print("(C열이 없어 키워드 검사를 수행하지 않았습니다)")
                print(f"\n원본 파일 D7:D999 합계: {original_d_sum:,.0f}")
        else:
            print("잘못된 번호입니다.")
    except ValueError:
        print("숫자를 입력해주세요.")
    except Exception as e:
        print(f"오류 발생: {e}")
else:
    print("처리할 파일이 없습니다.")

