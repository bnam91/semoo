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

from googleapiclient.discovery import build
from auth import get_credentials
from datetime import datetime
from pathlib import Path
import pandas as pd
import webbrowser

def main():
    # 구글 인증 자격 증명 가져오기
    creds = get_credentials()
    
    # 구글 시트 API 클라이언트 생성
    sheets_service = build('sheets', 'v4', credentials=creds)
    
    # 스프레드시트 ID (URL에서 추출)
    SPREADSHEET_ID = '1Qk_Jlchp0RczrWPsDfNqMMtRuGmJpd3wXD6hpxbxHuk'
    
    try:
        # 전월 계산 (예: 2025.11 -> 2510)
        today = datetime.now()
        if today.month == 1:
            prev_month = 12
            prev_year = today.year - 1
        else:
            prev_month = today.month - 1
            prev_year = today.year
        
        # 연월 형식으로 변환 (예: 2510)
        year_month = f"{str(prev_year)[2:]}{prev_month:02d}"
        print(f"전월: {year_month}")
        
        # 스프레드시트의 모든 시트 정보 가져오기
        spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
        sheets = spreadsheet.get('sheets', [])
        
        # 양식시트 찾기
        template_sheet_id = None
        for sheet in sheets:
            sheet_title = sheet.get('properties', {}).get('title', '')
            if sheet_title == '양식시트':
                template_sheet_id = sheet.get('properties', {}).get('sheetId')
                break
        
        if template_sheet_id is None:
            print("❌ '양식시트'를 찾을 수 없습니다.")
            return
        
        print(f"양식시트 ID: {template_sheet_id}")
        
        # A4 셀 값 읽기
        RANGE_NAME = '양식시트!A4'
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=RANGE_NAME
        ).execute()
        
        # 데이터 처리
        values = result.get('values', [])
        
        if values:
            # A4 셀 값 출력
            a4_value = values[0][0] if len(values[0]) > 0 else ''
            print(f"A4 셀 값: {a4_value}")
        else:
            print("A4 셀이 비어있습니다.")
        
        # 이미 같은 이름의 시트가 있는지 확인하고 고유한 이름 생성
        existing_sheet_names = {sheet.get('properties', {}).get('title', '') for sheet in sheets}
        new_sheet_name = year_month
        counter = 2
        
        while new_sheet_name in existing_sheet_names:
            new_sheet_name = f"{year_month}_{counter}"
            counter += 1
        
        if new_sheet_name != year_month:
            print(f"\n⚠️  '{year_month}' 시트가 이미 존재합니다. '{new_sheet_name}'로 생성합니다.")
        
        # 양식시트 복사
        print(f"\n양식시트를 '{new_sheet_name}'로 복사 중...")
        duplicate_request = {
            'requests': [{
                'duplicateSheet': {
                    'sourceSheetId': template_sheet_id,
                    'newSheetName': new_sheet_name
                }
            }]
        }
        
        response = sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body=duplicate_request
        ).execute()
        
        new_sheet_id = response.get('replies')[0].get('duplicateSheet').get('properties').get('sheetId')
        print(f"✅ 시트 복사 완료! 새 시트명: '{new_sheet_name}' (ID: {new_sheet_id})")
        
        # 로컬 폴더에서 '(전처리)'가 포함된 xls 파일 찾기
        current_dir = Path('.')
        data_files = []
        for file_path in current_dir.glob('*(전처리)*.xls*'):
            if file_path.is_file():
                data_files.append(file_path)
        
        data_files.sort()
        
        if not data_files:
            print("\n❌ '(전처리)'가 포함된 xls 파일을 찾을 수 없습니다.")
            return
        
        # 파일 리스트 출력
        print(f"\n총 {len(data_files)}개의 데이터 파일을 찾았습니다:\n")
        for idx, file_path in enumerate(data_files, 1):
            print(f"{idx}. {file_path.name}")
        
        # 사용자로부터 번호 입력 받기
        try:
            choice = int(input("\n데이터 파일 번호를 선택하세요: "))
            if not (1 <= choice <= len(data_files)):
                print("잘못된 번호입니다.")
                return
            
            selected_data_file = data_files[choice - 1]
            print(f"\n선택된 데이터 파일: {selected_data_file.name}")
            
            # 데이터 파일에서 '후처리' 시트 읽기
            print(f"\n데이터 파일 읽는 중...")
            try:
                df_data = pd.read_excel(selected_data_file, sheet_name='후처리')
                print(f"✅ 데이터 읽기 완료: {len(df_data)}행")
            except Exception as e:
                print(f"❌ '후처리' 시트를 찾을 수 없습니다: {e}")
                return
            
            # 데이터를 스프레드시트에 입력
            print(f"\n스프레드시트에 데이터 입력 중...")
            
            # A열에 '고야'가 포함된 행과 그렇지 않은 행 분리
            goya_rows = []
            other_rows = []
            excluded_count = 0
            
            for idx, row_data in df_data.iterrows():
                # A열 값 확인 (인덱스 0)
                a_value = ''
                if len(df_data.columns) > 0 and pd.notna(row_data.iloc[0]):
                    a_value = str(row_data.iloc[0])
                
                # '726호'가 포함된 행은 제외
                if '726호' in a_value:
                    excluded_count += 1
                    continue
                
                if '고야' in a_value:
                    goya_rows.append((idx, row_data))
                else:
                    other_rows.append((idx, row_data))
            
            print(f"  - '726호' 제외된 행: {excluded_count}개")
            print(f"  - '고야' 포함 행: {len(goya_rows)}개")
            print(f"  - 기타 행: {len(other_rows)}개")
            
            # '고야'가 포함된 행을 먼저, 나머지를 나중에 배치
            sorted_rows = goya_rows + other_rows
            
            # 데이터 준비 (4행부터 시작)
            start_row = 4
            values_to_update = []
            
            # B열에 입력할 전월 형식 (예: 2025-10)
            date_string = f"{prev_year}-{prev_month:02d}"
            
            def convert_date_format(date_str):
                """날짜 형식 변환: '251002' -> '2025.10.02'"""
                if pd.isna(date_str) or not date_str:
                    return ''
                
                date_str = str(date_str).strip()
                # 6자리 숫자 형식인지 확인 (예: 251002)
                if len(date_str) == 6 and date_str.isdigit():
                    year = '20' + date_str[:2]  # 25 -> 2025
                    month = date_str[2:4]  # 10
                    day = date_str[4:6]  # 02
                    return f"{year}.{month}.{day}"
                return date_str  # 형식이 맞지 않으면 원본 반환
            
            # '고야'가 포함된 행의 실제 행 번호를 추적
            goya_row_numbers = []
            j_column_sum = 0  # J열 값의 합계
            
            for row_idx, (original_idx, row_data) in enumerate(sorted_rows):
                row_num = start_row + row_idx  # 실제 행 번호 (4부터 시작)
                row_values = [''] * 13  # A~M열까지 (인덱스 0~12)
                
                # '고야'가 포함된 행인지 확인
                a_value = ''
                if len(df_data.columns) > 0 and pd.notna(row_data.iloc[0]):
                    a_value = str(row_data.iloc[0])
                
                is_goya = '고야' in a_value
                if is_goya:
                    goya_row_numbers.append(row_num)
                
                # A열에 번호 입력 (1부터 시작)
                row_values[0] = row_idx + 1
                
                # B열에 전월 입력 (예: 2025-10)
                row_values[1] = date_string
                
                # H열 (인덱스 7) → C열 (인덱스 2) - 날짜 형식 변환
                if len(df_data.columns) > 7 and pd.notna(row_data.iloc[7]):
                    date_value = row_data.iloc[7]
                    row_values[2] = convert_date_format(date_value)
                
                # B열 (인덱스 1) → D열 (인덱스 3)
                if len(df_data.columns) > 1 and pd.notna(row_data.iloc[1]):
                    row_values[3] = row_data.iloc[1]
                
                # F열 (인덱스 5) → E열 (인덱스 4) - 텍스트로 입력
                if len(df_data.columns) > 5 and pd.notna(row_data.iloc[5]):
                    e_value = str(row_data.iloc[5])
                    # 숫자로 변환 가능한 경우 .0 제거
                    try:
                        # 숫자로 변환 시도
                        num_value = float(e_value)
                        if num_value == int(num_value):
                            e_value = str(int(num_value))
                        else:
                            e_value = str(num_value).rstrip('0').rstrip('.')
                    except ValueError:
                        pass  # 숫자가 아니면 그대로 사용
                    row_values[4] = e_value
                
                # '고야'가 포함된 행의 경우 F열에 '급여' 입력 (인덱스 5)
                if is_goya:
                    row_values[5] = '급여'
                
                # H열에 '기타자영업' 입력 (인덱스 7)
                row_values[7] = '기타자영업'
                
                # G열 (인덱스 6) → J열 (인덱스 9)
                if len(df_data.columns) > 6 and pd.notna(row_data.iloc[6]):
                    g_value = row_data.iloc[6]
                    row_values[9] = g_value
                    
                    # J열 값 합계 계산
                    try:
                        j_column_sum += float(g_value)
                    except (ValueError, TypeError):
                        pass  # 숫자가 아닌 경우 무시
                    
                    # I열에 G열 값 / 0.967 계산된 값 입력 (인덱스 8)
                    try:
                        i_value = float(g_value) / 0.967
                        row_values[8] = round(i_value, 1)  # 소수점 1자리로 반올림
                    except (ValueError, TypeError):
                        pass  # 숫자가 아닌 경우 무시
                
                # K열에 0.03 입력 (인덱스 10)
                row_values[10] = 0.03
                
                # L열에 수식 입력 (인덱스 11): =J{row_num}*K{row_num}
                row_values[11] = f'=J{row_num}*K{row_num}'
                
                # M열에 수식 입력 (인덱스 12): =L{row_num}*0.1
                row_values[12] = f'=L{row_num}*0.1'
                
                values_to_update.append(row_values)
            
            # 스프레드시트에 데이터 쓰기 (A~M열)
            range_name = f"{new_sheet_name}!A{start_row}:M{start_row + len(values_to_update) - 1}"
            
            body = {
                'values': values_to_update
            }
            
            # 수식과 텍스트 입력을 위해 USER_ENTERED 옵션 사용
            sheets_service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=range_name,
                valueInputOption='USER_ENTERED',
                body=body
            ).execute()
            
            print(f"✅ 데이터 입력 완료! {len(values_to_update)}행 입력됨")
            print(f"📊 입력 범위: {range_name}")
            print(f"💰 J열 값의 합계: {j_column_sum:,.0f}")
            
            # E열 텍스트 형식 및 I열 소수점 1자리 형식 적용
            format_requests = []
            
            # E열 전체를 텍스트 형식으로 설정
            format_requests.append({
                'repeatCell': {
                    'range': {
                        'sheetId': new_sheet_id,
                        'startRowIndex': start_row - 1,  # 0-based index (4행 = 인덱스 3)
                        'endRowIndex': start_row + len(values_to_update) - 1,
                        'startColumnIndex': 4,  # E열 (인덱스 4)
                        'endColumnIndex': 5
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
            })
            
            # I열 전체를 소수점 1자리 형식으로 설정
            format_requests.append({
                'repeatCell': {
                    'range': {
                        'sheetId': new_sheet_id,
                        'startRowIndex': start_row - 1,  # 0-based index (4행 = 인덱스 3)
                        'endRowIndex': start_row + len(values_to_update) - 1,
                        'startColumnIndex': 8,  # I열 (인덱스 8)
                        'endColumnIndex': 9
                    },
                    'cell': {
                        'userEnteredFormat': {
                            'numberFormat': {
                                'type': 'NUMBER',
                                'pattern': '0.0'
                            }
                        }
                    },
                    'fields': 'userEnteredFormat.numberFormat'
                }
            })
            
            # '고야'가 포함된 행에 서식 적용 (배경색 노란색, F열 글자 빨간색)
            if goya_row_numbers:
                print(f"\n'고야' 포함 행에 서식 적용 중... ({len(goya_row_numbers)}개 행)")
                
                for row_num in goya_row_numbers:
                    # 전체 행 배경색 노란색
                    format_requests.append({
                        'repeatCell': {
                            'range': {
                                'sheetId': new_sheet_id,
                                'startRowIndex': row_num - 1,  # 0-based index
                                'endRowIndex': row_num,
                                'startColumnIndex': 0,  # A열부터
                                'endColumnIndex': 13  # M열까지
                            },
                            'cell': {
                                'userEnteredFormat': {
                                    'backgroundColor': {
                                        'red': 1.0,
                                        'green': 1.0,
                                        'blue': 0.0
                                    }
                                }
                            },
                            'fields': 'userEnteredFormat.backgroundColor'
                        }
                    })
                    
                    # F열 글자 빨간색
                    format_requests.append({
                        'repeatCell': {
                            'range': {
                                'sheetId': new_sheet_id,
                                'startRowIndex': row_num - 1,  # 0-based index
                                'endRowIndex': row_num,
                                'startColumnIndex': 5,  # F열 (인덱스 5)
                                'endColumnIndex': 6
                            },
                            'cell': {
                                'userEnteredFormat': {
                                    'textFormat': {
                                        'foregroundColor': {
                                            'red': 1.0,
                                            'green': 0.0,
                                            'blue': 0.0
                                        }
                                    }
                                }
                            },
                            'fields': 'userEnteredFormat.textFormat.foregroundColor'
                        }
                    })
                
            # 배치 요청 실행
            if format_requests:
                format_body = {
                    'requests': format_requests
                }
                sheets_service.spreadsheets().batchUpdate(
                    spreadsheetId=SPREADSHEET_ID,
                    body=format_body
                ).execute()
                
                print(f"✅ 서식 적용 완료!")
            
            # 구글 시트를 크롬으로 열지 물어보기
            spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1Qk_Jlchp0RczrWPsDfNqMMtRuGmJpd3wXD6hpxbxHuk/edit?gid=0#gid=0'
            open_browser = input(f"\n구글 시트를 크롬으로 열까요? (Y/N): ").strip().upper()
            
            if open_browser == 'Y':
                try:
                    import os
                    import sys
                    # Windows에서 크롬 경로 찾기
                    if sys.platform == 'win32':
                        chrome_paths = [
                            r'C:\Program Files\Google\Chrome\Application\chrome.exe',
                            r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
                            os.path.expanduser(r'~\AppData\Local\Google\Chrome\Application\chrome.exe')
                        ]
                        
                        chrome_found = False
                        for chrome_path in chrome_paths:
                            if os.path.exists(chrome_path):
                                os.startfile(spreadsheet_url)
                                chrome_found = True
                                print(f"✅ 크롬으로 구글 시트를 열었습니다.")
                                break
                        
                        if not chrome_found:
                            # 크롬을 찾을 수 없으면 기본 브라우저로 열기
                            webbrowser.open(spreadsheet_url)
                            print(f"✅ 브라우저로 구글 시트를 열었습니다.")
                    else:
                        # Windows가 아닌 경우 기본 브라우저로 열기
                        webbrowser.open(spreadsheet_url)
                        print(f"✅ 브라우저로 구글 시트를 열었습니다.")
                except Exception as e:
                    # 오류 발생 시 기본 브라우저로 열기
                    webbrowser.open(spreadsheet_url)
                    print(f"✅ 브라우저로 구글 시트를 열었습니다.")
            
        except ValueError:
            print("숫자를 입력해주세요.")
            
    except Exception as e:
        print(f"오류 발생: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()

