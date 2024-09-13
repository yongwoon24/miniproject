import os
import pandas as pd
import chardet

# 변환할 폴더 경로
folder_path = r'D:\rpawork\workspace\mainpj\miniProJect\2019~2023'

def detect_encoding(file_path):
    with open(file_path, 'rb') as file:
        result = chardet.detect(file.read())
        return result['encoding']

def read_file(file_path):
    encoding = detect_encoding(file_path)
    try:
        with open(file_path, 'r', encoding=encoding) as file:
            return file.read()
    except UnicodeDecodeError:
        raise ValueError(f"파일을 읽을 수 없습니다: {file_path}")

# 폴더 내의 모든 파일 확인
for file_name in os.listdir(folder_path):
    # 파일 경로 생성
    file_path = os.path.join(folder_path, file_name)
    
    # 텍스트 파일인지 확인
    if file_name.endswith('.txt'):
        try:
            # 텍스트 파일 읽기
            data = read_file(file_path)
            
            # 데이터를 DataFrame으로 변환
            df = pd.DataFrame([line.split('\t') for line in data.splitlines()])
            
            # 엑셀 파일로 저장할 경로 설정
            excel_file_path = os.path.join(folder_path, file_name.replace('.txt', '.xlsx'))
            
            # DataFrame을 엑셀 파일로 저장
            df.to_excel(excel_file_path, index=False, header=False, engine='openpyxl')
            
            print(f'변환 완료: {file_name} -> {os.path.basename(excel_file_path)}')
        except Exception as e:
            print(f'파일 변환 실패: {file_name}. 오류: {e}')
