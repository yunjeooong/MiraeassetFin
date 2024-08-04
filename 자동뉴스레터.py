  # -*- coding: utf-8 -*-
import requests
import pandas as pd
import os
import json

class CompletionExecutor:
    def __init__(self, host, api_key, api_key_primary_val, request_id):
        self._host = host
        self._api_key = api_key
        self._api_key_primary_val = api_key_primary_val
        self._request_id = request_id

    def execute(self, completion_request):
        headers = {
            'X-NCP-CLOVASTUDIO-API-KEY': self._api_key,
            'X-NCP-APIGW-API-KEY': self._api_key_primary_val,
            'X-NCP-CLOVASTUDIO-REQUEST-ID': self._request_id,
            'Content-Type': 'application/json; charset=utf-8',
            'Accept': 'text/event-stream'
        }

        response_text = ""
        with requests.post(self._host + '/testapp/v1/chat-completions/HCX-DASH-001',
                           headers=headers, json=completion_request, stream=True) as r:
            for line in r.iter_lines():
                if line:
                    decoded_line = line.decode("utf-8")
                    # "data:" 문자열이 있는지 확인
                    if "data:" in decoded_line:
                        try:
                            data = json.loads(decoded_line.split("data:")[1].strip())
                            if 'message' in data and 'content' in data['message']:
                                response_text += data['message']['content']
                        except json.JSONDecodeError:
                            continue
        return response_text

def read_questions_from_excel(file_path):
    df = pd.read_excel(file_path)
    return df

def save_answers_to_excel(df, file_path):
    # 디렉토리가 없으면 생성
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    # 중복을 피하기 위해 파일 이름 변경
    base, extension = os.path.splitext(file_path)
    counter = 1
    new_file_path = f"{base}_{counter}{extension}"
    while os.path.exists(new_file_path):
        counter += 1
        new_file_path = f"{base}_{counter}{extension}"
    df.to_excel(new_file_path, index=False)
    return new_file_path

if __name__ == '__main__':
    completion_executor = CompletionExecutor (

        host='https://clovastudio.stream.ntruss.com',
        api_key='api 키 입력',
        api_key_primary_val='입력',
        request_id='입력'
    )

    # 엑셀 파일의 절대 경로
    input_file_path = 'C:\\Users\\A\\Desktop\\새 폴더2\경제_top5종목.xlsx'
    df = read_questions_from_excel(input_file_path)

    # 열 이름 확인
    print(df.columns)

    # 필요한 열이 존재하는지 확인
    if 'top5종목' not in df.columns:
        raise KeyError("'종목' 열이 엑셀 파일에 존재하지 않습니다.")

    # 답변을 저장할 칼럼 추가
    df['투자 인사이트'] = ""

    for index, row in df.iterrows():
        if pd.isna(row['top5종목']) or row['top5종목'].strip() == "":
            print(f"Skipping empty summary at row {index + 1}")
            continue

        preset_text = [
            
         {"role": "system", "content": "2024년 해당 기업의 투자장보만을 줘 \n내가 준 기업에 대한 제일 최신  정보들을 '미래에셋증권에서 찾아' 그 내용을 요약해줘.\n 먼저, 미래에셋증권 홈페이지(https://securities.**********.***/main.do) 에 접속하시고 검색창에 해당 기업명을 입력해 . 그런 다음 화면 하단에서 '투자정보'를 누르고 '기업분석'을 클릭하여 제일 상단의 제일 최신의 자료를 요약해줘.\n 2024년 이전의 정보는 절대 주지마 "},
            {"role": "user", "content": row['top5종목']}
        ]

        request_data = {
            'messages': preset_text,
            'topP': 0.8,
            'topK': 0,
            'maxTokens': 256,
            'temperature': 0.75,
            'repeatPenalty': 5.0,
            'stopBefore': [],
            'includeAiFilters': True,
            'seed': 0
        }

        print(f"Processing row {index + 1}/{len(df)}")
        response = completion_executor.execute(request_data)
        df.at[index, '답변'] = response

    # 결과를 새로운 엑셀 파일로 저장
    output_file_path = 'C:\\Users\\A\\Desktop\\새 폴더2\\자동뉴스레터_경제.xlsx'
    new_file_path = save_answers_to_excel(df, output_file_path)

    print(f"Results saved to {new_file_path}")