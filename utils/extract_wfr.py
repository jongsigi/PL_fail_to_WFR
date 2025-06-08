import pandas as pd

def process_elec_items(df):
    """
    데이터프레임의 elec_item과 wafer_list 열을 처리하는 함수
    """
    pattern_dict = {
        'ERASE_CH0_CE0': '"1100":{1:{',  # 일반 따옴표 사용
        'ERASE_CH0_CE1': '"1400":{1:{'   # 일반 따옴표 사용
    }
    
    # 결과를 저장할 리스트
    results = []
    
    for index, row in df.iterrows():
        elec_items = [item.strip() for item in row['elec_item'].split(',')]
        # wafer_list 데이터의 따옴표 정규화
        wafer_data = str(row['wafer_list']).strip().replace('\u201C', '"').replace('\u201D', '"')
        lot_id = row['lot_id']
        unit_id = row['unit_id']
        
        for item in elec_items:
            if item in pattern_dict:
                pattern = pattern_dict[item]
                start_idx = wafer_data.find(pattern)
                
                if start_idx != -1:
                    start_idx += len(pattern)
                    end_idx = wafer_data.find("}", start_idx)
                    if end_idx != -1:
                        extracted_data = wafer_data[start_idx:end_idx]
                        # 데이터 파싱 (이제 일반 따옴표 사용)
                        wafer_info = extracted_data.split('"')[1]  # "59CR138-12" 에서 59CR138-12 추출
                        run_id, no = wafer_info.split('-')  # 59CR138과 12로 분리
                        coords = extracted_data.split(':')[1]  # 11_10 추출
                        x, y = coords.split('_')  # 11과 10으로 분리
                        
                        # 결과 딕셔너리 생성
                        result = {
                            'lot_id': lot_id,
                            'unit_id': unit_id,
                            'elec_item': item,
                            'run_id': run_id,
                            'no': no,
                            'x': x,
                            'y': y
                        }
                        results.append(result)
    
    # 결과를 데이터프레임으로 변환
    result_df = pd.DataFrame(results)
    return result_df

def save_to_excel(original_file, result_df, sheet_name='Parsed_Results'):
    """
    결과 데이터프레임을 엑셀 파일의 새로운 시트에 저장
    """
    with pd.ExcelWriter(original_file, mode='a', engine='openpyxl') as writer:
        result_df.to_excel(writer, sheet_name=sheet_name, index=False)