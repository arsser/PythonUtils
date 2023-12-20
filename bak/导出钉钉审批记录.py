import requests
import csv
import time
import pandas as pd

# 常量
APP_KEY = 'dingl5qlh2s1ddksf5ru'
APP_SECRET = 'mjr6mb1TpfMx1Q2W7IfSVvrTsD855sSMQcL_XNZQA_HVhCYBKT-FDm8VPqGpJxU4'
DINGTALK_API_BASE = 'https://oapi.dingtalk.com'
GET_TOKEN_URL = f'{DINGTALK_API_BASE}/gettoken?appkey={APP_KEY}&appsecret={APP_SECRET}'
APPROVAL_RECORD_URL = f'{DINGTALK_API_BASE}/topapi/processinstance/listids'

KQ_PROCESS_CODE = 'PROC-6B32AF04-FE29-4DAC-9B99-2D4D58183DC9' #外包考勤的OA审批
BX_PROCESS_CODE = '' #外包报销的OA审批
# 获取 AccessToken
def get_access_token():
    response = requests.get(GET_TOKEN_URL)
    if response.status_code == 200:
        return response.json()['access_token']
    else:
        raise Exception('Failed to get access token')


# 获取审批记录 ID 列表
def get_approval_record_ids(token, process_code, start_time, end_time):
    data = {
        'process_code': process_code,  # 替换为你的审批流程代码
        'start_time': start_time,
        'end_time': end_time,
        'size': 10,  # 每次请求的记录数，根据需要调整
        'cursor': 0  # 分页游标
    }
    headers = {'Content-Type': 'application/json'}
    response = requests.post(f"{APPROVAL_RECORD_URL}?access_token={token}", json=data, headers=headers)
    if response.status_code == 200:
        return response.json()['result']['list']
    else:
        raise Exception('Failed to get approval record IDs')

# 获取单个审批记录的详细数据
def get_approval_record_details(token, record_id):
    details_url = f'{DINGTALK_API_BASE}/topapi/processinstance/get?access_token={token}'
    data = {'process_instance_id': record_id}
    response = requests.post(details_url, json=data)
    if response.status_code == 200:
        #print(response.json())
        return response.json()['process_instance']
    else:
        raise Exception('Failed to get approval record details')

# 修改 export_to_csv 函数，以适应新的数据结构
# ...


# 获取审批记录
def get_approval_records(token, start_time, end_time):
    data = {
        'process_code': 'PROC-6B32AF04-FE29-4DAC-9B99-2D4D58183DC9',  # 替换为你的审批流程代码
        'start_time': start_time,
        'end_time': end_time,
        'access_token': token
    }
    response = requests.post(APPROVAL_RECORD_URL, data=data)
    if response.status_code == 200:
        return response.json()['result']['list']
    else:
        raise Exception('Failed to get approval records')

# 导出审批记录到 CSV
def export_to_csv(records, filename='approval_records.csv'):
    with open(filename, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        # 写入标题行
        writer.writerow(['审批ID', '标题', '状态', '发起时间', '结束时间'])
        for record in records:
            # 写入每条审批记录
            form_values = record['form_component_values']
            # 提取表单组件的值
            form_values_str = '; '.join([f"{value['name']}: {value['value']}" for value in form_values])


            # 写入每条审批记录
            writer.writerow([
                record.get('business_id', '未知'),
                record.get('title', '未知'),
                record.get('status', '未知'),
                record.get('create_time', '未知'),
                record.get('finish_time', '未知'),
                form_values_str
            ])
            #writer.writerow([record['process_instance_id'], record['title'], record['status'], record['create_time'], record['finish_time']])

def parse_form_values(records):
    all_fields = set()
    parsed_records = []

    # 解析每个记录的表单值
    for record in records:
        form_values = record['form_component_values']
        parsed_record = {val['name']: val['value'] for val in form_values}
        parsed_records.append(parsed_record)
        all_fields.update(parsed_record.keys())

    return all_fields, parsed_records

def export_to_excel(records, all_fields, filename='approval_records.xlsx'):
    # 准备数据，确保每条记录都有所有字段
    prepared_data = []
    for record in records:
        prepared_data.append({field: record.get(field, '') for field in all_fields})
    
    # 创建 DataFrame 并导出到 Excel
    df = pd.DataFrame(prepared_data)
    df.to_excel(filename, index=False)


# 主逻辑
if __name__ == '__main__':
    token = get_access_token()
    end_time = int(time.time() * 1000)  # 当前时间戳（毫秒）
    start_time = end_time - 30 * 24 * 60 * 60 * 1000  # 过去30天

    record_ids = get_approval_record_ids(token, KQ_PROCESS_CODE, start_time, end_time)
    detailed_records = [get_approval_record_details(token, record_id) for record_id in record_ids]

    #export_to_csv(detailed_records)
    
    all_fields, parsed_records = parse_form_values(detailed_records)
    
    export_to_excel(parsed_records, all_fields)
