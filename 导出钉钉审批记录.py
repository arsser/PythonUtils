import requests
import csv
import time

# 常量
APP_KEY = 'dingl5qlh2s1ddksf5ru'
APP_SECRET = 'mjr6mb1TpfMx1Q2W7IfSVvrTsD855sSMQcL_XNZQA_HVhCYBKT-FDm8VPqGpJxU4'
DINGTALK_API_BASE = 'https://oapi.dingtalk.com'
GET_TOKEN_URL = f'{DINGTALK_API_BASE}/gettoken?appkey={APP_KEY}&appsecret={APP_SECRET}'
APPROVAL_RECORD_URL = f'{DINGTALK_API_BASE}/topapi/processinstance/listids'

# 获取 AccessToken
def get_access_token():
    response = requests.get(GET_TOKEN_URL)
    if response.status_code == 200:
        return response.json()['access_token']
    else:
        raise Exception('Failed to get access token')

# 获取审批记录
def get_approval_records(token, start_time, end_time):
    data = {
        'process_code': 'your_process_code',  # 替换为你的审批流程代码
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
            writer.writerow([record['process_instance_id'], record['title'], record['status'], record['create_time'], record['finish_time']])

# 主逻辑
if __name__ == '__main__':
    token = get_access_token()
    # 示例：获取当前时间至过去30天的审批记录
    end_time = int(time.time() * 1000)  # 当前时间戳（毫秒）
    start_time = end_time - 30 * 24 * 60 * 60 * 1000  # 过去30天
    records = get_approval_records(token, start_time, end_time)
    export_to_csv(records)
