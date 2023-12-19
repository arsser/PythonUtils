import requests
import json

# 钉钉开发平台的应用唯一标识
APP_KEY = "xxxx"
# 钉钉开发平台的应用密钥
APP_SECRET = "xxxx"

# 获取access_token
def get_access_token():
    url = "https://oapi.dingtalk.com/gettoken"
    params = {
        "appkey": APP_KEY,
        "appsecret": APP_SECRET
    }
    response = requests.post(url, params=params)
    response.raise_for_status()
    data = json.loads(response.text)
    return data["access_token"]

# 获取审批记录
def get_approval_records(access_token, start_time, end_time):
    url = "https://oapi.dingtalk.com/topapi/approval/listbyuser"
    params = {
        "access_token": access_token,
        "start_time": start_time,
        "end_time": end_time
    }
    response = requests.post(url, params=params)
    response.raise_for_status()
    data = json.loads(response.text)
    return data

# 导出审批记录
def export_approval_records(approval_records, filename):
    with open(filename, "w", encoding="utf-8") as f:
        f.write("审批记录\n")
        for approval_record in approval_records:
            f.write("审批单号：" + approval_record["process_code"] + "\n")
            f.write("审批人：" + approval_record["user_id"] + "\n")
            f.write("审批类型：" + approval_record["process_type"] + "\n")
            f.write("审批状态：" + approval_record["status"] + "\n")
            f.write("审批时间：" + approval_record["create_time"] + "\n")
            f.write("审批意见：" + approval_record["comment"] + "\n")

if __name__ == "__main__":
    # 获取access_token
    access_token = get_access_token()

    # 获取审批记录
    start_time = "2023-07-20"
    end_time = "2023-07-30"
    approval_records = get_approval_records(access_token, start_time, end_time)

    # 导出审批记录
    filename = "approval_records.csv"
    export_approval_records(approval_records, filename)
