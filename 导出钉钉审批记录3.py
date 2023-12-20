import requests
import csv
import time
import pandas as pd
from datetime import datetime,timedelta

# 常量
APP_KEY = 'dingl5qlh2s1ddksf5ru'
APP_SECRET = 'mjr6mb1TpfMx1Q2W7IfSVvrTsD855sSMQcL_XNZQA_HVhCYBKT-FDm8VPqGpJxU4'
DINGTALK_API_BASE = 'https://oapi.dingtalk.com'
GET_TOKEN_URL = f'{DINGTALK_API_BASE}/gettoken?appkey={APP_KEY}&appsecret={APP_SECRET}'
APPROVAL_RECORD_URL = f'{DINGTALK_API_BASE}/topapi/processinstance/listids'

KQ_PROCESS_CODE = 'PROC-6B32AF04-FE29-4DAC-9B99-2D4D58183DC9' #外包考勤的OA审批,v1旧版版编辑模板时显示在url
BX_PROCESS_CODE = 'PROC-B193D2E7-4B1B-43EA-8581-CEF423B71C1F' #外包报销的OA审批

CORP_INFO_FILE= fr'\\10.12.21.65\Share\外包费用\新流程\2023-11\Step1-考勤数据\【HR】外包人员表.xlsx'

# 获取 AccessToken
def get_access_token():
    response = requests.get(GET_TOKEN_URL)
    if response.status_code == 200:
        return response.json()['access_token']
    else:
        raise Exception('Failed to get access token')

def datetime_to_timestamp(dt):
    return int(time.mktime(dt.timetuple()) * 1000)

# 获取审批记录 ID 列表
def get_all_approval_record_ids(token, process_code, start_time, end_time):
    all_record_ids = []
    size = 20
    next_cursor = 0
    #has_more = True
    # 在调用 API 之前，将 datetime 对象转换为时间戳
    start_timestamp = datetime_to_timestamp(start_time)
    end_timestamp = datetime_to_timestamp(end_time)
    while True:
        data = {
            'process_code': process_code,
            'start_time': start_timestamp,
            'end_time': end_timestamp,
            'size': size,
            'cursor': next_cursor
        }
        response = requests.post(f"{APPROVAL_RECORD_URL}?access_token={token}", json=data)
        
        response_data = response.json()

        # 检查 errcode
        if response_data['errcode'] != 0:
            # 打印或抛出异常信息
            error_msg = response_data.get('errmsg', 'Unknown error')
            raise Exception(f"获取{process_code}的审批列表 API Error: {error_msg}")
        
        if response.status_code == 200:
            result = response_data['result']
            all_record_ids.extend(result['list'])            
            if 'next_cursor' in result and result['next_cursor'] != 0:
                next_cursor = result['next_cursor']
            else:
                break  # 没有更多数据或next_cursor为0            
        else:
            raise Exception('Failed to get approval record IDs')
        break
    
    print(f'获取{process_code} 审批记录条数：{len(all_record_ids)}')
    return all_record_ids

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


def extract_creator_from_title(title):
    # 分割标题并提取创建人的名字
    parts = title.split('提交的')
    creator = parts[0] if len(parts) > 1 else None
    return creator

def prepare_data_for_excel(records):
    prepared_data = []
    # 遍历每条记录
    for record in records:
        # 解析表单组件的值
        form_values = {comp['name']: comp['value'] for comp in record['form_component_values']}

        if record['status'] == 'COMPLETED' and record['result'] == 'agree':
            # 整合其他字段
            combined_record = {
                '审批ID': record['business_id'],
                '审批结果': '审批通过',
                '审批状态': '已结束',
                '标题': record['title'],
                '创建人': extract_creator_from_title(record['title']),
                '创建人部门': record['originator_dept_name'],
                '发起时间': record['create_time'],
                '结束时间': record['finish_time'],
                **form_values  # 将表单值添加到记录中
            }
            prepared_data.append(combined_record)
    return prepared_data


def generate_filename():
 # 获取当前日期和时间，并格式化为字符串，格式为"年-月-日 时-分-秒"
    current_datetime_str = datetime.now().strftime("%Y-%m-%d-%H%M%S")

    # 设置文件名为当前日期时间
    filename = f'外包出勤审批_{current_datetime_str}.xlsx'
      
    filename='外包出勤审批_2023-12-20-114202.xlsx'
    return filename

def get_approval_data(token, process_code, start_time, end_time):
    record_ids = get_all_approval_record_ids(token, process_code, start_time, end_time)
    detailed_records = [get_approval_record_details(token, record_id) for record_id in record_ids[:3]]
    return detailed_records

def number_to_chinese(num):
    num_to_chinese = {
        1: '一',
        2: '二',
        3: '三',
        4: '四',
        5: '五',
        6: '六',
        7: '七',
        8: '八',
        9: '九',
        10: '十',
        11: '十一',
        12: '十二'
    }
    return num_to_chinese.get(num, '')


def get_kaoqin_ding_data(token, process_code,  start_time, end_time, kaoqin_month, excel_file):
    #获取考勤的钉钉审批记录，kaoqin_month表示要获取哪个月份的考勤记录;
    record_ids = get_all_approval_record_ids(token, process_code, start_time, end_time)
    detailed_records = [get_approval_record_details(token, record_id) for record_id in record_ids]
    prepared_records =  pd.DataFrame(prepare_data_for_excel(detailed_records))
    # 过滤出 '出勤月份' 为 kaoqin_month 的数据
    kaoqin_mon_chin = f'{number_to_chinese(kaoqin_month)}月'
    df_filtered = prepared_records[prepared_records['出勤月份'] == kaoqin_mon_chin]  
    df_filtered.to_excel(excel_file, index=False)
    return 0

def get_baoxiao_ding_data(token, process_code,  start_time, end_time, baoxiao_month, excel_file):
    return 0


def merge_kaoqin_hr_excel(kaoqin_excel, hr_excel, hr_excel_sheet, merged_excel):
    #把考勤的excel和hr的人员信息excel进行合并,该excel按月份建立sheet，记录当月的外包人员信息；
    # 获取上一个月的月份
    now = datetime.now()
    first_day_of_current_month = datetime(now.year, now.month, 1)
    last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)
    previous_month_str = last_day_of_previous_month.strftime("%Y-%m")

    hr_excel_sheet = previous_month_str

    # 读取审批记录表
    df_approval_records = pd.read_excel(kaoqin_excel)

    # 读取公司信息表的上个月数据
    df_company_info = pd.read_excel(CORP_INFO_FILE, sheet_name=hr_excel_sheet,usecols=['姓名','供应商','入职日期','离职日期'])
     
    # 使用 '姓名'（或其他合适的键）进行合并
    # 假设两个表都有一个共同的列 '姓名' 用于合并
    df_merged = pd.merge(df_filtered, df_company_info, left_on='创建人', right_on='姓名', how='outer')
    df_merged['入职日期'] = pd.to_datetime(df_merged['入职日期']).dt.date
    df_merged['离职日期'] = pd.to_datetime(df_merged['离职日期']).dt.date
    # 删除重复的列（例如，删除 "姓名" 列）
    #df_merged.drop(columns=['姓名'], inplace=True)
    # 指定新的列顺序
    new_column_order = ['供应商', '审批ID', '创建人','出勤月份', '入职日期','离职日期','姓名','本月缺勤天数','具体缺勤日期','本月加班天数','具体加班日期',
                        '加班原因','备注','审批状态','审批结果','标题','发起时间','结束时间','创建人部门','部门信息']

    # 重新排列列
    df_merged = df_merged[new_column_order]
    # 将合并后的数据导出到新的 Excel 文件
    df_merged.to_excel(merge_excel, index=False)

def merge_baoxiao_hr_excel(hr_excel, sheet_name, merge_excel):
    #把报销的excel和hr的人员信息excel进行合并
    return 0

def split_kaoqin_excel(excel):
    return 0

def split_baoxiao_excel():
    return 0

def mailto_supplier():
    return 0

# 主逻辑
if __name__ == '__main__':     
    # 获取当前日期和时间，并格式化为字符串，格式为"年-月-日-时分秒"
    current_datetime_str = datetime.now().strftime("%Y-%m-%d-%H%M%S")
    kaoqin_excel =       fr'\\10.12.21.65\share\外包费用\新流程\考勤报销数据\外包出勤审批_{current_datetime_str}.xlsx'
    kaoqin_merged_excel = fr'\\10.12.21.65\share\外包费用\新流程\考勤报销数据\外包出勤审批_merge_{current_datetime_str}.xlsx'
    hr_excel =           fr'\\10.12.21.65\share\外包费用\新流程\【HR】外包人员表.xlsx'
    #kaoqin_excel_file = f'\\10.12.21.65\share\外包费用\新流程\考勤报销数据\外包出勤审批_2023-12-20-114202.xlsx'
    baoxiao_excel = fr'\\10.12.21.65\share\外包费用\新流程\考勤报销数据\外包报销审批_{current_datetime_str}.xlsx'

    token = get_access_token()

    # 获取当前时间
    now = datetime.now()
    # 计算当前月份的第五天
    end_time = datetime(now.year, now.month, 5)
    # 计算上个月的第一天, 如果当前是1月，则上个月是去年的12月
    if now.month == 1:
        start_time = datetime(now.year - 1, 12, 1)
    else:
        start_time = datetime(now.year, now.month - 1, 1)    
    
    end_time = datetime(now.year, now.month, 10)


    #1. 获取考勤,报销的审批记录，并保存为excel文件；
    get_kaoqin_ding_data(token, KQ_PROCESS_CODE, start_time, end_time, start_time.month, kaoqin_excel)
    #get_baoxiao_ding_data(token,BX_PROCESS_CODE, start_time, end_time, baoxiao_excel)
    
    #2. merge
    merge_kaoqin_hr_excel(kaoqin_excel, hr_excel, '11', kaoqin_merged_excel)


    #3. 分拆
    split_kaoqin_excel(kaoqin_merged_excel)


    #4. 生成邮件草稿

    #4. 分拆报销的审批excel文件；
    split_baoxiao_excel()
    #5. make email draft；
    mailto_supplier()

    #1. 生成文件名

    #2. 获取审批记录 ID 和详细信息

    #3. 准备数据并导出到 Excel

    #4. 读取审批记录和公司信息表，进行数据过滤和合并

    #5. 调整 DataFrame 的列顺序并导出

    record_ids = get_all_approval_record_ids(token, KQ_PROCESS_CODE, start_time, end_time)
    detailed_records = [get_approval_record_details(token, record_id) for record_id in record_ids]

    prepared_records = prepare_data_for_excel(detailed_records)
    

    export_to_excel(prepared_records,filename)

    # 获取上一个月的月份
    now = datetime.now()
    first_day_of_current_month = datetime(now.year, now.month, 1)
    last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)
    previous_month_str = last_day_of_previous_month.strftime("%Y-%m")
    # 读取审批记录表
    df_approval_records = pd.read_excel(filename)

    # 读取公司信息表的上个月数据
    df_company_info = pd.read_excel(CORP_INFO_FILE, sheet_name=previous_month_str,usecols=['姓名','供应商','入职日期','离职日期'])
    # 过滤出 '出勤月份' 为 '11月' 的数据
    df_filtered = df_approval_records[df_approval_records['出勤月份'] == '十一月']    

    # 使用 '姓名'（或其他合适的键）进行合并
    # 假设两个表都有一个共同的列 '姓名' 用于合并
    df_merged = pd.merge(df_filtered, df_company_info, left_on='创建人', right_on='姓名', how='outer')
    df_merged['入职日期'] = pd.to_datetime(df_merged['入职日期']).dt.date
    df_merged['离职日期'] = pd.to_datetime(df_merged['离职日期']).dt.date
    # 删除重复的列（例如，删除 "姓名" 列）
    #df_merged.drop(columns=['姓名'], inplace=True)
    # 指定新的列顺序
    new_column_order = ['供应商', '审批ID', '创建人','出勤月份', '入职日期','离职日期','姓名','本月缺勤天数','具体缺勤日期','本月加班天数','具体加班日期',
                        '加班原因','备注','审批状态','审批结果','标题','发起时间','结束时间','创建人部门','部门信息']

    # 重新排列列
    df_merged = df_merged[new_column_order]
    # 将合并后的数据导出到新的 Excel 文件
    df_merged.to_excel(f"{filename}.1.xlsx", index=False)
   
    df = df_merged
    month = '11'
    # 拆分数据
    for value in df['供应商'].unique():
        df_subset = df[df['供应商'] == value]
        df_subset.to_excel(rf"\\10.12.21.65\share\外包费用\新流程\chai_{value}_{month}月出勤.xlsx", index=False)


