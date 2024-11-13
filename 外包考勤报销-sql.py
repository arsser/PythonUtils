#外包考勤报销，主要功能：
#1.从钉钉获取外包考勤、报销的审批记录；
#2.对审批记录做处理，包括：筛选上个月的数据，和hr的名单做merge
#3.按照供应商名单，将拆分后的文件生产邮件草稿

import requests
import csv
import time
import pandas as pd
from datetime import datetime,timedelta
import sys
import os
import win32com.client as win32
from openpyxl import load_workbook
import json
import glob
import pyodbc


def insert_kaoqin_records(records):
    """
    将钉钉考勤记录插入到外包考勤表，如果记录已存在则跳过
    
    Args:
        records (list): 包含考勤记录信息的列表
    """
    insert_sql = """INSERT INTO dbo.外包考勤 ([process_instance_id],[审批ID], [审批结果], [审批状态], [标题], [创建人], [创建人部门], [发起时间], [结束时间], [出勤月份], [部门信息], [本月加班天数], [具体加班日期], [加班原因], [本月缺勤天数], [具体缺勤日期], [备注]) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"""
    check_sql = """SELECT COUNT(1) FROM dbo.外包考勤 WHERE [process_instance_id] = ?"""
     
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            inserted_count = 0
            skipped_count = 0
            
            for record in records:
                # 检查记录是否已存在
                cursor.execute(check_sql, (record.get('process_instance_id'),))
                exists = cursor.fetchone()[0] > 0
                
                if exists:
                    print(f"记录 {record.get('process_instance_id')} 已存在，跳过")
                    skipped_count += 1
                    continue
                    
                # 插入新记录
                cursor.execute(insert_sql, (
                    record.get('process_instance_id'),
                    record.get('审批ID'),
                    record.get('审批结果'),
                    record.get('审批状态'),
                    record.get('标题'),
                    record.get('创建人'),
                    record.get('创建人部门'),
                    record.get('发起时间'),
                    record.get('结束时间'),
                    record.get('出勤月份'),
                    record.get('部门信息'),
                    record.get('本月加班天数'),
                    record.get('具体加班日期'),
                    record.get('加班原因'),
                    record.get('本月缺勤天数'),
                    record.get('具体缺勤日期'),
                    record.get('备注')                    
                ))
                inserted_count += 1
                print(f"插入记录: {record.get('process_instance_id')}") 
                
            conn.commit()
            print(f"成功插入 {inserted_count} 条考勤记录，跳过 {skipped_count} 条重复记录")
            
    except Exception as e:
        print(f"插入考勤记录失败: {str(e)}")
        print(f"SQL: {insert_sql}")
        raise

def insert_baoxiao_records(records):
    """
    将报销记录插入数据库
    
    Args:
        records (list): 包含报销记录信息的列表
    """
    insert_sql = """
        INSERT INTO dbo.外包报销 (
            [process_instance_id], [审批ID], [审批结果], [审批状态], 
            [标题], [创建人], [创建人部门], [发起时间], [结束时间],
            [报销归属月份], [报销总金额（元）], [表格], [其他附件], [备注]
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """
    check_sql = "SELECT COUNT(1) FROM dbo.外包报销 WHERE [process_instance_id] = ?"
    
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            inserted_count = 0
            skipped_count = 0
            
            for record in records:
                # 检查记录是否已存在
                cursor.execute(check_sql, (record.get('process_instance_id'),))
                exists = cursor.fetchone()[0] > 0
                
                if exists:
                    print(f"记录 {record.get('process_instance_id')} 已存在，跳过")
                    skipped_count += 1
                    continue
                
                try:
                    # 准备参数
                    params = (
                        record.get('process_instance_id'),
                        record.get('审批ID'),
                        record.get('审批结果'),
                        record.get('审批状态'),
                        record.get('标题'),
                        record.get('创建人'),
                        record.get('创建人部门'),
                        record.get('发起时间'),
                        record.get('结束时间'),
                        record.get('报销归属月份'),
                        record.get('报销总金额（元）'),
                        record.get('表格'),
                        record.get('其他附件'),
                        record.get('备注')
                    )
                    
                    cursor.execute(insert_sql, params)
                    inserted_count += 1
                    print(f"插入记录: {record.get('process_instance_id')}")
                    
                except Exception as e:
                    # 构建实际执行的SQL语句用于调试
                    param_values = [f"'{str(p)}'" if p is not None else 'NULL' for p in params]
                    actual_sql = insert_sql.replace('?', '{}').format(*param_values)
                    
                    print(f"插入记录失败: {record.get('process_instance_id')}")
                    print(f"错误信息: {str(e)}")
                    print(f"SQL语句: {actual_sql}")
                    print("参数值:")
                    field_names = [
                        "process_instance_id", "审批ID", "审批结果", "审批状态",
                        "标题", "创建人", "创建人部门", "发起时间", "结束时间",
                        "报销归属月份", "报销总金额（元）", "表格", "其他附件", "备注"
                    ]
                    for name, value in zip(field_names, params):
                        print(f"  {name}: {value} (类型: {type(value)})")
                    continue
            
            conn.commit()
            print(f"成功插入 {inserted_count} 条报销记录，跳过 {skipped_count} 条重复记录")
            
    except Exception as e:
        print(f"数据库操作失败: {str(e)}")
        raise


# 常量
APP_KEY = 'dingl5qlh2s1ddksf5ru'
APP_SECRET = 'mjr6mb1TpfMx1Q2W7IfSVvrTsD855sSMQcL_XNZQA_HVhCYBKT-FDm8VPqGpJxU4'
DINGTALK_API_BASE = 'https://oapi.dingtalk.com'
GET_TOKEN_URL = f'{DINGTALK_API_BASE}/gettoken?appkey={APP_KEY}&appsecret={APP_SECRET}'
APPROVAL_RECORD_URL = f'{DINGTALK_API_BASE}/topapi/processinstance/listids'

KQ_PROCESS_CODE = 'PROC-6B32AF04-FE29-4DAC-9B99-2D4D58183DC9' #外包考勤的OA审批,v1旧版版编辑模板时显示在url
BX_PROCESS_CODE = 'PROC-B193D2E7-4B1B-43EA-8581-CEF423B71C1F' #外包报销的OA审批

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
    return all_record_ids

# 获取单个审批记录的详细数据
def get_approval_record_details(token, process_instance_id):
    details_url = f'{DINGTALK_API_BASE}/topapi/processinstance/get?access_token={token}'
    data = {'process_instance_id': process_instance_id}
    response = requests.post(details_url, json=data)
    if response.status_code == 200:
        process_instance = response.json()['process_instance']
        # 添加 process_instance_id
        process_instance['process_instance_id'] = process_instance_id
        return process_instance
    else:
        raise Exception('Failed to get approval record details')

# 分割标题并提取创建人的名字
def extract_creator_from_title(title):   
    parts = title.split('提交的')
    creator = parts[0] if len(parts) > 1 else None
    return creator

def prepare_data_for_excel(records):
    prepared_data = []
    # 遍历每条记录
    for record in records: 
        #form_values = {comp['name']: comp['value'] for comp in record['form_component_values']}
        # 解析表单组件的值, 在字典推导式中处理 'null' 值
        form_values = {
            comp['name']: ('' if comp['value'] == 'null' else comp['value']) 
            for comp in record['form_component_values']
        }
        #if record['status'] == 'COMPLETED' and record['result'] == 'agree':
        # 整合其他字段
        combined_record = {  
            'process_instance_id': record['process_instance_id'],           
            '审批ID': record['business_id'],
            '审批结果': record['result'],
            '审批状态': record['status'],
            '标题': record['title'],
            '创建人': extract_creator_from_title(record['title']),
            '创建人部门': record['originator_dept_name'],
            '发起时间': record['create_time'],
            '结束时间': record.get('finish_time', None),            
            **form_values  # 将表单值添加到记录中
        }
        prepared_data.append(combined_record)
    return prepared_data

#处理钉钉审批单中有嵌套表格的情况，并将表格明细按行展开
def parse_table_field_and_expand(record, table_data, form_values):
    expanded_records = []
    for row in table_data:
        parsed_row = {}
        for cell in row['rowValue']:
            if cell['componentType'] == 'DDAttachment':  # 特殊处理附件字段
                attachment_info = '; '.join([f"{att['fileName']} ({att['fileSize']} bytes)" for att in cell['value']])
                parsed_row[cell['label']] = attachment_info
            else:
                parsed_row[cell['label']] = cell['value']

        # 将表格的每一行作为一个独立的记录，包含主记录的其他信息
        expanded_record = {
            '审批ID': record['business_id'],
            '审批结果': '审批通过',
            '审批状态': '已结束',
            '标题': record['title'],
            '创建人': extract_creator_from_title(record['title']),
            '创建人部门': record['originator_dept_name'],
            '发起时间': record['create_time'],
            '结束时间': record['finish_time'],
            **form_values,  # 包含除表格外的其他表单值
            **parsed_row  # 包含当前表格行的信息
        }
        expanded_records.append(expanded_record)

    return expanded_records

def prepare_data_for_excel_forbx_expand(records):
    prepared_data = []
    for record in records:
        form_values = {comp['name']: comp['value'] for comp in record['form_component_values'] if comp['component_type'] != 'TableField'}
        if record['status'] == 'COMPLETED' and record['result'] == 'agree':
            # 对于包含表格的记录，单独处理
            for comp in record['form_component_values']:
                if comp['component_type'] == 'TableField':
                    table_data = json.loads(comp['value'])
                    prepared_data.extend(parse_table_field_and_expand(record, table_data, form_values))
                else:
                    combined_record = {
                        '审批ID': record['business_id'],
                        '审批结果': '审批通过',
                        '审批状态': '已结束',
                        '标题': record['title'],
                        '创建人': extract_creator_from_title(record['title']),
                        '创建人部门': record['originator_dept_name'],
                        '发起时间': record['create_time'],
                        '结束时间': record['finish_time'],
                        **form_values
                    }
                    prepared_data.append(combined_record)
    return prepared_data

#处理钉钉审批单中有嵌套表格的情况
# def parse_table_field_old(table_data):
    parsed_table = []
    for row in table_data:
        parsed_row = {}
        for cell in row['rowValue']:
            if cell['componentType'] == 'DDAttachment':  # 特殊处理附件字段
                attachment_info = '; '.join([f"{att['fileName']} ({att['fileSize']} bytes)" for att in cell['value']])
                parsed_row[cell['label']] = attachment_info
            else:
                parsed_row[cell['label']] = cell['value']
        parsed_table.append(parsed_row)
    return parsed_table

def parse_table_field(table_data):
    parsed_table = []
    for row in table_data:
        parsed_row = []
        for cell in row['rowValue']:
            if cell['componentType'] == 'DDAttachment':  # 特殊处理附件字段
                try:
                    # 判断value是否为字符串，如果是则需要解析JSON
                    attachment_list = cell['value']
                    if isinstance(attachment_list, str):
                        try:
                            attachment_list = json.loads(attachment_list)
                        except json.JSONDecodeError as e:
                            print(f"JSON解析失败: {e}")
                            parsed_row.append(f"{cell['label']}: 附件数据解析失败")
                            continue
                    
                    # 处理附件信息
                    attachment_info = '; '.join([
                        f"{att['fileName']} ({att['fileSize']} bytes)" 
                        for att in attachment_list
                    ])
                    parsed_row.append(f"{cell['label']}: {attachment_info}")
                except Exception as e:
                    print(f"处理附件字段失败: {e}")
                    print(f"问题数据: {cell['value']}")
                    parsed_row.append(f"{cell['label']}: 附件处理失败")
            else:
                parsed_row.append(f"{cell['label']}: {cell['value']}")
        parsed_table.append('\n'.join(parsed_row))
    return '\n'.join(parsed_table)  # 每个表格行之间加上换行符


def extract_attachment_filenames(attachment_data):
    if not attachment_data:  # 检查数据是否为空或None
        return ''

    try:
        attachments = json.loads(attachment_data)
        if attachments is None:  # 检查解析结果是否为None
            return ''
        file_names = [att['fileName'] for att in attachments]
        return ', '.join(file_names)
    except json.JSONDecodeError:
        # 如果解析出错，返回原始数据或空字符串
        return ''

# 其余部分保持不变
def prepare_data_for_excel_forbx(records):
    prepared_data = []
    for record in records:
        form_values = {}
        for comp in record['form_component_values']:
            if comp['component_type'] == 'DDAttachment':
                # 特殊处理附件字段
                form_values[comp['name']] = extract_attachment_filenames(comp['value'])
            else:
                if comp['component_type'] == 'TableField':
                    # 解析表格
                    table_data = json.loads(comp['value'])
                    form_values[comp['name']] = parse_table_field(table_data)
                else:
                    value = comp['value']
                    if value == 'null':  # 将字符串'null'替换为空字符串
                        value = ''
                    form_values[comp['name']] = value

        #if record['status'] == 'COMPLETED' and record['result'] == 'agree':
        combined_record = {
            'process_instance_id': record['process_instance_id'], 
            '审批ID': record['business_id'],
            '审批结果': record['result'],
            '审批状态': record['status'],
            '标题': record['title'],
            '创建人': extract_creator_from_title(record['title']),
            '创建人部门': record['originator_dept_name'],
            '发起时间': record['create_time'],
            '结束时间': record.get('finish_time', None),
            **form_values
        }
        prepared_data.append(combined_record)
    return prepared_data

#获取考勤的钉钉审批记录
def get_kaoqin_ding_data(token, process_code,  start_time, end_time):
    record_ids = get_all_approval_record_ids(token, process_code, start_time, end_time)
    print(f'{start_time} - {end_time} 考勤审批记录条数：{len(record_ids)}')
    detailed_records = [get_approval_record_details(token, record_id) for record_id in record_ids]
    records = prepare_data_for_excel(detailed_records)
    return records

def get_baoxiao_ding_data(token, process_code,  start_time, end_time):
    record_ids = get_all_approval_record_ids(token, process_code, start_time, end_time)
    print(f'{start_time} - {end_time} 获取报销审批记录条数：{len(record_ids)}')
    detailed_records = [get_approval_record_details(token, record_id) for record_id in record_ids]
    records = prepare_data_for_excel_forbx(detailed_records)
    return records


def get_db_connection():
    """创建数据库连接"""
    conn_str = (
        'DRIVER={SQL Server};'
        'SERVER=powerbi01;'
        'DATABASE=DigiYitu;'
        'UID=partner;'
        'PWD=galaxy'
    )
    return pyodbc.connect(conn_str)

def get_pending_approval_records(table_name = "外包考勤"):
    """
    获取审批中的记录ID
    
    Args:
        isKaoqin: True表示获取考勤记录，False表示获取报销记录
        
    Returns:
        array: 审批中的record_id数组
    """
    #table_name = "外包考勤" if isKaoqin else "外包报销"
    
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        sql = f"""
            SELECT [process_instance_id]  
            FROM {table_name}
            WHERE 审批状态 = 'RUNNING'
        """
        
        cursor.execute(sql)
        records = [str(row[0]) for row in cursor.fetchall()]
        
        cursor.close()
        conn.close()
        
        return records
        
    except Exception as e:
        print(f"数据库查询出错: {str(e)}")
        return []

def get_max_create_time(table_name = "外包考勤"):
    """
    获取数据库中最新的发起时间
    
    Args:
        isKaoqin (bool): True表示考勤数据，False表示报销数据
    
    Returns:
        datetime: 最新的发起时间，如果没有记录则返回None
    """
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            
            sql = f"""
                SELECT MAX([发起时间]) 
                FROM {table_name}                
            """                
            cursor.execute(sql)
            max_time = cursor.fetchone()[0]
            print(f"{table_name}数据最后发起时间: {max_time}")
            return max_time
            
    except Exception as e:
        print(f"查询最大发起时间失败: {str(e)}")
        return None

def update_kq_record_to_db(record_details):
    """
    更新钉钉考勤记录到SQL Server数据库
    
    Args:
        record_details (dict): 包含考勤记录详细信息的字典
    
    Returns:
        bool: 更新成功返回True，失败抛出异常
    """
    sql = """UPDATE dbo.[外包考勤] SET [审批结果] = ?, [审批状态] = ?, [结束时间] = ?, [出勤月份] = ?, [部门信息] = ?, [本月加班天数] = ?, [具体加班日期] = ?, [加班原因] = ?, [本月缺勤天数] = ?, [具体缺勤日期] = ?, [备注] = ? WHERE [process_instance_id] = ?"""
    
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            
            # 从record_details中提取表单值
            form_values = {
                comp['name']: ('' if comp['value'] == 'null' else comp['value']) 
                for comp in record_details['form_component_values']
            }
            # 准备SQL参数
            params = (
                record_details['result'],
                record_details['status'],       
                record_details.get('finish_time', None),
                form_values.get('出勤月份', ''),
                form_values.get('部门信息', ''),
                form_values.get('本月加班天数', 0),
                form_values.get('具体加班日期', ''),
                form_values.get('加班原因', ''),
                form_values.get('本月缺勤天数', 0),
                form_values.get('具体缺勤日期', ''),
                form_values.get('备注', ''),
                record_details['process_instance_id']
            )
            cursor.execute(sql, params)            
            rows_affected = cursor.rowcount
            conn.commit()            
            if rows_affected == 0:
                print(f"警告: 没有找到process_instance_id为 {record_details['process_instance_id']} 的记录")
            else:
                print(f"成功更新 {rows_affected} 条考勤记录，process_instance_id为 {record_details['process_instance_id']}")
            
            return True
            
    except Exception as e:
        # 构建实际执行的SQL语句用于调试
        param_values = [f"'{str(p)}'" if p is not None else 'NULL' for p in params]
        actual_sql = sql.replace('?', '{}').format(*param_values)
        print(f"更新数据库失败: {str(e)}")
        print(f"SQL语句: {actual_sql}")
        raise

def update_bx_record_to_db(record_details):
    """
    更新钉钉报销记录到SQL Server数据库
    
    Args:
        record_details (dict): 包含报销记录详细信息的字典
    
    Returns:
        bool: 更新成功返回True，失败抛出异常
    """
    sql = """UPDATE dbo.[外包报销] SET [审批结果] = ?, [审批状态] = ?, [标题] = ?, [创建人] = ?, [创建人部门] = ?, [发起时间] = ?, [结束时间] = ?, [报销归属月份] = ?, [报销总金额（元）] = ?, [表格] = ?, [其他附件] = ?, [备注] = ? WHERE [process_instance_id] = ?"""
    
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            
            # 从record_details中提取表单值
            form_values = {}
            for comp in record_details['form_component_values']:
                if comp['component_type'] == 'DDAttachment':
                    # 特殊处理附件字段
                    form_values[comp['name']] = extract_attachment_filenames(comp['value'])
                elif comp['component_type'] == 'TableField':
                    # 解析表格
                    table_data = json.loads(comp['value'])
                    form_values[comp['name']] = parse_table_field(table_data)
                else:
                    value = comp['value']
                    form_values[comp['name']] = '' if value == 'null' else value
            
            # 准备参数
            params = (
                record_details['result'],
                record_details['status'],
                record_details['title'],
                extract_creator_from_title(record_details['title']),
                record_details['originator_dept_name'],
                record_details['create_time'],
                record_details.get('finish_time', None),
                form_values.get('报销归属月份', ''),
                form_values.get('报销总金额（元）', 0),
                form_values.get('表格', ''),
                form_values.get('其他附件', ''),
                form_values.get('备注', ''),
                record_details['process_instance_id']
            )
            
            # 执行更新
            try:
                cursor.execute(sql, params)
                rows_affected = cursor.rowcount
                conn.commit()
                
                if rows_affected == 0:
                    print(f"警告: 没有找到process_instance_id为 {record_details['process_instance_id']} 的记录")
                else:
                    print(f"成功更新 {rows_affected} 条报销记录，process_instance_id为 {record_details['process_instance_id']}")
                
                return True
                
            except Exception as e:
                # 构建实际执行的SQL语句用于调试
                param_values = [f"'{str(p)}'" if p is not None else 'NULL' for p in params]
                actual_sql = sql.replace('?', '{}').format(*param_values)
                
                print("更新数据库失败:")
                print(f"错误信息: {str(e)}")
                print(f"SQL语句: {actual_sql}")                
                raise
            
    except Exception as e:
        if not str(e).startswith("更新数据库失败"):  # 避免重复错误信息
            print(f"连接数据库失败: {str(e)}")
        raise

# 主逻辑
# 注意，record_id就是process_instance_id
if __name__ == '__main__':     
    token = get_access_token()
    # 1.连接SQL Server数据库，获取[审批状态]字段等于'审批中'的[record_id]字段
    kq_pending_records = get_pending_approval_records("外包考勤")
    bx_pending_records = get_pending_approval_records("外包报销")    
    print(f"当前审批中的考勤记录数: {len(kq_pending_records)}")
    print(f"当前审批中的报销记录数: {len(bx_pending_records)}")

    # 2.遍历pending_records，调用get_approval_record_details，获取审批记录的详细数据;
    for record_id in kq_pending_records:
        record_details = get_approval_record_details(token, record_id)
        #print(f"获取考勤记录ID :{record_id} 的详细信息 ")
        # 将钉钉记录更新到数据库
        try:
            update_kq_record_to_db(record_details)            
        except Exception as e:            
            continue
    for record_id in bx_pending_records:
        record_details = get_approval_record_details(token, record_id)
        #print(f"获取报销记录ID :{record_id} 的详细信息 ")
        # 将钉钉记录更新到数据库
        try:
            update_bx_record_to_db(record_details)            
        except Exception as e:            
            continue
    # 3.连接SQL Server数据库，并查询max([发起时间])，作为起始时间，把当前时间作为结束时间；
    kq_last_create_time = get_max_create_time("外包考勤")
    bx_last_create_time = get_max_create_time("外包报销")
    print(f"最后考勤记录时间: {kq_last_create_time}")
    print(f"最后报销记录时间: {bx_last_create_time}")
    if kq_last_create_time is None:
        kq_last_create_time = datetime.now() - timedelta(days=30)
    if bx_last_create_time is None:
        bx_last_create_time = datetime.now() - timedelta(days=30)
    print("--------------------------------")
    # 4.调用get_kaoqin_ding_data，获取KQ_PROCESS_CODE对应的记录，并插入到数据库
    new_kq_records = get_kaoqin_ding_data(token, KQ_PROCESS_CODE,kq_last_create_time, datetime.now())
    insert_kaoqin_records(new_kq_records)
    print("--------------------------------")
    new_bx_records = get_baoxiao_ding_data(token, BX_PROCESS_CODE, bx_last_create_time, datetime.now())
    insert_baoxiao_records(new_bx_records)
