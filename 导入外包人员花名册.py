import pandas as pd
import pyodbc
from datetime import datetime

def get_db_connection():
    """创建数据库连接"""
    conn_str = (
        'DRIVER={ODBC Driver 17 for SQL Server};'
        'SERVER=powerbi01;'
        'DATABASE=DigiYitu;'
        'UID=partner;'
        'PWD=galaxy'
    )
    return pyodbc.connect(conn_str)

def safe_convert(value, target_type, default=None):
    """安全地转换数据类型"""
    if pd.isna(value) or value == 'NaT' or value == 'nan':
        return default
    try:
        if target_type == str:
            return str(value) if value not in ('', None) else default
        elif target_type == int:
            return int(float(value)) if value not in ('', None) else default
        elif target_type == float:
            return float(value) if value not in ('', None) else default
        elif target_type == datetime.date:
            if isinstance(value, datetime):
                return value.date()
            return pd.to_datetime(value).date() if value not in ('', None) else default
        return value
    except:
        return default
    
def format_datetime(dt):
    """格式化日期时间为SQL Server可接受的格式"""
    if isinstance(dt, datetime):
        return dt.strftime('%Y-%m-%d %H:%M:%S')
    elif isinstance(dt, datetime.date):
        return dt.strftime('%Y-%m-%d')
    return None

def import_hr_data(data_date,excel_file, sheet_name=None):
    """
    将外包人员信息从Excel导入到SQL Server
    
    Args:
        excel_file (str): Excel文件路径
        sheet_name (str, optional): 工作表名称
    """
    # 定义数据库表的列名（带方括号）
    db_columns = [
        '[员工编号]', '[员工姓名]', '[供应商]', '[员工类别]', '[入职日期]', '[离职日期]', 
        '[上级主管]', '[部门负责人]', '[成本中心编号]', '[员工薪资]', '[基础收费]',
        '[透视分类（最小一级组织）]', '[项目编码 2024]', '[项目名称 2024]',
        '[24实际申请预算]', '[OA预算申请状态]', '[备注]'
    ]
    
    # Excel列名（不带方括号）
    excel_columns = [col.strip('[]') for col in db_columns]

    try:
        # 读取Excel文件
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        
        # 检查必要的列是否存在
        missing_columns = [col for col in excel_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(f"Excel中缺少以下列: {', '.join(missing_columns)}")
        
        # 只保留需要的列
        df = df[excel_columns]
        
        # 获取当前时间
        current_time = datetime.now()
        #data_year_month = datetime(data_date.year, data_date.month, 1)
        
        # 添加时间字段
        df['数据所属的年月'] = data_date
        df['数据导入时间'] = current_time
        
        # 处理日期格式
        df['入职日期'] = pd.to_datetime(df['入职日期']).dt.date
        df['离职日期'] = pd.to_datetime(df['离职日期']).dt.date
        
        # SQL语句（带方括号）
        delete_sql = """
            DELETE FROM dbo.[外包人员表] WHERE YEAR([数据所属的年月]) = ? AND MONTH([数据所属的年月]) = ?
        """
        
        # 动态生成INSERT语句（带方括号）
        columns = db_columns + ['[数据所属的年月]', '[数据导入时间]']
        placeholders = ','.join(['?' for _ in columns])
        insert_sql = f"""INSERT INTO dbo.[外包人员表] ({','.join(columns)}) VALUES ({placeholders})"""        
        with get_db_connection() as conn:
            cursor = conn.cursor()
            
            # 先删除同年月的数据
            cursor.execute(delete_sql, (data_date.year, data_date.month))
            deleted_count = cursor.rowcount
            print(f"已删除 {data_date.year}年{data_date.month}月 的历史数据：{deleted_count} 条记录")
            # 插入新数据
            inserted_count = 0
            error_count = 0
            
            for _, row in df.iterrows():
                try:
                    # 转换数据类型
                    params = [
                            safe_convert(row['员工编号'], str, ''),
                            safe_convert(row['员工姓名'], str, ''),
                            safe_convert(row['供应商'], str, ''),
                            safe_convert(row['员工类别'], str, ''),
                            safe_convert(row['入职日期'], datetime.date),
                            safe_convert(row['离职日期'], datetime.date),
                            safe_convert(row['上级主管'], str, ''),
                            safe_convert(row['部门负责人'], str, ''),
                            safe_convert(row['成本中心编号'], str, ''),
                            safe_convert(row['员工薪资'], int),
                            safe_convert(row['基础收费'], float),
                            safe_convert(row['透视分类（最小一级组织）'], str, ''),
                            safe_convert(row['项目编码 2024'], str),
                            safe_convert(row['项目名称 2024'], str),
                            safe_convert(row['24实际申请预算'], int),
                            safe_convert(row['OA预算申请状态'], str, ''),
                            safe_convert(row['备注'], str, ''),
                            format_datetime(data_date),  # 格式化日期时间
                            format_datetime(current_time)         # 格式化日期时间
                        ]
                    
                    cursor.execute(insert_sql, params)
                    inserted_count += 1
                    
                    # if inserted_count % 100 == 0:  # 每100条打印一次进度
                    #     print(f"已处理 {inserted_count} 条记录")
                    
                except Exception as e:
                    error_count += 1
                    print(f"插入数据失败: {str(e)}")
                    print(f"员工姓名: {row.get('员工姓名', 'Unknown')}")
                    # 构建实际执行的SQL语句用于调试
                    param_values = [f"'{str(p)}'" if p is not None else 'NULL' for p in params]
                    actual_sql = insert_sql.replace('?', '{}').format(*param_values)
                    print(f"失败的SQL语句: {actual_sql}")
                    # print("参数值:")
                    # for name, value in zip(excel_columns + ['数据所属的年月', '数据导入时间'], params):
                    #     print(f"  {name}: {value} (类型: {type(value)})")
                    continue
            
            conn.commit()
            print(f"数据导入完成: 成功 {inserted_count} 条，失败 {error_count} 条")
            
    except Exception as e:
        print(f"程序执行失败: {str(e)}")
        raise

if __name__ == '__main__':
    excel_file = r'\\10.12.21.65\share\外包费用\新流程\【HR】外包花名册.xlsx'
    # 获取当前年月作为sheet名
    current_date = datetime.strptime('2024-10-31','%Y-%m-%d')
    current_sheet = current_date.strftime("%Y-%m")    
    try:
        import_hr_data(current_date, excel_file, current_sheet)
    except Exception as e:
        print(f"导入失败: {str(e)}")