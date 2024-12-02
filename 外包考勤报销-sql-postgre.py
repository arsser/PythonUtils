import pyodbc
import psycopg2
from psycopg2 import sql# SQL Server连接信息

SQLSERVER_CONN_STR = '''
    DRIVER={SQL Server};
    SERVER=10.12.21.65;
    DATABASE=YITU_OA;
    UID=sa;
    PWD=Yitu@123
'''


# PostgreSQL连接信息
PG_DBNAME = "yitu_oa"
PG_USER = "postgres"
PG_PASSWORD = "postgres"
PG_HOST = "localhost"
PG_PORT = "5432"

# 添加数据库迁移相关函数
def get_sqlserver_connection():
    """建立SQL Server连接"""
    return pyodbc.connect(SQLSERVER_CONN_STR)

def get_postgresql_connection():
    """建立PostgreSQL连接"""
    return psycopg2.connect(
        dbname=PG_DBNAME,
        user=PG_USER,
        password=PG_PASSWORD,
        host=PG_HOST,
        port=PG_PORT
    )


def migrate_data_to_postgresql(df, table_name):
    """将数据从DataFrame迁移到PostgreSQL"""
    pg_conn = get_postgresql_connection()
    pg_cursor = pg_conn.cursor()
    
    try:
        # 获取现有记录的审批状态
        pg_cursor.execute(f"""
            SELECT "审批ID", "审批状态" 
            FROM {table_name}
        """)
        existing_records = dict(pg_cursor.fetchall())
        
        # 准备数据插入
        columns = df.columns.tolist()
        for _, row in df.iterrows():
            approval_id = row['审批ID']
            approval_status = row['审批状态']
            
            if approval_id not in existing_records:
                # 新记录，直接插入
                values = [row[col] for col in columns]
                placeholders = ','.join(['%s'] * len(columns))
                insert_query = f"""
                    INSERT INTO {table_name} 
                    ({','.join([f'"{col}"' for col in columns])})
                    VALUES ({placeholders})
                """
                pg_cursor.execute(insert_query, values)
            
            elif existing_records[approval_id] != approval_status:
                # 记录存在但状态不同，更新记录
                update_sets = [f'"{col}" = %s' for col in columns]
                values = [row[col] for col in columns]
                values.append(approval_id)  # WHERE条件的值
                
                update_query = f"""
                    UPDATE {table_name}
                    SET {', '.join(update_sets)}
                    WHERE "审批ID" = %s
                """
                pg_cursor.execute(update_query, values)
        
        pg_conn.commit()
        print(f"成功迁移数据到{table_name}表")
        
    except Exception as e:
        pg_conn.rollback()
        print(f"数据迁移失败: {str(e)}")
    finally:
        pg_cursor.close()
        pg_conn.close()