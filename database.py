import sqlite3
import logging
import time
import functools
import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

def retry_db_operation(max_attempts=3, delay=0.5):
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            for attempt in range(max_attempts):
                try:
                    return func(*args, **kwargs)
                except sqlite3.OperationalError as e:
                    if "database is locked" in str(e) and attempt < max_attempts - 1:
                        logging.warning(f"数据库锁冲突，重试 {attempt + 1}/{max_attempts}")
                        time.sleep(delay)
                    else:
                        raise
            raise Exception("数据库操作失败：超过最大重试次数")
        return wrapper
    return decorator

def init_db():
    with sqlite3.connect('hr_data.db') as conn:
        c = conn.cursor()
        c.execute('''CREATE TABLE IF NOT EXISTS personnel (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            real_name TEXT,
            gender TEXT,
            age INTEGER,
            id_number TEXT,
            phone TEXT,
            province TEXT,
            city TEXT,
            county TEXT,
            nickname TEXT,
            education TEXT,
            political_status TEXT,
            occupation TEXT,
            position TEXT,
            status TEXT,
            join_date TEXT,
            donation_days TEXT,
            address TEXT,
            bio TEXT,
            photo_path TEXT
        )''')
        c.execute('''CREATE TABLE IF NOT EXISTS operation_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            operation_type TEXT,
            operation_target TEXT,
            operation_time TEXT
        )''')
        c.execute('''CREATE TABLE IF NOT EXISTS talent_pool (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            person_id INTEGER,
            add_time TEXT,
            reason TEXT,
            FOREIGN KEY (person_id) REFERENCES personnel(id)
        )''')
        c.execute('''CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            password_hash TEXT,
            password_enabled INTEGER DEFAULT 1  -- 新增字段
        )''')
        conn.commit()
    logging.info("数据库初始化完成")

def migrate_db():
    with sqlite3.connect('hr_data.db') as conn:
        c = conn.cursor()
        c.execute("PRAGMA table_info(talent_pool)")
        columns = [info[1] for info in c.fetchall()]
        if 'reason' not in columns:
            c.execute("ALTER TABLE talent_pool ADD COLUMN reason TEXT")
            logging.info("数据库迁移：talent_pool 表添加 reason 列")
        c.execute("PRAGMA table_info(users)")
        columns = [info[1] for info in c.fetchall()]
        if not columns:
            c.execute("CREATE TABLE users (id INTEGER PRIMARY KEY AUTOINCREMENT, password_hash TEXT, password_enabled INTEGER DEFAULT 1)")
            logging.info("数据库迁移：创建 users 表")
        if 'password_enabled' not in columns:
            c.execute("ALTER TABLE users ADD COLUMN password_enabled INTEGER DEFAULT 1")
            c.execute("UPDATE users SET password_enabled = 1 WHERE password_enabled IS NULL")
            logging.info("数据库迁移：users 表添加 password_enabled 列")
        c.execute("SELECT COUNT(*) FROM users")
        if c.fetchone()[0] == 0:
            from utils import hash_password
            c.execute("INSERT INTO users (id, password_hash, password_enabled) VALUES (1, ?, 1)", (hash_password('123456'),))
            logging.info("初始化默认密码")
        conn.commit()
    logging.info("数据库迁移检查完成")

def load_admin_data():
    admin_data = {}
    with sqlite3.connect('hr_data.db') as conn:
        c = conn.cursor()
        c.execute("SELECT DISTINCT province FROM personnel WHERE province IS NOT NULL AND province != ''")
        provinces = sorted(set(row[0].replace("省", "") for row in c.fetchall()))
        for province in provinces:
            admin_data[province] = []
            c.execute("SELECT DISTINCT city FROM personnel WHERE province=? AND city IS NOT NULL AND city != ''", (province,))
            cities = sorted(set(row[0].replace("市", "") for row in c.fetchall()))
            admin_data[province] = cities
    logging.info("行政数据加载完成")
    return admin_data

@retry_db_operation()
def import_data(file_paths, refresh_callback):
    try:
        with sqlite3.connect('hr_data.db') as conn:
            c = conn.cursor()
            total_count = 0
            total_skipped = 0
            skipped_reasons = []

            c.execute("PRAGMA table_info(personnel)")
            existing_columns = {info[1] for info in c.fetchall() if info[1] not in ['id', 'photo_path']}

            column_mapping = {
                '姓名': 'real_name', '真实姓名': 'real_name', '性别': 'gender', '年龄': 'age',
                '身份证': 'id_number', '身份证号': 'id_number', '电话': 'phone', '手机号': 'phone',
                '省': 'province', '省份': 'province', '市': 'city', '城市': 'city',
                '县': 'county', '县区': 'county', '昵称': 'nickname', '学历': 'education',
                '政治面貌': 'political_status', '职业': 'occupation', '个人职业': 'occupation',
                '职务': 'position', '分会职务': 'position', '状态': 'status', '在职状态': 'status',
                '加入时间': 'join_date', '加入组织时间': 'join_date', '跟捐天数': 'donation_days',
                '地址': 'address', '家庭住址': 'address', '简历': 'bio', '个人简历': 'bio'
            }

            for file_path in file_paths:
                if file_path.endswith('.csv'):
                    df = pd.read_csv(file_path, encoding='utf-8')
                else:
                    df = pd.read_excel(file_path)

                import_columns = list(df.columns)
                mapped_columns = {}
                for col in import_columns:
                    found = False
                    for key, value in column_mapping.items():
                        if key in col:
                            mapped_columns[col] = value
                            found = True
                            break
                    if not found:
                        new_col = col.replace(" ", "_").replace("/", "_")
                        mapped_columns[col] = new_col
                        if new_col not in existing_columns:
                            c.execute(f"ALTER TABLE personnel ADD COLUMN '{new_col}' TEXT")
                            existing_columns.add(new_col)

                count = 0
                skipped = 0
                for _, row in df.iterrows():
                    real_name = str(row.get('真实姓名', row.get('姓名', '')))
                    phone = str(row.get('手机号', row.get('电话', '')))
                    if phone.strip():
                        c.execute("SELECT id FROM personnel WHERE real_name=? AND phone=?", (real_name, phone))
                    else:
                        c.execute("SELECT id FROM personnel WHERE real_name=?", (real_name,))
                    if c.fetchone():
                        skipped += 1
                        skipped_reasons.append(f"记录 '{real_name}' (手机号: {phone or '无'}) 已存在")
                        continue

                    data = {col: '' for col in existing_columns}
                    data['photo_path'] = ''
                    data['status'] = '在职'

                    for import_col, db_col in mapped_columns.items():
                        if import_col in row and pd.notna(row[import_col]):
                            if db_col == 'province':
                                data[db_col] = str(row[import_col]).strip().replace("省", "")
                            elif db_col == 'city':
                                data[db_col] = str(row[import_col]).strip().replace("市", "")
                            elif db_col == 'age':
                                try:
                                    data[db_col] = int(row[import_col])
                                except:
                                    data[db_col] = 0
                            else:
                                data[db_col] = str(row[import_col])

                    columns = list(data.keys())
                    values = list(data.values())
                    placeholders = ", ".join(["?" for _ in columns])
                    c.execute(f"INSERT INTO personnel ({', '.join(columns)}) VALUES ({placeholders})", values)
                    count += 1

                total_count += count
                total_skipped += skipped

            c.execute("INSERT INTO operation_log (operation_type, operation_target, operation_time) VALUES (?, ?, ?)",
                      ("导入数据", f"导入了{total_count}条数据，跳过了{total_skipped}条", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            conn.commit()

        refresh_callback()

        message = f"成功导入 {total_count} 条数据"
        if total_skipped > 0:
            message += f"\n跳过了 {total_skipped} 条数据，原因如下：\n" + "\n".join(skipped_reasons[:5])
        return message, None
    except Exception as e:
        logging.error(f"导入数据失败：{str(e)}")
        return None, f"导入失败：{str(e)}"

@retry_db_operation()
def export_data(export_type, province, city, admin_data):
    with sqlite3.connect('hr_data.db') as conn:
        query = "SELECT * FROM personnel"
        params = []
        default_filename = ""
        if export_type == "division":
            conditions = []
            if province != "全部":
                conditions.append("province=?")
                params.append(province)
                default_filename += province + "分会"
            if city != "全部":
                conditions.append("city=?")
                params.append(city)
                default_filename += city + "分会"
            if conditions:
                query += " WHERE " + " AND ".join(conditions)
            default_filename += "管理层名单"
        else:
            default_filename = "全部数据管理层名单"

        df = pd.read_sql_query(query, conn, params=params)

        if len(df) == 0:
            return None, "未查询到符合条件的数据，请检查选择的分会信息！"

        column_mapping = {
            'real_name': '真实姓名', 'gender': '性别', 'age': '年龄', 'id_number': '身份证号',
            'phone': '手机号', 'province': '省份', 'city': '城市', 'nickname': '昵称',
            'education': '学历', 'political_status': '政治面貌', 'occupation': '个人职业',
            'position': '分会职务', 'status': '在职状态', 'join_date': '加入组织时间',
            'donation_days': '跟捐天数', 'address': '家庭住址', 'bio': '个人简历'
        }
        df = df[list(column_mapping.keys())]
        df.rename(columns=column_mapping, inplace=True)

        return df, default_filename

@retry_db_operation()
def export_talent_pool():
    try:
        with sqlite3.connect('hr_data.db') as conn:
            query = """
                SELECT p.real_name, p.gender, p.age, p.phone, p.province, p.city, 
                       p.position, p.status, p.bio, t.reason, t.add_time
                FROM personnel p
                JOIN talent_pool t ON p.id = t.person_id
            """
            df = pd.read_sql_query(query, conn)
        df.columns = ['真实姓名', '性别', '年龄', '手机号', '省份', '城市', 
                     '分会职务', '在职状态', '个人简历', '加入人才库理由', '加入人才库时间']
        return df, None
    except Exception as e:
        logging.error(f"导出人才库失败：{str(e)}")
        return None, f"导出失败：{str(e)}"

@retry_db_operation()
def save_person(data, mode, person, from_talent=False):
    try:
        with sqlite3.connect('hr_data.db') as conn:
            c = conn.cursor()
            photo_updated = False
            if mode == "edit" and person and person[-1] != data[-1]:
                photo_updated = True
            if mode == "add":
                c.execute("INSERT INTO personnel (real_name, gender, age, id_number, phone, province, city, county, nickname, education, political_status, occupation, position, status, join_date, donation_days, address, bio, photo_path) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", data)
                person_id = c.lastrowid
                operation_type = "新增人员"
                message = "新增人员完成"
            else:
                c.execute("UPDATE personnel SET real_name=?, gender=?, age=?, id_number=?, phone=?, province=?, city=?, county=?, nickname=?, education=?, political_status=?, occupation=?, position=?, status=?, join_date=?, donation_days=?, address=?, bio=?, photo_path=? WHERE id=?", (*data, person[0]))
                person_id = person[0]
                operation_type = "编辑人员"
                message = "编辑信息完成"
                if from_talent:
                    talent_reason = data[-2]  # 假设理由在 data[-2]
                    c.execute("UPDATE talent_pool SET reason=? WHERE person_id=?", (talent_reason, person[0]))
                    operation_type = "编辑人员及人才库理由"
                    message = "编辑信息及人才库理由完成"
            c.execute("INSERT INTO operation_log (operation_type, operation_target, operation_time) VALUES (?, ?, ?)",
                      (operation_type, data[0], datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            conn.commit()
        if photo_updated:
            message += "\n照片已更新"
        return person_id, message, None
    except Exception as e:
        logging.error(f"保存人员失败：{str(e)}")
        return None, None, f"保存失败：数据库操作错误，请重试！"

@retry_db_operation()
def save_and_add_to_talent_pool(data, reason):
    try:
        with sqlite3.connect('hr_data.db') as conn:
            c = conn.cursor()
            c.execute("INSERT INTO personnel (real_name, gender, age, id_number, phone, province, city, county, nickname, education, political_status, occupation, position, status, join_date, donation_days, address, bio, photo_path) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", data)
            person_id = c.lastrowid
            c.execute("INSERT INTO talent_pool (person_id, add_time, reason) VALUES (?, ?, ?)",
                      (person_id, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), reason))
            c.execute("INSERT INTO operation_log (operation_type, operation_target, operation_time) VALUES (?, ?, ?)",
                      ("新增并加入人才库", data[0], datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            conn.commit()
        return person_id, f"新增人员 {data[0]} 并加入人才库完成", None
    except Exception as e:
        logging.error(f"保存并加入人才库失败：{str(e)}")
        return None, None, f"保存失败：数据库操作错误，请重试！"

@retry_db_operation()
def add_to_talent_pool(person_id, reason):
    try:
        with sqlite3.connect('hr_data.db') as conn:
            c = conn.cursor()
            c.execute("SELECT real_name FROM personnel WHERE id=?", (person_id,))
            result = c.fetchone()
            if not result:
                raise ValueError("人员不存在")
            real_name = result[0]
            c.execute("INSERT INTO talent_pool (person_id, add_time, reason) VALUES (?, ?, ?)",
                      (person_id, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), reason))
            c.execute("INSERT INTO operation_log (operation_type, operation_target, operation_time) VALUES (?, ?, ?)",
                      ("加入人才库", real_name, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            conn.commit()
        return f"人员 {real_name} 已加入人才库", None
    except Exception as e:
        logging.error(f"加入人才库失败：{str(e)}")
        return None, f"加入人才库失败：{str(e)}"

@retry_db_operation()
def delete_person(person_id):
    try:
        with sqlite3.connect('hr_data.db') as conn:
            c = conn.cursor()
            c.execute("SELECT real_name FROM personnel WHERE id=?", (person_id,))
            real_name = c.fetchone()[0]
            c.execute("DELETE FROM personnel WHERE id=?", (person_id,))
            c.execute("DELETE FROM talent_pool WHERE person_id=?", (person_id,))
            c.execute("INSERT INTO operation_log (operation_type, operation_target, operation_time) VALUES (?, ?, ?)",
                      ("删除人员", real_name, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            conn.commit()
        return "人员已删除", None
    except Exception as e:
        logging.error(f"删除人员失败：{str(e)}")
        return None, f"删除失败：{str(e)}"