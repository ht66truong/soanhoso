import sqlite3
import json
import os
import logging
from datetime import datetime

class DatabaseManager:
    def __init__(self, db_path='AppData/database.db'):
        self.db_path = db_path
        self.initialize_database()
        
    def get_connection(self):
        """Tạo và trả về kết nối đến cơ sở dữ liệu SQLite."""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row  # Cho phép truy cập vào cột bằng tên
        # THÊM DÒNG NÀY để bật foreign key constraints
        conn.execute("PRAGMA foreign_keys = ON;")
        return conn
        
    def initialize_database(self):
        """Khởi tạo cơ sở dữ liệu với các bảng cần thiết."""
        # Đảm bảo thư mục tồn tại
        os.makedirs(os.path.dirname(self.db_path), exist_ok=True)
        
        conn = self.get_connection()
        cursor = conn.cursor()
        
        # Bảng cấu hình
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS configs (
            id INTEGER PRIMARY KEY,
            name TEXT UNIQUE NOT NULL,
            data TEXT NOT NULL
        )
        ''')
        
        # Bảng dữ liệu đã lưu
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS entries (
            id INTEGER PRIMARY KEY,
            name TEXT NOT NULL,
            config_name TEXT NOT NULL,
            data TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(name, config_name)
        )
        ''')
        
        # Bảng thành viên
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS members (
            id INTEGER PRIMARY KEY,
            entry_id INTEGER,
            data TEXT NOT NULL,
            position INTEGER NOT NULL,
            FOREIGN KEY (entry_id) REFERENCES entries (id) ON DELETE CASCADE
        )
        ''')
        
        # Bảng ngành nghề
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS industries (
            id INTEGER PRIMARY KEY,
            entry_id INTEGER,
            data TEXT NOT NULL,
            position INTEGER NOT NULL,
            is_main INTEGER DEFAULT 0,
            FOREIGN KEY (entry_id) REFERENCES entries (id) ON DELETE CASCADE
        )
        ''')

        # THÊM MỚI: Bảng mã ngành tham khảo
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS industry_codes (
            ma_nganh TEXT PRIMARY KEY,
            ten_nganh TEXT NOT NULL
        )
        ''')
        
        # THÊM MỚI: Bảng nhân viên ủy quyền
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY,
            data TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        ''')
        
        conn.commit()
        conn.close()
    
    def save_config(self, config_name, config_data):
        """Lưu cấu hình vào cơ sở dữ liệu."""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute(
            "INSERT OR REPLACE INTO configs (name, data) VALUES (?, ?)",
            (config_name, json.dumps(config_data, ensure_ascii=False))
        )
        
        conn.commit()
        conn.close()
    
    def get_configs(self):
        """Lấy tất cả cấu hình từ cơ sở dữ liệu."""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute("SELECT name, data FROM configs")
        results = cursor.fetchall()
        
        configs = {}
        for row in results:
            configs[row['name']] = json.loads(row['data'])
        
        conn.close()
        return configs
    
    def get_entries(self, config_name):
        """Lấy tất cả các mục dữ liệu đã lưu cho một cấu hình."""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute("SELECT id, name, data FROM entries WHERE config_name = ?", (config_name,))
        entry_results = cursor.fetchall()
        
        entries = []
        for entry_row in entry_results:
            entry_id = entry_row['id']
            entry_data = json.loads(entry_row['data'])
            
            # Lấy thành viên
            cursor.execute(
                "SELECT data, position FROM members WHERE entry_id = ? ORDER BY position",
                (entry_id,)
            )
            members = [json.loads(row['data']) for row in cursor.fetchall()]
            
            # Lấy ngành nghề
            cursor.execute(
                "SELECT data, position, is_main FROM industries WHERE entry_id = ? ORDER BY position",
                (entry_id,)
            )
            industries = []
            for ind_row in cursor.fetchall():
                industry = json.loads(ind_row['data'])
                industry['la_nganh_chinh'] = bool(ind_row['is_main'])
                industries.append(industry)
            
            # Thêm thành viên và ngành nghề vào dữ liệu
            entry_data['thanh_vien'] = members
            entry_data['nganh_nghe'] = industries
            
            entries.append({
                "name": entry_row['name'],
                "data": entry_data
            })
        
        conn.close()
        return entries
    
    def save_entry(self, config_name, entry_name, entry_data):
        """Lưu hoặc cập nhật một mục dữ liệu."""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        try:
            # Tách dữ liệu thành viên và ngành nghề
            # Thay đổi từ pop sang get để tránh thay đổi dữ liệu gốc
            members = entry_data.get("thanh_vien", [])
            industries = entry_data.get("nganh_nghe", [])
            
            # Tạo bản sao dữ liệu để tránh ảnh hưởng đến dữ liệu gốc
            entry_data_copy = entry_data.copy()
            if "thanh_vien" in entry_data_copy: del entry_data_copy["thanh_vien"]
            if "nganh_nghe" in entry_data_copy: del entry_data_copy["nganh_nghe"]
            
            # Lưu dữ liệu chính
            cursor.execute(
                """
                INSERT OR REPLACE INTO entries (name, config_name, data, updated_at)
                VALUES (?, ?, ?, datetime('now'))
                """,
                (entry_name, config_name, json.dumps(entry_data_copy, ensure_ascii=False))
            )
            
            # Lấy ID 
            entry_id = cursor.lastrowid
            
            # Xóa data cũ
            cursor.execute("DELETE FROM members WHERE entry_id = ?", (entry_id,))
            cursor.execute("DELETE FROM industries WHERE entry_id = ?", (entry_id,))
            
            # Thêm thành viên mới với xử lý đặc biệt cho boolean - kết hợp cả hai vòng lặp thành một
            for position, member in enumerate(members):
                # Đảm bảo la_chu_tich luôn là boolean
                if "la_chu_tich" in member:
                    if isinstance(member["la_chu_tich"], str):
                        member["la_chu_tich"] = member["la_chu_tich"].lower() in ["true", "x", "1", "yes"]
                
                cursor.execute(
                    "INSERT INTO members (entry_id, data, position) VALUES (?, ?, ?)",
                    (entry_id, json.dumps(member, ensure_ascii=False), position)
                )
            
            # Thêm ngành nghề mới
            for position, industry in enumerate(industries):
                is_main = 1 if industry.get("la_nganh_chinh", False) else 0
                cursor.execute(
                    "INSERT INTO industries (entry_id, data, position, is_main) VALUES (?, ?, ?, ?)",
                    (entry_id, json.dumps(industry, ensure_ascii=False), position, is_main)
                )
            
            conn.commit()
        except Exception as e:
            conn.rollback()
            logging.error(f"Lỗi khi lưu dữ liệu: {str(e)}")
            raise e
        finally:
            conn.close()
        
    def delete_entry(self, config_name, entry_name):
        """Xóa một mục dữ liệu."""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute(
            "DELETE FROM entries WHERE name = ? AND config_name = ?",
            (entry_name, config_name)
        )
        
        conn.commit()
        conn.close()
    
    def rename_entry(self, config_name, old_name, new_name):
        """Đổi tên một mục dữ liệu."""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute(
            "UPDATE entries SET name = ? WHERE name = ? AND config_name = ?",
            (new_name, old_name, config_name)
        )
        
        conn.commit()
        conn.close()

    def create_backup(self):
        """Tạo bản sao lưu của cơ sở dữ liệu."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_dir = "backup"
        os.makedirs(backup_dir, exist_ok=True)
        backup_path = os.path.join(backup_dir, f"backup_{timestamp}.db")
        
        conn = sqlite3.connect(self.db_path)
        backup_conn = sqlite3.connect(backup_path)
        
        conn.backup(backup_conn)
        
        conn.close()
        backup_conn.close()
        
        return backup_path
    
    def get_industry_codes(self):
        """Lấy danh sách mã ngành từ cơ sở dữ liệu."""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute("SELECT ma_nganh, ten_nganh FROM industry_codes ORDER BY ma_nganh")
        results = cursor.fetchall()
        
        industry_codes = []
        for row in results:
            industry_codes.append({
                "ma_nganh": row['ma_nganh'],
                "ten_nganh": row['ten_nganh']
            })
        
        conn.close()
        return industry_codes
    
    def get_employees(self):
        """Lấy danh sách nhân viên ủy quyền từ cơ sở dữ liệu."""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute("SELECT id, data FROM employees ORDER BY id")
        results = cursor.fetchall()
        
        employees = []
        for row in results:
            employee = json.loads(row['data'])
            employees.append(employee)
        
        conn.close()
        return employees
    
    def save_employee(self, employee_data):
        """Thêm hoặc cập nhật thông tin một nhân viên ủy quyền."""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        try:
            # Tìm kiếm nhân viên theo họ tên (giả sử họ tên là duy nhất)
            ho_ten = employee_data.get("họ_tên_uq", "")
            cursor.execute(
                "SELECT id, data FROM employees WHERE json_extract(data, '$.họ_tên_uq') = ?",
                (ho_ten,)
            )
            result = cursor.fetchone()
            
            if result:
                # Cập nhật nhân viên đã tồn tại
                employee_id = result['id']
                cursor.execute(
                    "UPDATE employees SET data = ?, updated_at = datetime('now') WHERE id = ?",
                    (json.dumps(employee_data, ensure_ascii=False), employee_id)
                )
            else:
                # Thêm nhân viên mới
                cursor.execute(
                    "INSERT INTO employees (data, updated_at) VALUES (?, datetime('now'))",
                    (json.dumps(employee_data, ensure_ascii=False),)
                )
            
            conn.commit()
            return True
        except Exception as e:
            conn.rollback()
            logging.error(f"Lỗi khi lưu nhân viên: {str(e)}")
            return False
        finally:
            conn.close()
    
    def delete_employee(self, ho_ten):
        """Xóa một nhân viên ủy quyền theo tên."""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        try:
            cursor.execute(
                "DELETE FROM employees WHERE json_extract(data, '$.họ_tên_uq') = ?",
                (ho_ten,)
            )
            conn.commit()
            return True
        except Exception as e:
            conn.rollback()
            logging.error(f"Lỗi khi xóa nhân viên: {str(e)}")
            return False
        finally:
            conn.close()