from tkinter import ttk, messagebox, simpledialog
import os
import sys
import logging
from datetime import datetime
from modules.database import DatabaseManager
from modules.gui import create_popup

class ConfigManager:
    def __init__(self, app):
        self.app = app
        self.db_manager = DatabaseManager()
        self.configs = self.db_manager.get_configs()
        if self.configs:
            self.current_config_name = list(self.configs.keys())[0]
        else:
            self.current_config_name = None
    
    def initialize_default_config(self):
        """Khởi tạo cấu hình mặc định nếu chưa có."""
        default_config = {
            "field_groups": {
                "Thông tin công ty": self.app.default_fields[0:11],
                "Thông tin ĐDPL": self.app.default_fields[11:22],
                "Thông tin thành viên": [],
                "Ngành nghề kinh doanh": [],
                "Thông tin uỷ quyền": self.app.default_fields[22:]  
            },
            "templates": {},
            "entries": []
        }
        
        # Thêm cấu hình mặc định vào configs
        self.configs["Mặc định"] = default_config
        
        # Lưu vào cơ sở dữ liệu
        self.db_manager.save_config("Mặc định", default_config)
        
        # Cập nhật current_config_name
        self.current_config_name = "Mặc định"
        
        logging.info("Đã khởi tạo cấu hình mặc định")
        return default_config
    
    def save_configs(self):
        """Lưu cấu hình hiện tại vào cơ sở dữ liệu."""
        try:
            for config_name, config_data in self.configs.items():
                self.db_manager.save_config(config_name, config_data)
            logging.info("Đã lưu cấu hình")
        except Exception as e:
            logging.error(f"Lỗi khi lưu configs: {str(e)}")
            messagebox.showerror("Lỗi", f"Không thể lưu cấu hình: {str(e)}")
    
    def add_config(self, name):
        """Thêm cấu hình mới."""
        if name in self.configs:
            return False
        
        self.configs[name] = {
            "tabs": {},
            "templates": {}
        }
        self.db_manager.save_config(name, self.configs[name])
        return True
    
    def delete_config(self, name):
        """Xóa cấu hình."""
        if name in self.configs:
            del self.configs[name]
            # Xóa từ SQLite sẽ được thực hiện bởi database manager
            conn = self.db_manager.get_connection()
            cursor = conn.cursor()
            cursor.execute("DELETE FROM configs WHERE name = ?", (name,))
            conn.commit()
            conn.close()
            return True
        return False
    
    def rename_config(self, old_name, new_name):
        """Đổi tên cấu hình."""
        if old_name in self.configs and new_name not in self.configs:
            self.configs[new_name] = self.configs.pop(old_name)
            
            # Cập nhật trong SQLite
            conn = self.db_manager.get_connection()
            cursor = conn.cursor()
            cursor.execute(
                "UPDATE configs SET name = ? WHERE name = ?",
                (new_name, old_name)
            )
            # Cập nhật tên cấu hình cho các mục dữ liệu
            cursor.execute(
                "UPDATE entries SET config_name = ? WHERE config_name = ?",
                (new_name, old_name)
            )
            conn.commit()
            conn.close()
            
            if self.current_config_name == old_name:
                self.current_config_name = new_name
            return True
        return False

    def load_selected_config(self, event):
        """Tải cấu hình được chọn từ dropdown."""
        self.current_config_name = self.app.config_var.get()
        
        # Lấy cấu hình từ cơ sở dữ liệu
        config_data = self.configs.get(self.current_config_name, {})
        
        # Kiểm tra và chuyển đổi cấu trúc từ format cũ sang format mới nếu cần
        if "tabs" in config_data and "field_groups" not in config_data:
            # Chuyển đổi từ cấu trúc cũ (tabs) sang cấu trúc mới (field_groups)
            field_groups = {}
            for tab_name, tab_data in config_data["tabs"].items():
                field_list = list(tab_data.get("fields", {}).keys())
                field_groups[tab_name] = field_list
            
            config_data["field_groups"] = field_groups
            # Lưu cấu trúc đã chuyển đổi
            self.configs[self.current_config_name] = config_data
            self.db_manager.save_config(self.current_config_name, config_data)
        
        # Lấy dữ liệu field_groups từ cấu hình
        self.app.field_groups = config_data.get("field_groups", {})
        
        # Xóa các tab hiện tại và tạo lại
        self.app.clear_tabs()
        self.app.tab_manager.create_tabs()
        
        # Cập nhật dropdown danh sách tab
        self.app.tab_dropdown["values"] = list(self.app.field_groups.keys())
        if self.app.field_groups:
            self.app.tab_dropdown.set(list(self.app.field_groups.keys())[0])
        else:
            self.app.tab_dropdown.set("")
        
        # Xóa dữ liệu trong dropdown và không tự động chọn dữ liệu đầu tiên
        if hasattr(self.app, 'load_data_dropdown'):
            self.app.load_data_dropdown.set("")
            self.app.load_data_dropdown["values"] = [entry["name"] for entry in self.app.saved_entries]
        elif hasattr(self.app, 'search_combobox'):
            self.app.search_combobox.set("")
            self.app.search_combobox["values"] = [entry["name"] for entry in self.app.saved_entries]
            
        # Cập nhật cây templates
        self.app.update_template_tree()
        
        # Cập nhật danh sách trường
        self.app.update_field_dropdown()
        
        # Log thông tin
        logging.info(f"Đã tải cấu hình: {self.current_config_name}")

    def add_new_config(self):
        config_name = simpledialog.askstring("Thêm cấu hình", "Nhập tên cấu hình mới:")
        if config_name and config_name not in self.configs:
            self.configs[config_name] = {
                "field_groups": {
                    "Thông tin công ty": self.app.default_fields[0:11],
                    "Thông tin ĐDPL": self.app.default_fields[11:22],
                    "Thông tin thành viên": [],
                    "Ngành nghề kinh doanh": [],
                    "Thông tin uỷ quyền": self.app.default_fields[22:]  
                },
                "templates": {},
                "entries": []
            }
            # Lưu vào SQLite
            self.db_manager.save_config(config_name, self.configs[config_name])
            
            self.app.config_dropdown["values"] = list(self.configs.keys())
            self.app.config_dropdown.set(config_name)
            self.load_selected_config(None)
            logging.info(f"Thêm cấu hình '{config_name}'")

    def delete_current_config(self):
        """UI handler to delete the current configuration."""
        if not self.current_config_name:
            return
        if messagebox.askyesno("Xác nhận", f"Bạn có chắc muốn xóa cấu hình '{self.current_config_name}' không?"):
            # Get the name before deleting
            name_to_delete = self.current_config_name
            
            # Delete from configs dictionary and database
            self.delete_config(name_to_delete)
            
            # Update UI
            self.app.config_dropdown["values"] = list(self.configs.keys())
            if self.configs:
                self.app.config_dropdown.set(list(self.configs.keys())[0])
                self.load_selected_config(None)
            else:
                self.initialize_default_config()
                self.app.config_dropdown.set("Mặc định")
                self.load_selected_config(None)
            logging.info(f"Xóa cấu hình '{name_to_delete}'")

    def rename_current_config(self):
        """UI handler to rename the current configuration."""
        old_name = self.current_config_name
        if not old_name:
            return
        new_name = simpledialog.askstring("Sửa tên cấu hình", "Nhập tên mới:", initialvalue=old_name)
        if new_name and new_name != old_name and new_name not in self.configs:
            # Rename in database and configs dictionary
            self.rename_config(old_name, new_name)
            
            # Update UI
            self.app.config_dropdown["values"] = list(self.configs.keys())
            self.app.config_dropdown.set(new_name)
            messagebox.showinfo("Thành công", f"Đã đổi tên thành '{new_name}'!")
            logging.info(f"Đổi tên cấu hình từ '{old_name}' thành '{new_name}'")

class BackupManager:
    def __init__(self, app):
        self.app = app

    def auto_backup(self):
        """Tự động sao lưu dữ liệu."""
        try:
            # Tạo bản sao lưu của cơ sở dữ liệu SQLite
            backup_path = self.app.config_manager.db_manager.create_backup()
            logging.info(f"Đã tạo bản sao lưu tự động: {backup_path}")
            
            # Xóa các bản sao lưu cũ (giữ lại 10 bản gần nhất)
            self.cleanup_old_backups()
            
            # Thiết lập hẹn giờ cho bản sao lưu tiếp theo
            self.app.root.after(600000, self.auto_backup)
            
            return True
        except Exception as e:
            logging.error(f"Lỗi khi tạo bản sao lưu tự động: {str(e)}")
            return False
    
    def cleanup_old_backups(self, max_backups=10):
        """Xóa các bản sao lưu cũ, giữ lại tối đa max_backups bản gần nhất."""
        backup_dir = "backup"
        if not os.path.exists(backup_dir):
            return
            
        # Liệt kê các file backup theo thời gian tạo
        backups = [os.path.join(backup_dir, f) for f in os.listdir(backup_dir) 
                if f.startswith("backup_") and f.endswith(".db")]
        backups.sort(key=lambda x: os.path.getmtime(x), reverse=True)
        
        # Xóa các file cũ
        for old_backup in backups[max_backups:]:
            try:
                os.remove(old_backup)
                logging.info(f"Đã xóa bản sao lưu cũ: {old_backup}")
            except Exception as e:
                logging.error(f"Không thể xóa bản sao lưu: {str(e)}")

    def restore_from_backup(self):
        """Khôi phục dữ liệu từ bản sao lưu SQLite."""
        if not os.path.exists(self.app.backup_dir):
            messagebox.showwarning("Cảnh báo", "Không tìm thấy thư mục sao lưu!")
            return
            
        backups = [f for f in os.listdir(self.app.backup_dir) if f.startswith("backup_") and f.endswith(".db")]
        if not backups:
            messagebox.showwarning("Cảnh báo", "Không có bản sao lưu nào!")
            return
            
        backups.sort(reverse=True)  # Sắp xếp theo thứ tự giảm dần (mới nhất đầu tiên)
        
        # Tạo popup để hiển thị danh sách bản sao lưu
        popup = create_popup(self.app.root, "Chọn bản sao lưu để khôi phục", 390, 390)
        
        ttk.Label(popup, text="Chọn bản sao lưu:").pack(pady=5)
        
        backup_frame = ttk.Frame(popup)
        backup_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        backup_tree = ttk.Treeview(backup_frame, columns=("lastmod"), show="tree headings", height=10, selectmode="browse")
        backup_tree.heading("#0", text="Tên file")
        backup_tree.heading("lastmod", text="Ngày tạo")
        backup_tree.column("#0", width=200)
        backup_tree.column("lastmod", width=150)
        backup_tree.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(backup_frame, orient="vertical", command=backup_tree.yview)
        scrollbar.pack(side="right", fill="y")
        backup_tree.configure(yscrollcommand=scrollbar.set)
        
        for backup_file in backups:
            file_path = os.path.join(self.app.backup_dir, backup_file)
            try:
                # Định dạng: backup_20250404_061140.db -> 04/04/2025 06:11:40
                date_str = backup_file.split("_")[1]
                time_str = backup_file.split("_")[2].split(".")[0]
                formatted_date = f"{date_str[6:8]}/{date_str[4:6]}/{date_str[0:4]} {time_str[0:2]}:{time_str[2:4]}:{time_str[4:6]}"
            except:
                formatted_date = datetime.fromtimestamp(os.path.getctime(file_path)).strftime("%Y-%m-%d %H:%M:%S")
            
            backup_tree.insert("", "end", text=backup_file, values=(formatted_date,))
        
        def confirm_restore():
            selected = backup_tree.selection()
            if not selected:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn một bản sao lưu!")
                return
            
            backup_file = backup_tree.item(selected[0])["text"]
            backup_path = os.path.join(self.app.backup_dir, backup_file)
            
            if messagebox.askyesno("Xác nhận", f"Bạn có muốn khôi phục từ {backup_file}? Dữ liệu hiện tại sẽ bị ghi đè."):
                try:
                    # Đường dẫn đến database hiện tại
                    db_path = self.app.config_manager.db_manager.db_path
                    
                    # Đóng tất cả các kết nối database
                    try:
                        # Xóa tham chiếu đến connection pool
                        self.app.config_manager.db_manager.conn = None
                        
                        # Thử gọi garbage collector để đóng kết nối còn tồn tại
                        import gc
                        gc.collect()
                    except Exception as conn_error:
                        logging.error(f"Lỗi khi đóng kết nối: {str(conn_error)}")
                    
                    # Tạo bản sao lưu trước khi khôi phục
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    pre_restore_backup = os.path.join(self.app.backup_dir, f"pre_restore_{timestamp}.db")
                    
                    if os.path.exists(db_path):
                        try:
                            import shutil
                            shutil.copy2(db_path, pre_restore_backup)
                            logging.info(f"Đã tạo bản sao lưu trước khi khôi phục: {pre_restore_backup}")
                        except Exception as backup_error:
                            logging.error(f"Lỗi khi tạo bản sao lưu trước khôi phục: {str(backup_error)}")
                    
                    # Thay vì xóa file và tạo lại, sử dụng phương pháp an toàn hơn
                    try:
                        import shutil
                        temp_db_path = db_path + ".new"
                        
                        # Sao chép backup vào file tạm
                        shutil.copy2(backup_path, temp_db_path)
                        
                        # Đóng tất cả kết nối một lần nữa để đảm bảo
                        self.app.config_manager.db_manager.conn = None
                        gc.collect()
                        
                        # Thử xóa file cũ
                        if os.path.exists(db_path):
                            os.remove(db_path)
                        
                        # Đổi tên file mới
                        os.rename(temp_db_path, db_path)
                        logging.info(f"Đã khôi phục database từ: {backup_path}")
                    except Exception as replace_error:
                        logging.error(f"Lỗi khi thay thế file database: {str(replace_error)}")
                        raise replace_error
                    
                    # Thông báo thành công và khởi động lại
                    messagebox.showinfo("Thành công", "Đã khôi phục dữ liệu từ bản sao lưu! Ứng dụng sẽ khởi động lại.")
                    popup.destroy()
                    
                    # Đóng ứng dụng an toàn
                    self.app.root.destroy()
                    os.execl(sys.executable, sys.executable, *sys.argv)
                    
                except Exception as e:
                    error_msg = f"Không thể khôi phục từ bản sao lưu: {str(e)}"
                    logging.error(error_msg)
                    messagebox.showerror("Lỗi", error_msg)
        
        button_frame = ttk.Frame(popup)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="Khôi phục", command=confirm_restore, style="primary-outline").pack(side="left", padx=5)
        ttk.Button(button_frame, text="Đóng", command=popup.destroy, style="danger-outline").pack(side="left", padx=5)