import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import json
import os
import logging
from modules.utils import create_popup

class EmployeeManager:
    def __init__(self, app):
        self.app = app
        self.employees = []
        self.employee_fields = [
            "họ_tên_uq", "giới_tính_uq", "ngày_sinh_uq", 
            "số_cccd_uq", "ngày_cấp_uq", "nơi_cấp_uq", 
            "địa_chỉ_liên_lạc_uq", "sdt_uq", "email_uq"
        ]
        self.load_employees()
    
    def load_employees(self):
        """Load employees from database"""
        try:
            # Lấy dữ liệu từ SQLite thay vì từ file JSON
            self.employees = self.app.config_manager.db_manager.get_employees()
            
            # Kiểm tra xem có dữ liệu nào không
            if not self.employees:
                # Thử tìm và chuyển đổi từ file JSON sang SQLite
                employees_file = os.path.join("AppData", "employees.json")
                if os.path.exists(employees_file):
                    with open(employees_file, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                        # Chuyển đổi từ khóa không dấu sang có dấu nếu cần
                        self.employees = []
                        for emp in data:
                            # Tạo một bản ghi mới với khóa có dấu
                            new_emp = {}
                            # Map khóa từ không dấu sang có dấu
                            key_mapping = {
                                'ho_ten_uq': 'họ_tên_uq',
                                'gioi_tinh_uq': 'giới_tính_uq',
                                'ngay_sinh_uq': 'ngày_sinh_uq',
                                'so_cccd_uq': 'số_cccd_uq',
                                'ngay_cap_uq': 'ngày_cấp_uq',
                                'noi_cap_uq': 'nơi_cấp_uq',
                                'dia_chi_lien_lac_uq': 'địa_chỉ_liên_lạc_uq',
                                'sdt_uq': 'sdt_uq',
                                'email_uq': 'email_uq'
                            }
                            
                            # Chuyển đổi khóa
                            for old_key, new_key in key_mapping.items():
                                if old_key in emp:
                                    new_emp[new_key] = emp[old_key]
                                elif new_key in emp:  # Đã có khóa mới
                                    new_emp[new_key] = emp[new_key]
                                else:
                                    new_emp[new_key] = ""  # Giá trị mặc định nếu không tìm thấy
                            
                            self.employees.append(new_emp)
                        
                        # Lưu lại vào cơ sở dữ liệu SQLite
                        for employee in self.employees:
                            self.app.config_manager.db_manager.save_employee(employee)
        except Exception as e:
            logging.error(f"Error loading employees: {str(e)}")
            messagebox.showerror("Lỗi", f"Không thể tải danh sách nhân viên: {str(e)}")
    
    def save_employees(self):
        """Save employees to database"""
        # Phương thức này không còn cần thiết vì mọi thao tác đều được lưu trực tiếp vào cơ sở dữ liệu
        pass
    
    def add_employee(self):
        """Add a new employee"""
        popup = create_popup(self.app.root, "Thêm nhân viên", 450, 500)
        
        # Create a frame for employee details
        frame = ttk.Frame(popup, padding=10)
        frame.pack(fill="both", expand=True)
        
        # Create entry fields
        entries = {}
        row = 0
        for field in self.employee_fields:
            display_name = field.replace('_uq', '').replace('_', ' ').title()
            ttk.Label(frame, text=f"{display_name}:").grid(row=row, column=0, padx=5, pady=5, sticky="e")
            entry = ttk.Entry(frame, width=40)
            entry.grid(row=row, column=1, padx=5, pady=5, sticky="w")
            entries[field] = entry
            row += 1
        
        # Add buttons
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=row, column=0, columnspan=2, pady=10)
        
        def confirm_add():
            employee = {field: entries[field].get() for field in self.employee_fields}
            # Check if required fields are filled
            if not employee["họ_tên_uq"]:
                messagebox.showwarning("Cảnh báo", "Vui lòng nhập họ tên nhân viên!")
                return
            
            # Add employee to SQLite
            if self.app.config_manager.db_manager.save_employee(employee):
                # Reload employees from database
                self.employees = self.app.config_manager.db_manager.get_employees()
                messagebox.showinfo("Thành công", "Đã thêm nhân viên!")
                popup.destroy()
                
                # Update employee dropdown if it exists
                self.update_employee_dropdown()
            else:
                messagebox.showerror("Lỗi", "Không thể lưu thông tin nhân viên!")
        
        ttk.Button(button_frame, text="Lưu", command=confirm_add).pack(side="left", padx=10, pady=10, expand=True)
        ttk.Button(button_frame, text="Hủy", command=popup.destroy).pack(side="right", padx=10, pady=10, expand=True)
    
    def edit_employee(self):
        """Edit an existing employee"""
        if not self.employees:
            messagebox.showwarning("Cảnh báo", "Không có nhân viên nào để sửa!")
            return
        
        # Create dropdown to select employee
        select_popup = create_popup(self.app.root, "Sửa nhân viên", 300, 150)
        
        ttk.Label(select_popup, text="Chọn nhân viên:").pack(pady=10)
        employee_var = tk.StringVar()
        employee_dropdown = ttk.Combobox(select_popup, textvariable=employee_var, state="readonly", width=30)
        employee_dropdown["values"] = [emp["họ_tên_uq"] for emp in self.employees]
        employee_dropdown.pack(pady=10)
        
        def open_edit_form():
            selected = employee_var.get()
            if not selected:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn nhân viên!")
                return
            
            # Find selected employee
            employee = None
            for emp in self.employees:
                if emp["họ_tên_uq"] == selected:
                    employee = emp
                    break
            
            if not employee:
                return
            
            select_popup.destroy()
            
            # Open edit form
            edit_popup = create_popup(self.app.root, "Sửa thông tin nhân viên", 450, 500)   
            
            # Create a frame for employee details
            frame = ttk.Frame(edit_popup, padding=10)
            frame.pack(fill="both", expand=True)
            
            # Create entry fields
            entries = {}
            row = 0
            for field in self.employee_fields:
                display_name = field.replace('_uq', '').replace('_', ' ').title()
                ttk.Label(frame, text=f"{display_name}:").grid(row=row, column=0, padx=5, pady=5, sticky="e")
                entry = ttk.Entry(frame, width=40)
                entry.insert(0, employee.get(field, ""))
                entry.grid(row=row, column=1, padx=5, pady=5, sticky="w")
                entries[field] = entry
                row += 1
            
            # Add buttons
            button_frame = ttk.Frame(frame)
            button_frame.grid(row=row, column=0, columnspan=2, pady=10)
            
            def confirm_edit():
                updated = {field: entries[field].get() for field in self.employee_fields}
                # Check if required fields are filled
                if not updated["họ_tên_uq"]:
                    messagebox.showwarning("Cảnh báo", "Vui lòng nhập họ tên nhân viên!")
                    return
                
                # Update employee in SQLite
                if self.app.config_manager.db_manager.save_employee(updated):
                    # Reload employees from database
                    self.employees = self.app.config_manager.db_manager.get_employees()
                    messagebox.showinfo("Thành công", "Đã cập nhật thông tin nhân viên!")
                    edit_popup.destroy()
                    
                    # Update employee dropdown if it exists
                    self.update_employee_dropdown()
                else:
                    messagebox.showerror("Lỗi", "Không thể cập nhật thông tin nhân viên!")
            
            ttk.Button(button_frame, text="Lưu", command=confirm_edit).pack(side="left", padx=10)
            ttk.Button(button_frame, text="Hủy", command=edit_popup.destroy).pack(side="right", padx=10)
        
        button_frame = ttk.Frame(select_popup)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="Chọn", command=open_edit_form).pack(side="left", padx=10)
        ttk.Button(button_frame, text="Hủy", command=select_popup.destroy).pack(side="right", padx=10)
    
    def delete_employee(self):
        """Delete an existing employee"""
        if not self.employees:
            messagebox.showwarning("Cảnh báo", "Không có nhân viên nào để xóa!")
            return
        
        # Create dropdown to select employee
        select_popup = create_popup(self.app.root, "Xóa nhân viên", 300, 150)
        
        ttk.Label(select_popup, text="Chọn nhân viên cần xóa:").pack(pady=10)
        employee_var = tk.StringVar()
        employee_dropdown = ttk.Combobox(select_popup, textvariable=employee_var, state="readonly", width=30)
        employee_dropdown["values"] = [emp["họ_tên_uq"] for emp in self.employees]
        employee_dropdown.pack(pady=10)
        
        def confirm_delete():
            selected = employee_var.get()
            if not selected:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn nhân viên!")
                return
            
            if messagebox.askyesno("Xác nhận", f"Bạn có chắc muốn xóa nhân viên '{selected}' không?"):
                # Delete employee from SQLite
                if self.app.config_manager.db_manager.delete_employee(selected):
                    # Reload employees from database
                    self.employees = self.app.config_manager.db_manager.get_employees()
                    messagebox.showinfo("Thành công", "Đã xóa nhân viên!")
                    select_popup.destroy()
                    
                    # Update employee dropdown if it exists
                    self.update_employee_dropdown()
                else:
                    messagebox.showerror("Lỗi", "Không thể xóa nhân viên!")
        
        button_frame = ttk.Frame(select_popup)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="Xóa", command=confirm_delete).pack(side="left", padx=10)
        ttk.Button(button_frame, text="Hủy", command=select_popup.destroy).pack(side="right", padx=10)
    
    def create_employee_management_ui(self):
        """Create UI for employee management"""
        popup = create_popup(self.app.root, "Quản lý nhân viên ủy quyền", 700, 400)
        
        # Create a frame for employee list
        frame = ttk.Frame(popup, padding=10)
        frame.pack(fill="both", expand=True)
        
        # Create Treeview to display employees
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        employee_tree = ttk.Treeview(tree_frame, columns=self.employee_fields, show="headings", height=15)
        
        # Define column headings
        for field in self.employee_fields:
            display_name = field.replace('_uq', '').replace('_', ' ').title()
            employee_tree.heading(field, text=display_name)
            employee_tree.column(field, width=100)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=employee_tree.yview)
        employee_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        employee_tree.pack(side="left", fill="both", expand=True)
        
        # Add buttons
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill="x", padx=5, pady=10)
        
        # Load employee data into Treeview
        for emp in self.employees:
            employee_tree.insert("", "end", values=[emp.get(field, "") for field in self.employee_fields])
        
        # Thêm chức năng làm mới dữ liệu
        def refresh_data():
            # Xóa dữ liệu cũ
            for item in employee_tree.get_children():
                employee_tree.delete(item)
            
            # Tải lại dữ liệu từ cơ sở dữ liệu
            self.employees = self.app.config_manager.db_manager.get_employees()
            
            # Hiển thị dữ liệu mới
            for emp in self.employees:
                employee_tree.insert("", "end", values=[emp.get(field, "") for field in self.employee_fields])
        
        ttk.Button(button_frame, text="Thêm nhân viên", command=self.add_employee).pack(side="left", padx=10)
        ttk.Button(button_frame, text="Sửa nhân viên", command=self.edit_employee).pack(side="left", padx=10)
        ttk.Button(button_frame, text="Xóa nhân viên", command=self.delete_employee).pack(side="left", padx=10)
        ttk.Button(button_frame, text="Làm mới", command=refresh_data).pack(side="left", padx=10)
        ttk.Button(button_frame, text="Đóng", command=popup.destroy).pack(side="right", padx=10)
    
    def update_employee_dropdown(self):
        """Update employee dropdown in auth info tab if it exists"""
        current_tab = self.app.notebook.tab(self.app.notebook.select(), "text")
        if current_tab == "Thông tin uỷ quyền" and hasattr(self.app, 'employee_dropdown'):
            # Lấy dữ liệu mới từ cơ sở dữ liệu
            self.employees = self.app.config_manager.db_manager.get_employees()
            self.app.employee_dropdown["values"] = ["-- Chọn nhân viên --"] + [emp["họ_tên_uq"] for emp in self.employees]
    
    def add_employee_dropdown_to_tab(self, tab_frame):
        """Add employee dropdown to authority info tab"""
        # Create a frame for the dropdown at the top of the tab
        dropdown_frame = ttk.Frame(tab_frame)
        dropdown_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="ew")
        
        ttk.Label(dropdown_frame, text="Chọn nhân viên ủy quyền:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        
        # Create employee dropdown
        employee_var = tk.StringVar()
        self.app.employee_dropdown = ttk.Combobox(dropdown_frame, textvariable=employee_var, state="readonly", width=30)
        self.app.employee_dropdown["values"] = ["-- Chọn nhân viên --"] + [emp["họ_tên_uq"] for emp in self.employees]
        self.app.employee_dropdown.set("-- Chọn nhân viên --")
        self.app.employee_dropdown.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        # Add button
        ttk.Button(dropdown_frame, text="Quản lý nhân viên", command=self.create_employee_management_ui).grid(row=0, column=2, padx=5, pady=5, sticky="w")
        
        # Adjust the rows of the existing widgets
        # We'll shift all form fields in the tab down one row
        for i, field in enumerate(self.app.field_groups.get("Thông tin uỷ quyền", [])):
            if field in self.app.labels and field in self.app.entries:
                # Get current widgets
                label = self.app.labels[field]
                entry = self.app.entries[field]
                
                # Shift them down by one row
                label.grid(row=i+1, column=0, padx=5, pady=2, sticky="e")
                entry.grid(row=i+1, column=1, padx=5, pady=2, sticky="ew")
        
        # Bind selection event
        def on_employee_selected(event):
            selected = employee_var.get()
            if selected == "-- Chọn nhân viên --":
                return
            
            # Find selected employee
            for emp in self.employees:
                if emp["họ_tên_uq"] == selected:
                    # Fill in the form fields
                    for field in self.employee_fields:
                        if field in self.app.entries:
                            self.app.entries[field].delete(0, tk.END)
                            self.app.entries[field].insert(0, emp.get(field, ""))
                    break
        
        self.app.employee_dropdown.bind("<<ComboboxSelected>>", on_employee_selected)