import tkinter as tk
from tkinter import ttk, messagebox
import logging
from modules.utils import create_popup_with_notebook


class MemberManager:
    def __init__(self, app):
        self.app = app
#hàm tạp tab thông tin thành viên
    def create_member_tab(self):
        """Tạo tab Thông tin thành viên với Treeview và scrollbar dọc."""
        tab_name = "Thông tin thành viên"
        tab = ttk.Frame(self.app.notebook)
        self.app.notebook.add(tab, text=tab_name)
        
        # Tạo Canvas và Scrollbar để hỗ trợ cuộn dọc
        canvas = tk.Canvas(tab, bg="#ffffff", highlightthickness=0)
        scrollbar_y = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, padding=5)
        scrollable_frame.bind("<Configure>", lambda e, c=canvas: c.configure(scrollregion=c.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar_y.set)
        
        # Đặt vị trí Canvas và Scrollbar dọc
        canvas.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        scrollbar_y.pack(side="right", fill="y")
        
        # Tạo frame chứa Treeview
        tree_frame = ttk.Frame(scrollable_frame)
        tree_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Chỉ hiển thị các cột quan trọng
        self.app.displayed_columns = ["ho_ten", "so_cccd", "dia_chi_thuong_tru", "von_gop", "ty_le_gop", "la_chu_tich"]
        self.app.member_tree = ttk.Treeview(tree_frame, columns=self.app.displayed_columns, show="headings", height=15)  # Tăng height để hiển thị nhiều hàng hơn
        
        # Định nghĩa tiêu đề và chiều rộng cột
        column_widths = {
            "ho_ten": 200,      # Họ tên rộng hơn để dễ đọc
            "so_cccd": 100,
            "dia_chi_thuong_tru": 400,
            "von_gop": 120,
            "ty_le_gop": 100,
            "la_chu_tich": 80,  # Cột Chủ tịch nhỏ gọn
        }
        for col in self.app.displayed_columns:
            self.app.member_tree.heading(col, text=col.replace('_', ' ').title())
            width = column_widths.get(col, 100)
            self.app.member_tree.column(col, width=width, anchor="center")
        
        # Đặt Treeview vào frame
        self.app.member_tree.pack(side="top", fill="both", expand=True)
        
        # Tạo menu ngữ cảnh (context menu)
        context_menu = tk.Menu(self.app.root, tearoff=0)
        context_menu.add_command(label="Thêm thành viên", command=self.add_member)
        context_menu.add_command(label="Sửa thành viên", command=self.edit_member)
        context_menu.add_command(label="Xóa thành viên", command=self.delete_member)
        context_menu.add_command(label="Xem chi tiết", command=self.view_member_details)

        def show_context_menu(event):
            selected = self.app.member_tree.identify_row(event.y)
            if selected:
                self.app.member_tree.selection_set(selected)
            context_menu.post(event.x_root, event.y_root)

        # Gắn sự kiện nhấp chuột phải để hiển thị menu ngữ cảnh
        self.app.member_tree.bind("<Button-3>", show_context_menu)
        
        # Gắn sự kiện nhấp đúp để xem chi tiết
        self.app.member_tree.bind("<Double-1>", self.view_member_details)

        # Gắn các sự kiện kéo thả
        self.app.member_tree.bind("<Button-1>", self.start_drag_member)
        self.app.member_tree.bind("<B1-Motion>", self.drag_member)
        self.app.member_tree.bind("<ButtonRelease-1>", self.drop_member)
        
        # Tạo frame cho các nút
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="Thêm thành viên", command=self.add_member).pack(side="left", padx=10, expand=True)
        ttk.Button(button_frame, text="Sửa thành viên", command=self.edit_member).pack(side="left", padx=10, expand=True)
        ttk.Button(button_frame, text="Xóa thành viên", command=self.delete_member).pack(side="left", padx=10, expand=True)
        ttk.Button(button_frame, text="Xem chi tiết", command=self.view_member_details).pack(side="left", padx=10, expand=True)
        
        # Kích hoạt cuộn chuột dọc
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))
        
        # Gọi load_member_data để hiển thị dữ liệu ngay khi tab được tạo
        self.load_member_data()

#hàm tải dữ liệu thành viên từ cơ sở dữ liệu vào Treeview
    def load_member_data(self):
        """Tải dữ liệu thành viên vào Treeview."""
        # Xóa dữ liệu hiện tại
        self.app.member_tree.delete(*self.app.member_tree.get_children())
        selected_name = self.app.load_data_var.get()
        
        # Trực tiếp lấy dữ liệu mới nhất từ cơ sở dữ liệu thay vì dùng cached data
        entries = self.app.config_manager.db_manager.get_entries(
            self.app.config_manager.current_config_name
        )
        
        # Cập nhật saved_entries để đồng bộ với database
        self.app.saved_entries = entries
        
        # Tìm và hiển thị thành viên
        members = []
        for entry in entries:
            if entry["name"] == selected_name:
                members = entry["data"].get("thanh_vien", [])
                break
        
        # Hiển thị thành viên
        for member in members:
            values = [member.get(col, "") for col in self.app.displayed_columns]
            # Chuyển đổi giá trị la_chu_tich từ boolean sang "X" hoặc ""
            is_chairman = False
            if isinstance(member.get("la_chu_tich"), bool):
                is_chairman = member.get("la_chu_tich")
            elif member.get("la_chu_tich") == "X" or member.get("la_chu_tich") == "x":
                is_chairman = True
                
            values[self.app.displayed_columns.index("la_chu_tich")] = "X" if is_chairman else ""
            self.app.member_tree.insert("", "end", values=values)

#Hàm thao tác 
    def add_member(self):
        tabs = {
            "Thông tin cá nhân": ["ho_ten", "gioi_tinh", "ngay_sinh", "dan_toc", "quoc_tich"],
            "Thông tin giấy tờ": ["loai_giay_to", "so_cccd", "ngay_cap", "noi_cap", "ngay_het_han"],
            "Thông tin địa chỉ": ["dia_chi_thuong_tru", "dia_chi_lien_lac"],
            "Thông tin vốn góp": ["von_gop", "ty_le_gop", "ngay_gop_von"]
        }
        popup, entries = create_popup_with_notebook(self.app.root, "Thêm thành viên", 450, 550, tabs)

        # Thêm checkbox "Chức danh chủ tịch" vào dưới cùng của popup
        chairman_var = tk.BooleanVar(value=False)
        chairman_check = ttk.Checkbutton(popup, text="Chức danh chủ tịch", variable=chairman_var)
        chairman_check.pack(pady=10)

        # Nút xác nhận
        def confirm_add():
            member = {col: entries[col].get() for col in self.app.member_columns}
            member["la_chu_tich"] = chairman_var.get()
            selected_name = self.app.load_data_var.get()
            
            # Tìm entry hiện tại
            for entry in self.app.saved_entries:
                if entry["name"] == selected_name:
                    members = entry["data"].setdefault("thanh_vien", [])
                    # Nếu thành viên mới là chủ tịch, bỏ chọn chủ tịch cũ
                    if member["la_chu_tich"]:
                        for m in members:
                            m["la_chu_tich"] = False
                    # Thêm thành viên mới vào danh sách
                    members.append(member)
                    
                    # Lưu vào SQLite
                    self.app.config_manager.db_manager.save_entry(
                        self.app.config_manager.current_config_name,
                        selected_name,
                        entry["data"]
                    )
                    # QUAN TRỌNG: Cập nhật saved_entries từ cơ sở dữ liệu
                    self.app.saved_entries = self.app.config_manager.db_manager.get_entries(
                        self.app.config_manager.current_config_name
                    )
                    
                    # Tải lại dữ liệu
                    self.load_member_data()
                    popup.destroy()
                    break

        ttk.Button(popup, text="Thêm", command=confirm_add).pack(pady=10)

    def delete_member(self):
        """Xóa thành viên đã chọn."""
        selected_items = self.app.member_tree.selection()
        if not selected_items:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn ít nhất một thành viên để xóa!")
            return
        
        if not messagebox.askyesno("Xác nhận", f"Bạn có chắc muốn xóa {len(selected_items)} thành viên đã chọn không?"):
            return
            
        selected_name = self.app.load_data_var.get()
        
        try:
            for entry in self.app.saved_entries:
                if entry["name"] == selected_name:
                    # Lấy danh sách thành viên hiện tại
                    members = entry["data"].get("thanh_vien", [])
                    # Tạo danh sách mới không bao gồm các thành viên đã chọn
                    new_members = []
                    for i, member in enumerate(members):
                        if i not in [self.app.member_tree.index(item) for item in selected_items]:
                            new_members.append(member)
                    
                    # Tạo bản sao dữ liệu trước khi sửa đổi
                    updated_data = entry["data"].copy()
                    # Cập nhật danh sách thành viên
                    updated_data["thanh_vien"] = new_members
                    
                    # Lưu vào SQLite với dữ liệu mới
                    self.app.config_manager.db_manager.save_entry(
                        self.app.config_manager.current_config_name,
                        selected_name,
                        updated_data
                    )
                    
                    # Cập nhật saved_entries từ cơ sở dữ liệu để đồng bộ
                    self.app.saved_entries = self.app.config_manager.db_manager.get_entries(
                        self.app.config_manager.current_config_name
                    )
                    
                    # Tải lại dữ liệu thành viên
                    self.load_member_data()
                    messagebox.showinfo("Thành công", f"Đã xóa {len(selected_items)} thành viên!")
                    break
        except Exception as e:
            logging.error(f"Lỗi khi xóa thành viên: {str(e)}")
            messagebox.showerror("Lỗi", f"Không thể xóa thành viên: {str(e)}")

    def edit_member(self):
        """Sửa thông tin thành viên đã chọn."""
        selected_item = self.app.member_tree.selection()
        if not selected_item:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn thành viên để sửa!")
            return
        idx = self.app.member_tree.index(selected_item)
        selected_name = self.app.load_data_var.get()
        member = None
        for entry in self.app.saved_entries:
            if entry["name"] == selected_name:
                member = entry["data"]["thanh_vien"][idx]
                break

        tabs = {
            "Thông tin cá nhân": ["ho_ten", "gioi_tinh", "ngay_sinh", "dan_toc", "quoc_tich"],
            "Thông tin giấy tờ": ["loai_giay_to", "so_cccd", "ngay_cap", "noi_cap", "ngay_het_han"],
            "Thông tin địa chỉ": ["dia_chi_thuong_tru", "dia_chi_lien_lac"],
            "Thông tin vốn góp": ["von_gop", "ty_le_gop", "ngay_gop_von"]
        }
        popup, entries = create_popup_with_notebook(self.app.root, "Sửa thành viên", 450, 550, tabs)

        # Điền dữ liệu hiện tại vào các trường
        for field, entry in entries.items():
            entry.insert(0, member.get(field, ""))

        # Thêm checkbox "Chức danh chủ tịch" - FIX: Ép kiểu về boolean
        is_chairman = False
        if isinstance(member.get("la_chu_tich"), bool):
            is_chairman = member.get("la_chu_tich")
        elif member.get("la_chu_tich") == "X" or member.get("la_chu_tich") == "x":
            is_chairman = True
        
        chairman_var = tk.BooleanVar(value=is_chairman)
        chairman_check = ttk.Checkbutton(popup, text="Chức danh chủ tịch", variable=chairman_var)
        chairman_check.pack(pady=5)

        # Nút xác nhận
        def confirm_edit():
            updated_member = {col: entries[col].get() for col in self.app.member_columns}
            updated_member["la_chu_tich"] = chairman_var.get()
            
            for entry in self.app.saved_entries:
                if entry["name"] == selected_name:
                    members = entry["data"].get("thanh_vien", [])
                    
                    # Nếu thành viên được cập nhật là chủ tịch, bỏ chọn chủ tịch cũ
                    if updated_member["la_chu_tich"]:
                        for m in members:
                            if m != members[idx]:  # Không cập nhật chính nó
                                m["la_chu_tich"] = False
                    
                    # Cập nhật thành viên
                    members[idx] = updated_member
                    
                    # Lưu vào SQLite
                    self.app.config_manager.db_manager.save_entry(
                        self.app.config_manager.current_config_name,
                        selected_name,
                        entry["data"]
                    )
                    # Cập nhật saved_entries từ cơ sở dữ liệu
                    self.app.saved_entries = self.app.config_manager.db_manager.get_entries(
                        self.app.config_manager.current_config_name
                    )
                    
                    # Tải lại dữ liệu
                    self.load_member_data()
                    popup.destroy()
                    break

        ttk.Button(popup, text="Lưu", command=confirm_edit).pack(pady=10)

    def view_member_details(self, event=None):
        """Hiển thị chi tiết thông tin thành viên."""
        selected_items = self.app.member_tree.selection()
        if not selected_items:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn thành viên để xem chi tiết!")
            return
        
        # Kiểm tra nếu chọn nhiều thành viên
        if len(selected_items) > 1:
            messagebox.showwarning("Cảnh báo", "Vui lòng chỉ chọn một thành viên để xem chi tiết!")
            return
            
        selected_item = selected_items[0]  # Lấy item đầu tiên từ selection
        idx = self.app.member_tree.index(selected_item)
        selected_name = self.app.load_data_var.get()
        member = None
        for entry in self.app.saved_entries:
            if entry["name"] == selected_name:
                member = entry["data"]["thanh_vien"][idx]
                break

        tabs = {
            "Thông tin cá nhân": ["ho_ten", "gioi_tinh", "ngay_sinh", "dan_toc", "quoc_tich"],
            "Thông tin giấy tờ": ["loai_giay_to", "so_cccd", "ngay_cap", "noi_cap", "ngay_het_han"],
            "Thông tin địa chỉ": ["dia_chi_thuong_tru", "dia_chi_lien_lac"],
            "Thông tin vốn góp": ["von_gop", "ty_le_gop", "ngay_gop_von"]
        }
        popup, entries = create_popup_with_notebook(self.app.root, "Chi tiết thành viên", 450, 550, tabs)

        # Điền dữ liệu hiện tại vào các trường và đặt chúng ở trạng thái chỉ đọc
        for field, entry in entries.items():
            entry.insert(0, member.get(field, ""))
            entry.config(state="readonly")

        # Hiển thị trạng thái chủ tịch
        ttk.Label(popup, text="Chức danh: " + ("Chủ tịch" if member.get("la_chu_tich", False) else "Thành viên")).pack(pady=5)

        # Nút đóng
        ttk.Button(popup, text="Chỉnh sửa", command=lambda: [popup.destroy(), self.edit_member()]).pack(side="left", padx=10, expand=True)
        ttk.Button(popup, text="Đóng", command=popup.destroy).pack(side="right", pady=10, expand=True)

    def start_drag_member(self, event):
        """Bắt đầu kéo một thành viên."""
        item = self.app.member_tree.identify_row(event.y)
        if item:
            self.app.drag_item = item

    def drag_member(self, event):
        """Di chuyển thành viên khi kéo."""
        if self.app.drag_item:
            self.app.member_tree.selection_set(self.app.drag_item)

    def drop_member(self, event):
        """Thả thành viên vào vị trí mới và cập nhật danh sách."""
        if self.app.drag_item:
            target = self.app.member_tree.identify_row(event.y)
            if target and target != self.app.drag_item:
                # Lấy vị trí hiện tại và mục tiêu
                current_index = self.app.member_tree.index(self.app.drag_item)
                target_index = self.app.member_tree.index(target)
                
                # Cập nhật thứ tự trong bộ nhớ
                selected_name = self.app.load_data_var.get()
                for entry in self.app.saved_entries:
                    if entry["name"] == selected_name:
                        members = entry["data"].get("thanh_vien", [])
                        # Di chuyển thành viên
                        member = members.pop(current_index)
                        members.insert(target_index, member)
                        
                        # Lưu vào SQLite
                        self.app.config_manager.db_manager.save_entry(
                            self.app.config_manager.current_config_name,
                            selected_name,
                            entry["data"]
                        )
                        # Cập nhật saved_entries từ cơ sở dữ liệu
                        self.app.saved_entries = self.app.config_manager.db_manager.get_entries(
                            self.app.config_manager.current_config_name
                        )
                        # Tải lại dữ liệu
                        self.load_member_data()
                        break
            
            self.app.drag_item = None
