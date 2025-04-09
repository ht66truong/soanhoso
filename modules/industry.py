import tkinter as tk
from tkinter import ttk, messagebox
from modules.utils import create_popup

class IndustryManager:
    def __init__(self, app):
        self.app = app
# Khởi tạo các tab ngành nghề kinh doanh
    def create_industry_tab(self):
        """Tạo tab Ngành nghề kinh doanh với Treeview và các nút quản lý."""
        tab_name = "Ngành nghề kinh doanh"
        tab = ttk.Frame(self.app.notebook)
        self.app.notebook.add(tab, text=tab_name)
        
        # Tạo Canvas và Scrollbar để hỗ trợ cuộn
        canvas = tk.Canvas(tab, bg="#ffffff", highlightthickness=0)
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, padding=5)
        scrollable_frame.bind("<Configure>", lambda e, c=canvas: c.configure(scrollregion=c.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="center")  # Căn giữa scrollable_frame
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Căn giữa canvas và scrollbar trong tab
        canvas.pack(side="left", fill="both", expand=True, padx=20, pady=20)  # Thêm padding để tạo khoảng cách
        scrollbar.pack(side="right", fill="y")
        
        # Tăng kích thước Treeview và căn giữa
        self.app.industry_tree = ttk.Treeview(scrollable_frame, columns=("ten_nganh", "ma_nganh", "la_nganh_chinh"), show="headings", height=15)
        self.app.industry_tree.heading("ten_nganh", text="Tên ngành")
        self.app.industry_tree.heading("ma_nganh", text="Mã ngành")
        self.app.industry_tree.heading("la_nganh_chinh", text="Ngành chính")
        
        # Tăng chiều rộng cột để hiển thị nội dung rõ ràng hơn
        self.app.industry_tree.column("ten_nganh", width=600)  # Tăng từ 400 lên 600
        self.app.industry_tree.column("ma_nganh", width=150)   # Tăng từ 100 lên 150
        self.app.industry_tree.column("la_nganh_chinh", width=150)  # Tăng từ 100 lên 150
        
        # Căn giữa Treeview trong scrollable_frame
        self.app.industry_tree.pack(fill="x", expand=True, padx=10, pady=10)
        
        # Tạo menu ngữ cảnh (context menu)
        context_menu = tk.Menu(self.app.root, tearoff=0)
        context_menu.add_command(label="Thêm ngành", command=self.add_industry)
        context_menu.add_command(label="Sửa ngành", command=self.edit_industry)
        context_menu.add_command(label="Xóa ngành", command=self.delete_industry)
        context_menu.add_command(label="Ngành chính", command=self.set_main_industry)  # Thêm tùy chọn "Ngành chính"
        context_menu.add_command(label="Xem chi tiết", command=self.view_industry_details)

        def show_context_menu(event):
            selected = self.app.industry_tree.identify_row(event.y)
            if selected:
                self.app.industry_tree.selection_set(selected)
            context_menu.post(event.x_root, event.y_root)

        # Gắn sự kiện nhấp chuột phải để hiển thị menu ngữ cảnh
        self.app.industry_tree.bind("<Button-3>", show_context_menu)
        
        # Gắn sự kiện nhấp đúp để xem chi tiết
        self.app.industry_tree.bind("<Double-1>", self.view_industry_details)
        
        # Tạo frame cho các nút và căn giữa
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.pack(pady=10)
        
        # Tạo các nút và đặt chúng trong button_frame (xóa nút "Đặt ngành chính")
        ttk.Button(button_frame, text="Thêm ngành", command=self.add_industry, style="primary-outline").pack(side="left", padx=10, expand=True)
        ttk.Button(button_frame, text="Xóa ngành", command=self.delete_industry, style="danger-outline").pack(side="left", padx=10, expand=True)
        ttk.Button(button_frame, text="Sửa ngành", command=self.edit_industry, style="primary-outline").pack(side="left", padx=10, expand=True)
        ttk.Button(button_frame, text="Xem chi tiết", command=self.view_industry_details, style="secondary-outline").pack(side="left", padx=10, expand=True)

        # Kích hoạt cuộn chuột
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))
        
        # Load dữ liệu ngành nghề
        self.load_industry_data()

    def create_additional_industry_tab(self):
        """Tạo tab Ngành bổ sung với Treeview và các nút quản lý."""
        tab_name = "Ngành bổ sung"
        tab = ttk.Frame(self.app.notebook)
        self.app.notebook.add(tab, text=tab_name)

        # Tạo Canvas và Scrollbar để hỗ trợ cuộn
        canvas = tk.Canvas(tab, bg="#ffffff", highlightthickness=0)
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, padding=5)
        scrollable_frame.bind("<Configure>", lambda e, c=canvas: c.configure(scrollregion=c.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True, padx=20, pady=20)
        scrollbar.pack(side="right", fill="y")

        # Tạo Treeview
        self.app.additional_industry_tree = ttk.Treeview(scrollable_frame, columns=("ten_nganh", "ma_nganh", "la_nganh_chinh"), show="headings", height=15)
        self.app.additional_industry_tree.heading("ten_nganh", text="Tên ngành")
        self.app.additional_industry_tree.heading("ma_nganh", text="Mã ngành")
        self.app.additional_industry_tree.heading("la_nganh_chinh", text="Ngành chính")
        self.app.additional_industry_tree.column("ten_nganh", width=600)
        self.app.additional_industry_tree.column("ma_nganh", width=150)
        self.app.additional_industry_tree.column("la_nganh_chinh", width=150)
        self.app.additional_industry_tree.pack(fill="x", expand=True, padx=10, pady=10)
        
        # Tạo các nút quản lý
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.pack(pady=10)

        # Tạo menu ngữ cảnh (context menu)
        context_menu = tk.Menu(self.app.root, tearoff=0)
        context_menu.add_command(label="Thêm ngành", command=self.add_industry)
        context_menu.add_command(label="Sửa ngành", command=self.edit_industry)
        context_menu.add_command(label="Xóa ngành", command=self.delete_industry)
        context_menu.add_command(label="Ngành chính", command=self.set_main_industry)  # Thêm tùy chọn "Ngành chính"
        context_menu.add_command(label="Xem chi tiết", command=self.view_industry_details)

        def show_context_menu(event):
            selected = self.app.additional_industry_tree.identify_row(event.y)
            if selected:
                self.app.additional_industry_tree.selection_set(selected)
            context_menu.post(event.x_root, event.y_root)

        # Gắn sự kiện nhấp chuột phải để hiển thị menu ngữ cảnh
        self.app.additional_industry_tree.bind("<Button-3>", show_context_menu)
        
        # Gắn sự kiện nhấp đúp để xem chi tiết
        self.app.additional_industry_tree.bind("<Double-1>", self.view_industry_details)

        ttk.Button(button_frame, text="Thêm ngành", command=self.add_industry, style="primary-outline").pack(side="left", padx=10, expand=True)
        ttk.Button(button_frame, text="Xóa ngành", command=self.delete_industry, style="danger-outline").pack(side="left", padx=10, expand=True)
        ttk.Button(button_frame, text="Sửa ngành", command=self.edit_industry, style="primary-outline").pack(side="left", padx=10, expand=True)
        ttk.Button(button_frame, text="Xem chi tiết", command=self.view_industry_details, style="secondary-outline").pack(side="left", padx=10, expand=True)

        # Kích hoạt cuộn chuột
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        # Load dữ liệu ngành bổ sung
        self.load_additional_industry_data()

    def create_removed_industry_tab(self):
        """Tạo tab Ngành giảm với Treeview và các nút quản lý."""
        tab_name = "Ngành giảm"
        tab = ttk.Frame(self.app.notebook)
        self.app.notebook.add(tab, text=tab_name)

        # Tạo Canvas và Scrollbar để hỗ trợ cuộn
        canvas = tk.Canvas(tab, bg="#ffffff", highlightthickness=0)
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, padding=5)
        scrollable_frame.bind("<Configure>", lambda e, c=canvas: c.configure(scrollregion=c.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True, padx=20, pady=20)
        scrollbar.pack(side="right", fill="y")

        # Tạo Treeview
        self.app.removed_industry_tree = ttk.Treeview(scrollable_frame, columns=("ten_nganh", "ma_nganh", "la_nganh_chinh"), show="headings", height=15)
        self.app.removed_industry_tree.heading("ten_nganh", text="Tên ngành")
        self.app.removed_industry_tree.heading("ma_nganh", text="Mã ngành")
        self.app.removed_industry_tree.heading("la_nganh_chinh", text="Ngành chính")
        self.app.removed_industry_tree.column("ten_nganh", width=600)
        self.app.removed_industry_tree.column("ma_nganh", width=150)
        self.app.removed_industry_tree.column("la_nganh_chinh", width=150)
        self.app.removed_industry_tree.pack(fill="x", expand=True, padx=10, pady=10)

        # Tạo các nút quản lý
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.pack(pady=10)

        # Tạo menu ngữ cảnh (context menu)
        context_menu = tk.Menu(self.app.root, tearoff=0)
        context_menu.add_command(label="Thêm ngành", command=self.add_industry)
        context_menu.add_command(label="Sửa ngành", command=self.edit_industry)
        context_menu.add_command(label="Xóa ngành", command=self.delete_industry)
        context_menu.add_command(label="Ngành chính", command=self.set_main_industry)  # Thêm tùy chọn "Ngành chính"
        context_menu.add_command(label="Xem chi tiết", command=self.view_industry_details)

        def show_context_menu(event):
            selected = self.app.removed_industry_tree.identify_row(event.y)
            if selected:
                self.app.removed_industry_tree.selection_set(selected)
            context_menu.post(event.x_root, event.y_root)

        # Gắn sự kiện nhấp chuột phải để hiển thị menu ngữ cảnh
        self.app.removed_industry_tree.bind("<Button-3>", show_context_menu)
        
        # Gắn sự kiện nhấp đúp để xem chi tiết
        self.app.removed_industry_tree.bind("<Double-1>", self.view_industry_details)

        ttk.Button(button_frame, text="Thêm ngành giảm", command=self.add_industry, style="primary-outline").pack(side="left", padx=10, expand=True)
        #ttk.Button(button_frame, text="Xóa ngành", command=self.delete_industry, style="primary-outline").pack(side="left", padx=10, expand=True)
        #ttk.Button(button_frame, text="Sửa ngành", command=self.edit_industry, style="primary-outline").pack(side="left", padx=10, expand=True)
        ttk.Button(button_frame, text="Xem chi tiết", command=self.view_industry_details, style="secondary-outline").pack(side="left", padx=10, expand=True)

        # Kích hoạt cuộn chuột
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        # Load dữ liệu ngành giảm
        self.load_removed_industry_data()

    def create_adjusted_industry_tab(self):
        """Tạo tab Ngành điều chỉnh với Treeview và các nút quản lý."""
        tab_name = "Ngành điều chỉnh"
        tab = ttk.Frame(self.app.notebook)
        self.app.notebook.add(tab, text=tab_name)

        # Tạo Canvas và Scrollbar để hỗ trợ cuộn
        canvas = tk.Canvas(tab, bg="#ffffff", highlightthickness=0)
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, padding=5)
        scrollable_frame.bind("<Configure>", lambda e, c=canvas: c.configure(scrollregion=c.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True, padx=20, pady=20)
        scrollbar.pack(side="right", fill="y")

        # Tạo Treeview
        self.app.adjusted_industry_tree = ttk.Treeview(scrollable_frame, columns=("ten_nganh", "ma_nganh", "la_nganh_chinh"), show="headings", height=15)
        self.app.adjusted_industry_tree.heading("ten_nganh", text="Tên ngành")
        self.app.adjusted_industry_tree.heading("ma_nganh", text="Mã ngành")
        self.app.adjusted_industry_tree.heading("la_nganh_chinh", text="Ngành chính")
        self.app.adjusted_industry_tree.column("ten_nganh", width=600)
        self.app.adjusted_industry_tree.column("ma_nganh", width=150)
        self.app.adjusted_industry_tree.column("la_nganh_chinh", width=150)
        self.app.adjusted_industry_tree.pack(fill="x", expand=True, padx=10, pady=10)

        # Tạo các nút quản lý
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.pack(pady=10)

        # Tạo menu ngữ cảnh (context menu)
        context_menu = tk.Menu(self.app.root, tearoff=0)
        context_menu.add_command(label="Thêm ngành", command=self.add_industry)
        context_menu.add_command(label="Sửa ngành", command=self.edit_industry)
        context_menu.add_command(label="Xóa ngành", command=self.delete_industry)
        context_menu.add_command(label="Ngành chính", command=self.set_main_industry)
        context_menu.add_command(label="Xem chi tiết", command=self.view_industry_details)

        def show_context_menu(event):
            selected = self.app.adjusted_industry_tree.identify_row(event.y)
            if selected:
                self.app.adjusted_industry_tree.selection_set(selected)
            context_menu.post(event.x_root, event.y_root)

        # Gắn sự kiện nhấp chuột phải để hiển thị menu ngữ cảnh
        self.app.adjusted_industry_tree.bind("<Button-3>", show_context_menu)
        
        # Gắn sự kiện nhấp đúp để xem chi tiết
        self.app.adjusted_industry_tree.bind("<Double-1>", self.view_industry_details)

        ttk.Button(button_frame, text="Thêm ngành", command=self.add_industry, style="primary-outline").pack(side="left", padx=10, expand=True)
        #ttk.Button(button_frame, text="Xóa ngành", command=self.delete_industry, style="primary-outline").pack(side="left", padx=10, expand=True)
        ttk.Button(button_frame, text="Điều chỉnh ngành", command=self.edit_industry, style="primary-outline").pack(side="left", padx=10, expand=True)
        #ttk.Button(button_frame, text="Xem chi tiết", command=self.view_industry_details, style="primary-outline").pack(side="left", padx=10, expand=True)

        # Kích hoạt cuộn chuột
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        # Load dữ liệu ngành điều chỉnh
        self.load_adjusted_industry_data()

 # Load dữ liệu ngành nghề kinh doanh       
    def load_industry_data(self):
        """Tải dữ liệu ngành nghề kinh doanh vào Treeview."""
        self.app.industry_tree.delete(*self.app.industry_tree.get_children())
        selected_name = self.app.load_data_var.get()
        main_industries = []
        
        for entry in self.app.saved_entries:
            if entry["name"] == selected_name:
                main_industries = entry["data"].get("nganh_nghe", [])
                break

        for industry in main_industries:
            is_main = "Có" if industry.get("la_nganh_chinh", False) else ""
            self.app.industry_tree.insert("", "end", values=(
                industry.get("ten_nganh", ""),
                industry.get("ma_nganh", ""),
                is_main
            ))

    def load_additional_industry_data(self):
        """Tải dữ liệu ngành bổ sung vào Treeview."""
        self.app.additional_industry_tree.delete(*self.app.additional_industry_tree.get_children())
        selected_name = self.app.load_data_var.get()
        additional_industries = []
        for entry in self.app.saved_entries:
            if entry["name"] == selected_name:
                additional_industries = entry["data"].get("nganh_bo_sung", [])
                break
        for industry in additional_industries:
            self.app.additional_industry_tree.insert("", "end", values=(industry["ten_nganh"], industry["ma_nganh"], "X" if industry["la_nganh_chinh"] else ""))

    def load_removed_industry_data(self):
        """Tải dữ liệu ngành giảm vào Treeview."""
        self.app.removed_industry_tree.delete(*self.app.removed_industry_tree.get_children())
        selected_name = self.app.load_data_var.get()
        removed_industries = []
        for entry in self.app.saved_entries:
            if entry["name"] == selected_name:
                removed_industries = entry["data"].get("nganh_giam", [])
                break
        for industry in removed_industries:
            self.app.removed_industry_tree.insert("", "end", values=(industry["ten_nganh"], industry["ma_nganh"], "X" if industry["la_nganh_chinh"] else ""))

    def load_adjusted_industry_data(self):
        """Tải dữ liệu ngành điều chỉnh vào Treeview."""
        self.app.adjusted_industry_tree.delete(*self.app.adjusted_industry_tree.get_children())
        selected_name = self.app.load_data_var.get()
        adjusted_industries = []
        for entry in self.app.saved_entries:
            if entry["name"] == selected_name:
                adjusted_industries = entry["data"].get("nganh_dieu_chinh", [])
                break
        for industry in adjusted_industries:
            self.app.adjusted_industry_tree.insert("", "end", values=(industry["ten_nganh"], industry["ma_nganh"], "X" if industry["la_nganh_chinh"] else ""))
    
    def load_industry_data_for_current_tab(self, tree, industries):
        """Tải lại dữ liệu cho Treeview và danh sách ngành của tab hiện tại."""
        tree.delete(*tree.get_children())
        for industry in industries:
            tree.insert("", "end", values=(industry["ten_nganh"], industry["ma_nganh"], "X" if industry.get("la_nganh_chinh", False) else ""))
            # Xóa dòng này vì không cần thiết và gây lỗi khi lưu:
            # self.app.config_manager.save_configs()

#Các hàm helper chung     
    def get_current_tab_tree_and_data(self):
        """Xác định Treeview và danh sách ngành dựa trên tab hiện tại."""
        current_tab = self.app.notebook.tab(self.app.notebook.select(), "text")
        selected_name = self.app.load_data_var.get()
        tree = None
        industries = []

        for entry in self.app.saved_entries:
            if entry["name"] == selected_name:
                if current_tab == "Ngành nghề kinh doanh":
                    tree = self.app.industry_tree
                    industries = entry["data"].get("nganh_nghe", [])
                elif current_tab == "Ngành bổ sung" and hasattr(self.app, 'additional_industry_tree'):
                    tree = self.app.additional_industry_tree
                    industries = entry["data"].get("nganh_bo_sung", [])
                elif current_tab == "Ngành giảm" and hasattr(self.app, 'removed_industry_tree'):
                    tree = self.app.removed_industry_tree
                    industries = entry["data"].get("nganh_giam", [])
                elif current_tab == "Ngành điều chỉnh" and hasattr(self.app, 'adjusted_industry_tree'):
                    tree = self.app.adjusted_industry_tree
                    industries = entry["data"].get("nganh_dieu_chinh", [])
                break
        
        return tree, industries
      
    def sync_main_industry_tab(self, action, current_tab, updated_industry=None, original_industry=None):
        """
        Đồng bộ dữ liệu của tab Ngành nghề kinh doanh dựa trên các thay đổi từ các tab khác.
        
        :param action: Loại hành động ("add", "edit", "delete").
        :param current_tab: Tab hiện tại ("Ngành bổ sung", "Ngành giảm", "Ngành điều chỉnh").
        :param updated_industry: Ngành mới được thêm hoặc sửa (nếu có).
        :param original_industry: Ngành cũ trước khi sửa (chỉ áp dụng cho "edit").
        """
        selected_name = self.app.load_data_var.get()

        for entry in self.app.saved_entries:
            if entry["name"] == selected_name:
                main_industries = entry["data"].setdefault("nganh_nghe", [])

                if current_tab == "Ngành bổ sung":
                    if action == "add" and updated_industry:
                        # Kiểm tra nếu ngành đã tồn tại trong "Ngành nghề kinh doanh"
                        if not any(industry["ma_nganh"] == updated_industry["ma_nganh"] for industry in main_industries):
                            main_industries.append(updated_industry)
                    
                    elif action == "edit" and updated_industry and original_industry:
                        # Tìm và cập nhật ngành tương ứng trong "Ngành nghề kinh doanh"
                        for industry in main_industries:
                            if industry["ma_nganh"] == original_industry["ma_nganh"]:
                                industry.update({
                                    "ma_nganh": updated_industry["ma_nganh"],
                                    "ten_nganh": updated_industry["ten_nganh"],
                                    "la_nganh_chinh": updated_industry["la_nganh_chinh"]
                                })
                                break

                    elif action == "delete" and updated_industry:
                        # Xóa ngành tương ứng khỏi "Ngành nghề kinh doanh"
                        main_industries = [
                            industry for industry in main_industries
                            if industry["ma_nganh"] != updated_industry["ma_nganh"]
                        ]

                elif current_tab == "Ngành giảm":
                    # Khi thêm ngành vào "Ngành giảm", xóa ngành tương ứng khỏi "Ngành nghề kinh doanh"
                    if action == "add" and updated_industry:
                        main_industries = [
                            industry for industry in main_industries
                            if industry["ma_nganh"] != updated_industry["ma_nganh"]
                        ]

                elif current_tab == "Ngành điều chỉnh":
                    # Khi sửa ngành trong "Ngành điều chỉnh", cập nhật thông tin ngành trong "Ngành nghề kinh doanh"
                    if action == "edit" and updated_industry and original_industry:
                        for industry in main_industries:
                            if industry["ma_nganh"] == original_industry["ma_nganh"]:
                                industry.update(updated_industry)
                                break

                # Cập nhật lại dữ liệu trong "Ngành nghề kinh doanh"
                entry["data"]["nganh_nghe"] = main_industries
                break

        # Lưu cấu hình và tải lại dữ liệu cho tab "Ngành nghề kinh doanh"
        self.app.config_manager.db_manager.save_entry(
            self.app.config_manager.current_config_name,
            selected_name,
            entry["data"]
        )
        
        # Cập nhật lại saved_entries từ database
        self.app.saved_entries = self.app.config_manager.db_manager.get_entries(
            self.app.config_manager.current_config_name
        )
        self.load_industry_data()

#Các thao tác CRUD (Create, Read, Update, Delete)        
    def add_industry(self):
        """Thêm một ngành nghề mới với logic tùy chỉnh cho từng tab."""
        current_tab = self.app.notebook.tab(self.app.notebook.select(), "text")
        selected_name = self.app.load_data_var.get()

        if current_tab in ["Ngành giảm", "Ngành điều chỉnh"]:
            # Lấy danh sách ngành từ tab "Ngành nghề kinh doanh"
            main_industries = []
            for entry in self.app.saved_entries:
                if entry["name"] == selected_name:
                    main_industries = entry["data"].get("nganh_nghe", [])
                    break

            if not main_industries:
                messagebox.showwarning("Cảnh báo", "Không có ngành nghề nào trong 'Ngành nghề kinh doanh' để thêm!")
                return

            # Tạo popup
            popup = create_popup(self.app.root, f"Thêm {current_tab.lower()}", 600, 250)
            ttk.Label(popup, text="Chọn ngành từ danh sách ngành nghề kinh doanh:").pack(pady=5)
            industry_var = tk.StringVar()
            industry_combo = ttk.Combobox(popup, textvariable=industry_var, width=90, state="readonly")
            industry_combo.pack(pady=5)

            # Tạo danh sách các ngành để hiển thị trong combobox
            combo_values = []
            for industry in main_industries:
                display_text = f"{industry['ma_nganh']} - {industry['ten_nganh']}"
                combo_values.append(display_text)
            industry_combo['values'] = combo_values

            def confirm_add():
                selected_industry = industry_var.get()
                if not selected_industry:
                    messagebox.showwarning("Cảnh báo", "Vui lòng chọn ngành để thêm!")
                    return

                ma_nganh, ten_nganh = selected_industry.split(" - ", 1)

                # Tìm ngành được chọn trong danh sách ngành chính
                selected_industry_data = None
                for industry in main_industries:
                    if industry["ma_nganh"] == ma_nganh:
                        selected_industry_data = industry.copy()
                        break

                if not selected_industry_data:
                    messagebox.showwarning("Cảnh báo", "Không tìm thấy thông tin ngành!")
                    return

                # Kiểm tra xem ngành đã tồn tại trong danh sách đích chưa
                for entry in self.app.saved_entries:
                    if entry["name"] == selected_name:
                        target_industries = entry["data"].setdefault(
                            "nganh_giam" if current_tab == "Ngành giảm" else "nganh_dieu_chinh", 
                            []
                        )
                        
                        # Kiểm tra trùng lặp
                        if any(industry["ma_nganh"] == ma_nganh for industry in target_industries):
                            messagebox.showwarning("Cảnh báo", f"Ngành này đã tồn tại trong {current_tab}!")
                            return

                        # Thêm ngành vào danh sách đích
                        target_industries.append(selected_industry_data)
                        self.app.config_manager.save_configs()

                        # Tải lại dữ liệu và cập nhật giao diện
                        if current_tab == "Ngành giảm":
                            self.load_removed_industry_data()
                            self.sync_main_industry_tab(
                                action="add",
                                current_tab=current_tab,
                                updated_industry=selected_industry_data
                            )
                        else:  # Ngành điều chỉnh
                            self.load_adjusted_industry_data()
                            self.sync_main_industry_tab(
                                action="add",
                                current_tab=current_tab,
                                updated_industry=selected_industry_data
                            )
                        popup.destroy()
                        break

            ttk.Button(popup, text="Thêm", command=confirm_add, style="primary-outline").pack(side="bottom", padx=10, pady=10)


        elif current_tab == "Ngành bổ sung":
            # Logic thêm ngành cho tab "Ngành bổ sung"
            self.add_industry_default_logic(current_tab)

            # Đồng bộ ngành bổ sung với "Ngành nghề kinh doanh"
            for entry in self.app.saved_entries:
                if entry["name"] == selected_name:
                    industries = entry["data"].get("nganh_bo_sung", [])
                    if industries:
                        updated_industry = industries[-1]  # Lấy ngành vừa thêm
                        self.sync_main_industry_tab(action="add", current_tab=current_tab, updated_industry=updated_industry)
                        self.app.industry_manager.load_industry_data()  # Tải lại dữ liệu cho tab "Ngành nghề kinh doanh"
                    break

        else:
            # Logic mặc định cho các tab khác (ví dụ: "Ngành nghề kinh doanh")
            self.add_industry_default_logic(current_tab)

    def add_industry_default_logic(self, current_tab):
        """Logic mặc định để thêm ngành cho các tab khác."""
        
        # Lấy từ SQLite:
        all_industries = self.app.config_manager.db_manager.get_industry_codes()
        
        # Tạo popup
        popup = create_popup(self.app.root, "Thêm ngành", 600, 250)
        popup.title("Thêm ngành")

        ttk.Label(popup, text="Chọn mã ngành và tên ngành:").pack(pady=5)
        industry_var = tk.StringVar()
        industry_combo = ttk.Combobox(popup, textvariable=industry_var, width=90)
        industry_combo.pack(pady=5)

        full_list = [f"{code['ma_nganh']} - {code['ten_nganh']}" for code in all_industries]

        def update_list(*args):
            search_term = industry_var.get().lower()
            filtered_list = [item for item in full_list if search_term in item.lower()] if search_term else full_list
            industry_combo['values'] = filtered_list

        industry_combo.bind('<KeyRelease>', update_list)
        industry_combo['values'] = full_list

        ttk.Label(popup, text="Chi tiết ngành (nếu có):").pack(pady=5)
        chi_tiet_var = tk.StringVar()
        chi_tiet_entry = ttk.Entry(popup, textvariable=chi_tiet_var, width=90)
        chi_tiet_entry.pack(pady=5)

        main_industry_var = tk.BooleanVar(value=False)
        main_industry_check = ttk.Checkbutton(popup, text="Ngành chính", variable=main_industry_var)
        main_industry_check.pack(pady=5)

        def confirm_add():
            selected_industry = industry_var.get()
            chi_tiet = chi_tiet_var.get().strip()
            if not selected_industry:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn mã ngành và tên ngành!")
                return
            try:
                ma_nganh, ten_nganh = selected_industry.split(" - ", 1)
            except ValueError:
                messagebox.showerror("Lỗi", "Định dạng mã ngành không hợp lệ!")
                return

            if chi_tiet:
                ten_nganh = f"{ten_nganh} - {chi_tiet}"

            selected_name = self.app.load_data_var.get()
            for entry in self.app.saved_entries:
                if entry["name"] == selected_name:
                    # Xác định danh sách ngành theo tab hiện tại
                    if current_tab == "Ngành bổ sung":
                        industries = entry["data"].setdefault("nganh_bo_sung", [])
                    else:  # Mặc định là "Ngành nghề kinh doanh"
                        industries = entry["data"].setdefault("nganh_nghe", [])

                    # Kiểm tra trùng lặp mã ngành
                    if any(industry["ma_nganh"] == ma_nganh for industry in industries):
                        messagebox.showwarning("Cảnh báo", "Mã ngành này đã tồn tại!")
                        return

                    # Nếu ngành mới là ngành chính, bỏ chọn ngành chính cũ
                    if main_industry_var.get():
                        for existing_industry in industries:
                            if existing_industry.get("la_nganh_chinh", False):
                                existing_industry["la_nganh_chinh"] = False

                    # Tạo ngành mới
                    new_industry = {
                        "ma_nganh": ma_nganh,
                        "ten_nganh": ten_nganh,
                        "la_nganh_chinh": main_industry_var.get()
                    }

                    # Thêm ngành mới vào danh sách
                    industries.append(new_industry)

                    # Lưu vào SQLite
                    self.app.config_manager.db_manager.save_entry(
                        self.app.config_manager.current_config_name,
                        selected_name,
                        entry["data"]
                    )
                    
                    # Cập nhật saved_entries từ database
                    self.app.saved_entries = self.app.config_manager.db_manager.get_entries(
                        self.app.config_manager.current_config_name
                    )
                    
                    # Tải lại dữ liệu cho tất cả các tab
                    if current_tab == "Ngành bổ sung":
                        self.load_additional_industry_data()
                        # Đồng bộ với tab "Ngành nghề kinh doanh"
                        self.sync_main_industry_tab(
                            action="add",
                            current_tab=current_tab,
                            updated_industry=new_industry
                        )
                        self.load_industry_data()  # Tải lại tab chính
                    elif current_tab == "Ngành nghề kinh doanh":
                        self.load_industry_data()
                    elif current_tab == "Ngành giảm":
                        self.load_removed_industry_data()
                    elif current_tab == "Ngành điều chỉnh":
                        self.load_adjusted_industry_data()
                    
                    popup.destroy()
                    break

        ttk.Button(popup, text="Thêm", command=confirm_add, style="primary-outline").pack(side="bottom", padx=10, pady=10)

    def delete_industry(self):
        """Xóa nhiều ngành nghề đã chọn."""
        tree, industries = self.get_current_tab_tree_and_data()
        if not tree or not industries:
            messagebox.showwarning("Cảnh báo", "Không tìm thấy dữ liệu ngành nghề!")
            return

        selected_items = tree.selection()
        if not selected_items:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn ít nhất một ngành nghề để xóa!")
            return

        if messagebox.askyesno("Xác nhận", f"Bạn có muốn xóa {len(selected_items)} ngành nghề đã chọn không?"):
            indices = [tree.index(item) for item in selected_items]
            indices.sort(reverse=True)  # Xóa từ cuối danh sách để tránh lỗi chỉ số
            removed_industries = []
            selected_name = self.app.load_data_var.get()
            
            for entry in self.app.saved_entries:
                if entry["name"] == selected_name:
                    # Xóa ngành nghề theo chỉ số
                    for idx in indices:
                        try:
                            removed_industries.append(industries.pop(idx))
                        except IndexError:
                            continue
                    
                    # Lưu vào SQLite
                    self.app.config_manager.db_manager.save_entry(
                        self.app.config_manager.current_config_name,
                        selected_name,
                        entry["data"]
                    )
                    
                    # Cập nhật saved_entries từ database
                    self.app.saved_entries = self.app.config_manager.db_manager.get_entries(
                        self.app.config_manager.current_config_name
                    )
                    
                    # Tải lại dữ liệu
                    self.load_industry_data_for_current_tab(tree, industries)
                    
                    
                        
                    # Đồng bộ với tab ngành nghề chính nếu cần
                    current_tab = self.app.notebook.tab(self.app.notebook.select(), "text")
                    if current_tab == "Ngành bổ sung":
                        for removed_industry in removed_industries:
                            self.sync_main_industry_tab(
                                action="delete",
                                current_tab=current_tab,
                                updated_industry=removed_industry
                            )

                    self.load_industry_data()
                    if hasattr(self.app, 'additional_industry_tree'):
                        self.load_additional_industry_data()
                    if hasattr(self.app, 'removed_industry_tree'):
                        self.load_removed_industry_data()
                    if hasattr(self.app, 'adjusted_industry_tree'):
                        self.load_adjusted_industry_data()
                        
                    break

    def edit_industry(self, event=None):
        """Sửa chi tiết ngành nghề với bố cục tương tự Thêm ngành và checkbox Ngành chính."""
        # Lấy từ SQLite:
        all_industries = self.app.config_manager.db_manager.get_industry_codes()

        # Xác định Treeview và danh sách ngành dựa trên tab hiện tại
        tree, industries = self.get_current_tab_tree_and_data()
        if not tree or not industries:
            messagebox.showwarning("Cảnh báo", "Không tìm thấy dữ liệu ngành nghề!")
            return

        # Kiểm tra xem ngành đã được chọn chưa
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn ngành nghề để sửa!")
            return

        # Lấy chỉ số ngành được chọn
        idx = tree.index(selected_item)
        try:
            industry = industries[idx]
        except IndexError:
            messagebox.showwarning("Cảnh báo", "Không tìm thấy ngành nghề để sửa!")
            return

        # Tạo popup sửa ngành
        popup = create_popup(self.app.root, "Sửa ngành nghề", 600, 250)
        ttk.Label(popup, text="Chọn mã ngành và tên ngành:").pack(pady=10)
        industry_var = tk.StringVar()

        # Tách chi tiết (nếu có) khỏi tên ngành
        ten_nganh = industry["ten_nganh"]
        chi_tiet = ""
        if " - " in ten_nganh:
            base_ten_nganh, chi_tiet = ten_nganh.rsplit(" - ", 1)
            for code in all_industries:
                if code["ten_nganh"] == base_ten_nganh:
                    industry_var.set(f"{industry['ma_nganh']} - {base_ten_nganh}")
                    break
        else:
            industry_var.set(f"{industry['ma_nganh']} - {ten_nganh}")
        
        industry_combo = ttk.Combobox(popup, textvariable=industry_var, width=90)
        industry_combo.pack(pady=10)
        
        # Danh sách đầy đủ các mã ngành và tên ngành
        full_list = [f"{code['ma_nganh']} - {code['ten_nganh']}" for code in all_industries]
        
        # Hàm cập nhật danh sách lọc
        def update_list(*args):
            search_term = industry_var.get().lower()  # Lấy chuỗi tìm kiếm và chuyển thành chữ thường
            if search_term:
                # Lọc các mục chứa chuỗi tìm kiếm (trong mã ngành hoặc tên ngành)
                filtered_list = [item for item in full_list if search_term in item.lower()]
            else:
                # Nếu không có chuỗi tìm kiếm, hiển thị toàn bộ danh sách
                filtered_list = full_list
            industry_combo['values'] = filtered_list  # Cập nhật danh sách trong combobox
        
        # Gắn sự kiện <KeyRelease> để gọi hàm update_list khi người dùng gõ
        industry_combo.bind('<KeyRelease>', update_list)
        
        # Khởi tạo combobox với danh sách đầy đủ ban đầu
        industry_combo['values'] = full_list
        
        ttk.Label(popup, text="Chi tiết ngành (nếu có):").pack(pady=5)
        chi_tiet_var = tk.StringVar(value=chi_tiet)
        chi_tiet_entry = ttk.Entry(popup, textvariable=chi_tiet_var, width=90)
        chi_tiet_entry.pack(pady=10)

        # Thêm checkbox "Ngành chính"
        main_industry_var = tk.BooleanVar(value=industry.get("la_nganh_chinh", False))
        main_industry_check = ttk.Checkbutton(popup, text="Ngành chính", variable=main_industry_var)
        main_industry_check.pack(pady=5)
        
        def confirm_edit():
            selected_industry = industry_var.get()
            chi_tiet = chi_tiet_var.get().strip()
            if not selected_industry:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn mã ngành và tên ngành!")
                return
            ma_nganh, ten_nganh = selected_industry.split(" - ", 1)
            if chi_tiet:
                ten_nganh = f"{ten_nganh} - {chi_tiet}"
            for entry in self.app.saved_entries:
                if entry["name"] == self.app.load_data_var.get():
                    # Nếu ngành được sửa thành ngành chính, bỏ chọn ngành chính cũ
                    if main_industry_var.get() and not industry.get("la_nganh_chinh", False):
                        for i, existing_industry in enumerate(industries):
                            if i != idx and existing_industry.get("la_nganh_chinh", False):
                                existing_industry["la_nganh_chinh"] = False
                    industries[idx] = {"ma_nganh": ma_nganh, "ten_nganh": ten_nganh, "la_nganh_chinh": main_industry_var.get()}
                    
                    # Lưu vào SQLite
                    self.app.config_manager.db_manager.save_entry(
                        self.app.config_manager.current_config_name,
                        entry["name"],
                        entry["data"]
                    )
                    
                    # Cập nhật saved_entries từ database
                    self.app.saved_entries = self.app.config_manager.db_manager.get_entries(
                        self.app.config_manager.current_config_name
                    )

                    # Tải lại dữ liệu cho tab hiện tại
                    self.load_industry_data_for_current_tab(tree, industries)
                    self.sync_main_industry_tab(
                        action="edit",
                        current_tab=self.app.notebook.tab(self.app.notebook.select(), "text"),
                        updated_industry={"ma_nganh": ma_nganh, "ten_nganh": ten_nganh, "la_nganh_chinh": main_industry_var.get()},
                        original_industry=industry
                    )
                    self.load_industry_data()
                    if hasattr(self.app, 'additional_industry_tree'):
                        self.load_additional_industry_data()
                    if hasattr(self.app, 'removed_industry_tree'):
                        self.load_removed_industry_data()
                    if hasattr(self.app, 'adjusted_industry_tree'):
                        self.load_adjusted_industry_data()

                    popup.destroy()
                    break
        
        ttk.Button(popup, text="Lưu", command=confirm_edit, style="primary-outline").pack(side="bottom", padx=10, pady=10, expand=True)
 
    def view_industry_details(self, event=None):
        """Hiển thị chi tiết thông tin ngành nghề."""
        tree, industries = self.get_current_tab_tree_and_data()
        if not tree or not industries:
            messagebox.showwarning("Cảnh báo", "Không tìm thấy dữ liệu ngành nghề!")
            return

        selected_items = tree.selection()
        if not selected_items:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn ngành nghề để xem chi tiết!")
            return
            
        # Kiểm tra nếu chọn nhiều ngành nghề
        if len(selected_items) > 1:
            messagebox.showwarning("Cảnh báo", "Vui lòng chỉ chọn một ngành nghề để xem chi tiết!")
            return
            
        selected_item = selected_items[0]  # Lấy item đầu tiên từ selection
        idx = tree.index(selected_item)
        try:
            industry = industries[idx]
        except IndexError:
            messagebox.showwarning("Cảnh báo", "Không tìm thấy ngành nghề để xem chi tiết!")
            return

        # Tạo popup chi tiết ngành nghề
        popup = create_popup(self.app.root, "Chi tiết ngành nghề", 600, 250)
        ttk.Label(popup, text="Tên ngành:").pack(pady=5)
        ttk.Label(popup, text=industry.get("ten_nganh", ""), wraplength=450).pack(pady=5)

        ttk.Label(popup, text="Mã ngành:").pack(pady=5)
        ttk.Label(popup, text=industry.get("ma_nganh", "")).pack(pady=5)

        ttk.Label(popup, text="Ngành chính:").pack(pady=5)
        ttk.Label(popup, text="Có" if industry.get("la_nganh_chinh", False) else "Không").pack(pady=5)

        ttk.Button(popup, text="Đóng", command=popup.destroy, style="secondary-outline").pack(side="right", padx=5, expand=True)
        ttk.Button(popup, text="Chỉnh sửa", command=lambda: [popup.destroy(), self.edit_industry()], style="primary-outline").pack(side="left", padx=5, expand=True)
    
    def set_main_industry(self):
        """Đặt ngành nghề được chọn làm ngành chính."""
        selected = self.app.industry_tree.selection()
        if not selected:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn ngành nghề để đặt làm ngành chính!")
            return
        
        idx = self.app.industry_tree.index(selected[0])
        selected_name = self.app.load_data_var.get()
        
        for entry in self.app.saved_entries:
            if entry["name"] == selected_name:
                # Đặt tất cả các ngành không phải là ngành chính
                for industry in entry["data"].get("nganh_nghe", []):
                    industry["la_nganh_chinh"] = False
                
                # Đặt ngành được chọn làm ngành chính
                if 0 <= idx < len(entry["data"].get("nganh_nghe", [])):
                    entry["data"]["nganh_nghe"][idx]["la_nganh_chinh"] = True
                
                # Lưu vào SQLite
                self.app.config_manager.db_manager.save_entry(
                    self.app.config_manager.current_config_name,
                    selected_name,
                    entry["data"]
                )
                # Cập nhật saved_entries từ database
                self.app.saved_entries = self.app.config_manager.db_manager.get_entries(
                    self.app.config_manager.current_config_name
                )
                # Tải lại dữ liệu
                self.load_industry_data()
                break

