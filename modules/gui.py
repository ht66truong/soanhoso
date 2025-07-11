# Thư viện chuẩn
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import sys
import logging
from datetime import datetime

# Thư viện bên thứ ba
from PIL import Image, ImageTk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinterdnd2 import DND_FILES

# Module nội bộ
from modules.utils import ToolTip, create_popup
from modules.export import ExportManager
from modules.member import MemberManager
from modules.industry import IndustryManager
from modules.config import ConfigManager, BackupManager
from modules.data import DataManager, FieldManager, TabManager, TemplateManager
from modules.employee import EmployeeManager

# Thiết lập thư mục AppData và logging
appdata_dir = "AppData"
if not os.path.exists(appdata_dir):
    os.makedirs(appdata_dir)
logging.basicConfig(filename=os.path.join(appdata_dir, "app.log"), level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logging.getLogger("docxtpl").setLevel(logging.ERROR)


class DataEntryApp:
    def __init__(self, root):
        """Khởi tạo ứng dụng nhập liệu hồ sơ kinh doanh."""
        self.root = root
        self.root.title("Ứng dụng nhập liệu hồ sơ Kinh Doanh")
        window_width = 1280
        window_height = 720
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        #self.root.state('zoomed')  # Tùy chọn: Phóng to tối đa khi khởi động     
        self.root.configure(bg="#f8f9fa")
        logging.info("Ứng dụng khởi động")

        # Cấu hình file trong thư mục AppData
        self.appdata_dir = appdata_dir
        self.configs_file = os.path.join(self.appdata_dir, "configs.json")
        self.templates_dir = "templates"
        self.backup_dir = "backup"

        if not os.path.exists(self.templates_dir):
            os.makedirs(self.templates_dir)
        if not os.path.exists(self.backup_dir):
            os.makedirs(self.backup_dir)

        self.default_fields = ["tên_doanh_nghiệp", "tên_nước_ngoài", "tên_viết_tắt", "số_nhà_tên_đường", "xã_phường", "tỉnh_thành_phố", 
                               "số_điện_thoại", "vốn_điều_lệ",
                               "họ_tên", "giới_tính", "ngày_sinh", "số_cccd", "ddpl_số_nhà_tên_đường", "ddpl_xã_phường", "ddpl_tỉnh_thành_phố", "quốc_gia",
                               "họ_tên_uq", "giới_tính_uq", "ngày_sinh_uq", "số_cccd_uq", "địa_chỉ_liên_lạc_uq", "sdt_uq", "email_uq"]
        # Định nghĩa các cột cho bảng Thành viên/Chủ sở hữu
        self.member_columns = ["ho_ten", "gioi_tinh", "ngay_sinh", "so_cccd", "so_nha_ten_duong", "xa_phuong", "tinh_thanh_pho", "quoc_gia", "von_gop", "ty_le_gop", "ngay_gop_von"]

        self.configs = {}
        self.current_config_name = None
        self.field_groups = {}
        self.fields = []
        self.saved_entries = []
        self.drag_item = None
        self.entries = {}
        self.labels = {}
        self.current_tab_index = 0

        # Initialize manager classes first
        self.config_manager = ConfigManager(self)
        self.field_manager = FieldManager(self)
        self.tab_manager = TabManager(self)
        self.data_manager = DataManager(self)
        self.member_manager = MemberManager(self)
        self.industry_manager = IndustryManager(self)
        self.template_manager = TemplateManager(self)
        self.export_manager = ExportManager(self)
        self.backup_manager = BackupManager(self)
        self.employee_manager = EmployeeManager(self)

        # Styles
        style = ttk.Style()
        style.theme_use("flatly")
        style.configure("TButton", font=("Segoe UI", 10), borderwidth=1, relief="raised", padding=5)

        def resource_path(relative_path): # Hàm hỗ trợ lấy đường dẫn tài nguyên
            """Lấy đường dẫn tuyệt đối đến tài nguyên, hoạt động cho cả mã nguồn và file thực thi PyInstaller."""
            if hasattr(sys, '_MEIPASS'):
                return os.path.join(sys._MEIPASS, relative_path)
            return os.path.join(os.path.abspath("."), relative_path)

        # Load icons
        try:
            add_icon = Image.open(resource_path("icon/add_icon.png")).resize((20, 20), Image.Resampling.LANCZOS)
            delete_icon = Image.open(resource_path("icon/delete_icon.png")).resize((20, 20), Image.Resampling.LANCZOS)
            save_icon = Image.open(resource_path("icon/save_icon.png")).resize((20, 20), Image.Resampling.LANCZOS)
            clear_icon = Image.open(resource_path("icon/clear_icon.png")).resize((20, 20), Image.Resampling.LANCZOS)
            import_icon = Image.open(resource_path("icon/imxls_icon.png")).resize((20, 20), Image.Resampling.LANCZOS)
            xls_icon = Image.open(resource_path("icon/xls_icon.png")).resize((20, 20), Image.Resampling.LANCZOS)
            preview_icon = Image.open(resource_path("icon/preview_icon.png")).resize((40, 40), Image.Resampling.LANCZOS)
            export_icon = Image.open(resource_path("icon/export_icon.png")).resize((40, 40), Image.Resampling.LANCZOS)
            remove_icon = Image.open(resource_path("icon/remove_icon.png")).resize((20, 20), Image.Resampling.LANCZOS)
            edit_icon = Image.open(resource_path("icon/edit_icon.png")).resize((20, 20), Image.Resampling.LANCZOS)
            search_icon = Image.open(resource_path("icon/search_icon.png")).resize((20, 20), Image.Resampling.LANCZOS)
            find_icon = Image.open(resource_path("icon/find_icon.png")).resize((20, 20), Image.Resampling.LANCZOS)
            restorebackup_icon = Image.open(resource_path("icon/restorebackup_icon.png")).resize((20, 20), Image.Resampling.LANCZOS)
            self.add_icon_img = ImageTk.PhotoImage(add_icon)
            self.delete_icon_img = ImageTk.PhotoImage(delete_icon)
            self.save_icon_img = ImageTk.PhotoImage(save_icon)
            self.clear_icon_img = ImageTk.PhotoImage(clear_icon)
            self.import_icon_img = ImageTk.PhotoImage(import_icon)
            self.xls_icon_img = ImageTk.PhotoImage(xls_icon)
            self.preview_icon_img = ImageTk.PhotoImage(preview_icon)
            self.export_icon_img = ImageTk.PhotoImage(export_icon)
            self.remove_icon_img = ImageTk.PhotoImage(remove_icon)
            self.edit_icon_img = ImageTk.PhotoImage(edit_icon)
            self.search_icon_img = ImageTk.PhotoImage(search_icon)
            self.find_icon_img = ImageTk.PhotoImage(find_icon)
            self.restorebackup_icon_img = ImageTk.PhotoImage(restorebackup_icon)
        except Exception as e:
            logging.error(f"Lỗi khi tải biểu tượng: {str(e)}")
            messagebox.showerror("Lỗi", "Không thể tải biểu tượng, kiểm tra thư mục icon!")

        # Top frame
        self.top_frame = ttk.Frame(self.root)
        self.top_frame.pack(side="top", fill="x", padx=5, pady=5)
        
        self.main_frame = ttk.Frame(root, padding=5)
        self.main_frame.pack(fill="both", expand=True)

        # Config Frame
        self.config_frame = ttk.LabelFrame(self.main_frame, text="Quản lý cấu hình", padding=5)
        self.config_frame.pack(fill="x", pady=(0, 5))
        ttk.Label(self.config_frame, text="Chọn loại hình:").pack(side="left", padx=5)
        self.config_var = tk.StringVar()
        self.config_dropdown = ttk.Combobox(self.config_frame, textvariable=self.config_var, values=list(self.configs.keys()), state="readonly", width=25)
        self.config_dropdown.pack(side="left", padx=5)
        ToolTip(self.config_dropdown, "Chọn cấu hình làm việc\nNhấp chuột phải để mở menu quản lý cấu hình")
        self.config_dropdown.bind("<<ComboboxSelected>>", self.load_selected_config)
        # Thêm bind chuột phải để hiển thị menu ngữ cảnh cấu hình
        self.config_dropdown.bind("<Button-3>", self.show_config_context_menu)
        
        # Thêm một công tắc để ẩn hiện các mục quản lý tab và trường
        ttk.Separator(self.config_frame, orient="vertical").pack(side="left", padx=5, fill="y")
        self.advanced_mode = tk.BooleanVar(value=False)  # Mặc định không hiển thị
        self.toggle_button = ttk.Checkbutton(self.config_frame, text="Hiện quản lý tab & trường", 
            variable=self.advanced_mode, command=self.toggle_advanced_controls, bootstyle="round-toggle")
        self.toggle_button.pack(side="left", padx=5)
        ToolTip(self.toggle_button, "Bật/tắt hiển thị các điều khiển quản lý tab và trường")
        
        # Tạo một frame riêng cho các điều khiển quản lý tab và trường
        self.advanced_frame = ttk.Frame(self.config_frame)
        
        # Quản lý tab - Di chuyển vào advanced_frame
        ttk.Separator(self.advanced_frame, orient="vertical").pack(side="left", padx=5, fill="y")
        ttk.Label(self.advanced_frame, text="Quản lý tab:").pack(side="left", padx=5)
        self.tab_var = tk.StringVar()
        self.tab_dropdown = ttk.Combobox(self.advanced_frame, textvariable=self.tab_var, state="readonly", width=20)
        ToolTip(self.tab_dropdown, "Chọn tab để thêm/sửa/xóa\nNhấp chuột phải để mở menu quản lý tab")
        self.tab_dropdown.pack(side="left", padx=5)
        self.tab_dropdown.bind("<Button-3>", self.show_tab_context_menu)
        
        # Quản lý trường - Di chuyển vào advanced_frame
        ttk.Separator(self.advanced_frame, orient="vertical").pack(side="left", padx=5, fill="y")
        ttk.Label(self.advanced_frame, text="Quản lý trường:").pack(side="left", padx=5)
        self.field_var = tk.StringVar()
        self.field_dropdown = ttk.Combobox(self.advanced_frame, textvariable=self.field_var, state="readonly", width=20)
        ToolTip(self.field_dropdown, "Chọn trường để thêm/sửa/xóa\nNhấp chuột phải để mở menu quản lý trường")
        self.field_dropdown.pack(side="left", padx=5)
        self.field_dropdown.bind("<Button-3>", self.show_field_context_menu)
        
        # Quản lý template - Vẫn nằm bên ngoài advanced_frame để luôn hiển thị
        ttk.Separator(self.config_frame, orient="vertical").pack(side="left", padx=5, fill="y")
        template_button = ttk.Button(self.config_frame, text="Quản lý danh sách mẫu", command=self.show_template_manager_popup, bootstyle="info-outline")
        template_button.pack(side="left", padx=5)
        ToolTip(template_button, "Nhấp để mở cửa sổ quản lý mẫu\nBạn có thể thêm, xóa, kéo thả sắp xếp mẫu")
        
        # Đảm bảo template_tree vẫn được tạo nhưng không hiển thị trên giao diện chính
        self.template_tree = ttk.Treeview(self.root, columns=(), show="tree", height=4, selectmode="extended")
        self.template_tree.column("#0", width=250)
        self.update_template_tree()

        # Control Frame
        self.control_frame = ttk.LabelFrame(self.main_frame, text="Quản lý dữ liệu", padding=5)
        self.control_frame.pack(fill="x", pady=5)

        # Thêm khung tìm kiếm
        search_frame = ttk.Frame(self.control_frame)
        search_frame.grid(row=0, column=0, columnspan=2, pady=5, sticky="w")
        ttk.Label(search_frame, text="Tìm kiếm:").pack(side="left", padx=5)
        self.search_var = tk.StringVar()
        self.search_combobox = ttk.Combobox(search_frame, textvariable=self.search_var, width=60, state="normal")
        self.search_combobox.pack(side="left", padx=5)
        ToolTip(self.search_combobox, "Tìm kiếm hoặc chọn dữ liệu\nNhập tên để tìm kiếm công ty")
        self.search_combobox["values"] = [entry["name"] for entry in self.saved_entries]
        self.search_combobox.bind("<<ComboboxSelected>>", self.load_data_from_search)
        self.search_combobox.bind("<KeyRelease>", self.update_search_suggestions)
        
        # Thêm biến load_data_var để tránh lỗi AttributeError
        self.load_data_var = self.search_var
        
        # Tạo nút quản lý dữ liệu với menu ngữ cảnh thay vì nhiều nút
        data_menu_button = ttk.Button(search_frame, image=self.find_icon_img, bootstyle="success-outline")
        data_menu_button.pack(side="left", padx=5)
        ToolTip(data_menu_button, "Nhấp để tìm kiếm\nNhấp chuột phải để mở menu quản lý dữ liệu")
        data_menu_button.bind("<Button-1>", lambda e: self.search_data())
        data_menu_button.bind("<Button-3>", self.show_data_context_menu)

        ttk.Separator(self.control_frame, orient="vertical").grid(row=0, column=3, padx=5, pady=5, sticky="ns")
        
        # Gộp nút xuất nhập dữ liệu vào một nút dropdown
        data_tools_frame = ttk.Frame(self.control_frame)
        data_tools_frame.grid(row=0, column=4, padx=5, pady=5)
        
        # Tạo menu cho MenuButton
        data_tools_menu = tk.Menu(self.root, tearoff=0)
        data_tools_menu.add_command(label="Xuất dữ liệu ra Excel", command=self.export_data)
        data_tools_menu.add_command(label="Nhập dữ liệu từ Excel", command=self.import_from_file)
        data_tools_menu.add_command(label="Danh sách Placeholder", command=self.show_placeholder_popup)
        data_tools_menu.add_separator()
        data_tools_menu.add_command(label="Khôi phục backup", command=self.restore_from_backup)
        data_tools_menu.add_separator()
        data_tools_menu.add_command(label="Thông tin ứng dụng", command=self.show_about_info)
        
        # Thay thế Button thông thường bằng MenuButton
        data_tools_button = ttk.Menubutton(data_tools_frame, text="Công cụ dữ liệu", bootstyle="primary-outline-menubutton", width=15, menu=data_tools_menu)
        data_tools_button.pack(side="left", padx=5)
        ToolTip(data_tools_button, "Nhấp để mở các công cụ quản lý dữ liệu")

        # Notebook với tiêu đề "Quản lý nhập liệu"
        self.notebook_frame = ttk.LabelFrame(self.main_frame, text="Quản lý nhập liệu", padding=5)
        self.notebook_frame.pack(fill="both", expand=True, pady=5)
        self.notebook = ttk.Notebook(self.notebook_frame)
        self.notebook.pack(fill="both", expand=True)
        # Gắn sự kiện khi tab thay đổi
        self.notebook.bind("<<NotebookTabChanged>>", lambda event: self.update_field_dropdown())
        
        # Export Frame
        self.export_frame = ttk.Frame(self.main_frame)
        self.export_frame.pack(fill="x", pady=5)
        self.export_frame.grid_columnconfigure(0, weight=1)
        self.export_frame.grid_columnconfigure(3, weight=1)
        self.preview_word_button = ttk.Button(self.export_frame, image=self.preview_icon_img, command=self.preview_word, bootstyle="secondary-outline")
        ToolTip(self.preview_word_button, "Xem trước")
        self.preview_word_button.grid(row=0, column=1, padx=5)
        self.export_file_button = ttk.Button(self.export_frame, image=self.export_icon_img, command=self.export_file, bootstyle="success-outline")
        ToolTip(self.export_file_button, "Xuất file")
        self.export_file_button.grid(row=0, column=2, padx=5)
       
        # Initialize default config if none exists
        if not self.config_manager.configs:
            self.config_manager.initialize_default_config()
            
        # Update dropdown with configs
        self.config_dropdown["values"] = list(self.config_manager.configs.keys())
        if self.config_manager.configs:
            config_name = list(self.config_manager.configs.keys())[0]
            self.config_dropdown.set(config_name)
            self.config_manager.current_config_name = config_name
            self.config_manager.load_selected_config(None)

        self.root.after(600000, self.backup_manager.auto_backup)


    def load_selected_config(self, event):
        self.config_manager.load_selected_config(event)
        self.initialize_dropdowns()  # Cập nhật dropdowns sau khi tải cấu hình mới

    def load_selected_entry(self, event):
        self.data_manager.load_selected_entry(event)

    def update_field_dropdown(self):
        self.field_manager.update_field_dropdown()

    def add_new_config(self):
        self.config_manager.add_new_config()

    def delete_config(self):
        self.config_manager.delete_current_config()  # Change from delete_config

    def rename_config(self):
        self.config_manager.rename_current_config()  # Change from rename_config

    def add_field(self):
        self.field_manager.add_field()

    def delete_selected_field(self):
        self.field_manager.delete_selected_field()

    def rename_selected_field(self):
        self.field_manager.rename_selected_field()

    def rename_field(self, field):
        self.field_manager.rename_field(field)

    def delete_field(self, field):
        self.field_manager.delete_field(field)

    def add_tab(self):
        self.tab_manager.add_tab()

    def delete_tab(self):
        self.tab_manager.delete_tab()

    def rename_tab(self):
        self.tab_manager.rename_tab()

    def clear_entries(self):
        self.data_manager.clear_entries()

    def clear_tabs(self): 
        self.tab_manager.clear_tabs()

    def add_entry_data(self):
         self.data_manager.add_entry_data()

    def save_entry_data(self):
        self.data_manager.save_entry_data()

    def delete_entry_data(self):
        self.data_manager.delete_entry_data()

    def rename_entry_data(self):
        self.data_manager.rename_entry_data()

    def create_member_tab(self):
        self.member_manager.create_member_tab()

    def load_member_data(self):
        self.member_manager.load_member_data()

    def add_member(self):
        self.member_manager.add_member()

    def delete_member(self):
        self.member_manager.delete_member()

    def edit_member(self):
        self.member_manager.edit_member()

    def view_member_details(self, event=None):
        self.member_manager.view_member_details(event)

    def start_drag_member(self, event):
        self.member_manager.start_drag_member(event)

    def drag_member(self, event):
        self.member_manager.drag_member(event)

    def drop_member(self, event):
        self.member_manager.drop_member(event)

    def create_industry_tab(self):
        self.industry_manager.create_industry_tab()

    def load_industry_data(self):
        self.industry_manager.load_industry_data()

    def add_industry(self):
        self.industry_manager.add_industry()

    def delete_industry(self):
        self.industry_manager.delete_industry()

    def edit_industry(self, event=None):
        self.industry_manager.edit_industry(event)

    def view_industry_details(self, event=None):
        self.industry_manager.view_industry_details(event)

    def set_main_industry(self):
        self.industry_manager.set_main_industry()

    def update_template_tree(self):
        self.template_manager.update_template_tree()

    def show_template_manager_popup(self):
        self.template_manager.show_template_manager_popup()
        
    def drop_template_files_to_popup(self, event, popup_tree, popup):
        self.template_manager.drop_template_files_to_popup(event, popup_tree, popup)
        
    def delete_template_from_popup(self, popup_tree):
        self.template_manager.delete_template_from_popup(popup_tree)

    def add_template_from_popup(self, popup_tree):
        self.template_manager.add_template_from_popup(popup_tree)

    def export_data(self):
        self.export_manager.export_data()

    def import_from_file(self):
        self.export_manager.import_from_file()

    def show_placeholder_popup(self):
        self.export_manager.show_placeholder_popup()

    def copy_placeholder_from_popup(self, placeholder_list):
        self.export_manager.copy_placeholder_from_popup(placeholder_list)

    def export_placeholders(self):
        self.export_manager.export_placeholders()

    def check_template_placeholders(self, doc_paths, data_lower):
        self.export_manager.check_template_placeholders(doc_paths, data_lower)

    def preview_word(self):
        self.export_manager.preview_word()

    def export_file(self):
        self.export_manager.export_file()

    def show_export_popup(self, export_type):
        self.export_manager.show_export_popup(export_type)

    def export_preview(self, doc_paths, data_lower, mode):
        self.export_manager.export_preview(doc_paths, data_lower, mode)

    def export_to_word(self, doc_paths, data_lower, mode):
        self.export_manager.export_to_word(doc_paths, data_lower, mode)

    def export_to_pdf(self, doc_paths, data_lower, mode):
        self.export_manager.export_to_pdf(doc_paths, data_lower, mode)

    def auto_backup(self):
        self.backup_manager.auto_backup()

    def restore_from_backup(self):
        self.backup_manager.restore_from_backup()

    def update_field_dropdown(self):
        """Cập nhật danh sách trường trong dropdown dựa trên tab hiện tại."""
        if not self.notebook.tabs():  # Check if there are no tabs
            self.field_dropdown["values"] = []
            self.field_var.set("")
            return

        current_tab = self.notebook.tab(self.notebook.select(), "text")
        fields = self.field_groups.get(current_tab, [])
        self.field_dropdown["values"] = fields
        self.field_var.set(fields[0] if fields else "")
        
    def update_all_dropdowns(self):
        """Cập nhật tất cả dropdown trong ứng dụng."""
        # Cập nhật dropdown tìm kiếm
        entries = self.saved_entries
        if entries:
            self.search_combobox["values"] = [entry["name"] for entry in entries]
        else:
            # Lấy dữ liệu từ cơ sở dữ liệu nếu saved_entries trống
            if hasattr(self, 'config_manager') and self.config_manager.current_config_name:
                entries = self.config_manager.db_manager.get_entries(
                    self.config_manager.current_config_name
                )
                self.saved_entries = entries
                self.search_combobox["values"] = [entry["name"] for entry in entries]
        
        # Cập nhật các dropdown khác trong ứng dụng
        # Danh sách trường và tab
        self.update_field_dropdown()
        
        # Cập nhật tab_dropdown nếu có
        if hasattr(self, 'tab_dropdown'):
            if self.notebook.tabs():
                self.tab_dropdown["values"] = [self.notebook.tab(tab_id, "text") for tab_id in self.notebook.tabs()]
                if self.tab_dropdown["values"]:
                    self.tab_var.set(self.tab_dropdown["values"][0])
        
        logging.info("Đã cập nhật tất cả dropdown trong ứng dụng")

    def initialize_dropdowns(self):
        """Khởi tạo và tải dữ liệu cho tất cả các dropdown, không phụ thuộc vào tab."""
        # Đảm bảo cơ sở dữ liệu được truy cập
        if hasattr(self, 'config_manager') and self.config_manager.current_config_name:
            # Tải tất cả dữ liệu từ cấu hình hiện tại
            entries = self.config_manager.db_manager.get_entries(
                self.config_manager.current_config_name
            )
            self.saved_entries = entries
            
            # Cập nhật dropdown tìm kiếm
            self.search_combobox["values"] = [entry["name"] for entry in entries]
            
            # Cập nhật các dropdown khác
            self.update_all_dropdowns()
            
            logging.info("Đã khởi tạo tất cả dropdown")
        
    def add_entry_context_menu(self, entry):
        """Thêm menu ngữ cảnh cho ô nhập liệu."""
        context_menu = tk.Menu(self.root, tearoff=0)
        context_menu.add_command(label="Sao chép", command=lambda: entry.event_generate("<<Copy>>"))
        context_menu.add_command(label="Dán", command=lambda: entry.event_generate("<<Paste>>"))
        context_menu.add_command(label="Xóa", command=lambda: entry.delete(0, tk.END))
        context_menu.add_command(label="Xóa trắng", command=self.clear_entries)
        # Lưu menu vào thuộc tính context_menu của entry
        entry.context_menu = context_menu

    def show_field_context_menu(self, event):
        """Hiển thị menu ngữ cảnh cho field_dropdown."""
        # Thêm menu ngữ cảnh cho field_dropdown thay vì nhiều nút
        field_context_menu = tk.Menu(self.root, tearoff=0)
        field_context_menu.add_command(label="Thêm trường", command=self.add_field)
        field_context_menu.add_command(label="Xóa trường", command=self.delete_selected_field)
        field_context_menu.add_command(label="Sửa tên trường", command=self.rename_selected_field)
        field_context_menu.tk_popup(event.x_root, event.y_root)

    def show_tab_context_menu(self, event):
        """Hiển thị menu ngữ cảnh cho tab_dropdown."""
        selected_tab = self.tab_var.get()
        if selected_tab:
            # Tạo menu ngữ cảnh cho các tab
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label="Thêm tab", command=self.add_tab)
            menu.add_command(label="Xóa tab", command=self.delete_tab)
            menu.add_command(label="Sửa tên tab", command=self.rename_tab)
            menu.tk_popup(event.x_root, event.y_root)

    def search_data(self):
        """Tìm kiếm dữ liệu dựa trên từ khóa."""
        keyword = self.search_var.get().lower()
        if not keyword:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập từ khóa tìm kiếm!")
            return

        # Lọc dữ liệu dựa trên từ khóa
        results = []
        for entry in self.saved_entries:
            if keyword in entry["name"].lower() or any(keyword in str(value).lower() for value in entry["data"].values()):
                results.append(entry)

        if not results:
            messagebox.showinfo("Kết quả", "Không tìm thấy dữ liệu phù hợp!")
        else:
            # Hiển thị kết quả tìm kiếm (có thể cập nhật Treeview hoặc popup)
            self.display_search_results(results)
    
    def load_data_from_search(self, event):
        """Tải dữ liệu từ dropdown tìm kiếm."""
        selected_name = self.search_var.get()
        if not selected_name:
            return

        # Đồng bộ giá trị vào load_data_var
        self.load_data_var.set(selected_name)

        # Tìm và tải dữ liệu tương ứng
        for entry in self.saved_entries:
            if entry["name"] == selected_name:
                # Điền dữ liệu vào các trường nhập liệu
                for field, value in entry["data"].items():
                    if field in self.entries:
                        self.entries[field].delete(0, tk.END)
                        self.entries[field].insert(0, value)

                # Tải dữ liệu ngành nghề (nếu có)
                if hasattr(self, 'industry_tree'):
                    self.industry_tree.delete(*self.industry_tree.get_children())
                    for industry in entry["data"].get("nganh_nghe", []):
                        self.industry_tree.insert(
                            "", "end", values=(industry["ten_nganh"], industry["ma_nganh"], "X" if industry["la_nganh_chinh"] else "")
                        )

                # Tải dữ liệu thành viên (nếu có)
                if hasattr(self, 'member_tree'):
                    self.member_tree.delete(*self.member_tree.get_children())
                    for member in entry["data"].get("thanh_vien", []):
                        self.member_tree.insert(
                            "", "end", values=(
                                member.get("ho_ten", ""),
                                member.get("so_cccd", ""),
                                member.get("von_gop", ""),
                                member.get("ty_le_gop", ""),
                                member.get("la_chu_tich", "")
                            )
                        )

                # Xóa nội dung trong ô tìm kiếm
                #self.search_var.set("")
                messagebox.showinfo("Thành công", f"Đã tải dữ liệu '{selected_name}'!")
                break

    def display_search_results(self, results):
        """Hiển thị kết quả tìm kiếm trong popup."""
        popup = create_popup(self.root, "Kết quả tìm kiếm", 600, 400)
        tree = ttk.Treeview(popup, columns=("name", "details"), show="headings")
        tree.heading("name", text="Tên")
        tree.heading("details", text="Chi tiết")
        tree.column("name", width=200)
        tree.column("details", width=350)
        tree.pack(fill="both", expand=True,padx=5, pady=5)

        for result in results:
            tree.insert("", "end", values=(result["name"], str(result["data"])))

        def load_selected_data():
            """Tải dữ liệu được chọn vào khu vực 'Quản lý nhập liệu'."""
            selected_item = tree.selection()
            if not selected_item:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn một mục để tải!")
                return
            selected_name = tree.item(selected_item[0])["values"][0]
            for entry in self.saved_entries:
                if entry["name"] == selected_name:
                    self.load_data_var.set(selected_name)
                    self.load_selected_entry(None)
                    popup.destroy()
                    self.search_var.set("")
                    messagebox.showinfo("Thành công", f"Đã tải dữ liệu '{selected_name}'!")
                    break

        # Gắn sự kiện nhấn đúp chuột vào Treeview
        def on_double_click(event):
            load_selected_data()

        tree.bind("<Double-1>", on_double_click)

        ttk.Button(popup, text="Chọn", command=load_selected_data, bootstyle="primary-outline").pack(side="left", padx=5, pady=5)
        ttk.Button(popup, text="Đóng", command=popup.destroy, bootstyle="secondary-outline").pack(side="right", padx=5, pady=5)

    def update_treeview_with_results(self, results):
        """Cập nhật Treeview chính với kết quả tìm kiếm."""
        self.template_tree.delete(*self.template_tree.get_children())
        for result in results:
            self.template_tree.insert("", "end", text=result["name"], values=(result["data"]))
    
    def update_search_suggestions(self, event):
        """Cập nhật danh sách gợi ý trong ô tìm kiếm."""
        input_text = self.search_var.get().lower()
        suggestions = [
            entry["name"]
            for entry in self.saved_entries
            if input_text in entry["name"].lower()
        ]
        self.search_combobox["values"] = suggestions

    def show_data_context_menu(self, event):
        """Hiển thị menu ngữ cảnh cho quản lý dữ liệu."""
        # Tạo menu ngữ cảnh cho quản lý dữ liệu
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Thêm mới", command=self.add_entry_data)
        menu.add_command(label="Đổi tên", command=self.rename_entry_data)
        menu.add_command(label="Xóa dữ liệu", command=self.delete_entry_data)
        menu.add_separator()
        menu.add_command(label="Lưu/Cập nhật", command=self.save_entry_data)
        
        
        menu.tk_popup(event.x_root, event.y_root)

    def show_config_context_menu(self, event):
        """Hiển thị menu ngữ cảnh cho config_dropdown."""
        selected_config = self.config_var.get()
        if selected_config:
            # Tạo menu ngữ cảnh cho các cấu hình
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label="Thêm cấu hình", command=self.add_new_config)
            menu.add_command(label="Xóa cấu hình", command=self.delete_config)
            menu.add_command(label="Sửa tên cấu hình", command=self.rename_config)
            menu.tk_popup(event.x_root, event.y_root)

    def show_about_info(self):
        """Hiển thị thông tin về ứng dụng và tác giả."""
        about_text = """
        Ứng dụng soạn hồ sơ Đăng Ký Doanh Nghiệp
        
        Phiên bản: 6.1.0
        Phát hành: 11/07/2025
        
        © 2025 CÔNG TY TNHH GIẢI PHÁP SME
        Bản quyền thuộc về CÔNG TY TNHH GIẢI PHÁP SME
        Mọi quyền được bảo lưu

        Liên hệ:
        Email: lienhe@giaiphapsme.com
        Điện thoại: 02866.866.800
        Website: www.giaiphapsme.com
        """
        messagebox.showinfo("Thông tin ứng dụng", about_text)

    def toggle_advanced_controls(self):
        """Ẩn hiện các điều khiển quản lý tab và trường."""
        if self.advanced_mode.get():
            # Hiển thị các điều khiển
            self.advanced_frame.pack(side="left", fill="y")
            self.toggle_button.config(text="Ẩn quản lý tab & trường")
        else:
            # Ẩn các điều khiển
            self.advanced_frame.pack_forget()
            self.toggle_button.config(text="Hiện quản lý tab & trường")