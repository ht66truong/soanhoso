from ttkbootstrap.constants import *
from tkinterdnd2 import TkinterDnD
from modules.gui import DataEntryApp
import sqlite3
import os


def check_sqlite_version():
    """Kiểm tra phiên bản SQLite và hiển thị thông báo nếu không đủ."""
    try:
        sqlite_version = sqlite3.sqlite_version_info
        min_version = (3, 24, 0)  # Phiên bản tối thiểu cho các tính năng mới
        
        if sqlite_version < min_version:
            print(f"Cảnh báo: Phiên bản SQLite hiện tại ({sqlite3.sqlite_version}) thấp hơn phiên bản khuyến nghị ({'.'.join(map(str, min_version))}).")
            print("Một số tính năng có thể không hoạt động đúng.")
    except:
        print("Không thể kiểm tra phiên bản SQLite.")

if __name__ == "__main__":
    # Kiểm tra môi trường SQLite
    check_sqlite_version()
    
    # Đảm bảo thư mục AppData tồn tại
    os.makedirs("AppData", exist_ok=True)
    
    # Khởi động ứng dụng
    root = TkinterDnD.Tk()
    app = DataEntryApp(root)
    root.mainloop()
    