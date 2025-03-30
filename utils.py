import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog, Toplevel, Text
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docxcompose.composer import Composer
import unicodedata

class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event):
        if self.tooltip:
            self.tooltip.destroy()
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        screen_width = self.widget.winfo_screenwidth()
        screen_height = self.widget.winfo_screenheight()
        if x + 200 > screen_width:
            x = self.widget.winfo_rootx() - 200
        if y + 30 > screen_height:
            y = self.widget.winfo_rooty() - 30
        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        label = ttk.Label(self.tooltip, text=self.text, background="#ffffe0", padding=2, relief="solid", wraplength=200)
        label.pack()
        self.tooltip.transient(self.widget)
        self.tooltip.lift()
        self.tooltip.attributes('-topmost', True)

    def hide_tooltip(self, event):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

def create_centered_popup(parent, title, width, height):
    """Tạo một popup căn giữa màn hình với kích thước và tiêu đề tùy chỉnh."""
    popup = tk.Toplevel(parent)
    popup.title(title)
    
    screen_width = popup.winfo_screenwidth()
    screen_height = popup.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    
    popup.geometry(f"{width}x{height}+{x}+{y}")
    popup.transient(parent)
    popup.grab_set()
    
    return popup

def add_section_break(doc):
    last_paragraph = doc.paragraphs[-1]
    new_section = OxmlElement('w:p')
    break_element = OxmlElement('w:br')
    break_element.set(qn('w:type'), 'page')
    new_section.append(break_element)
    last_paragraph._p.addnext(new_section)

def number_to_words(number):
    """Chuyển đổi số thành chữ tiếng Việt."""
    try:
        number = int(float(str(number).replace(",", "").replace(".", "")))
    except (ValueError, TypeError):
        logging.warning(f"Giá trị không hợp lệ cho von_đieu_le: {number}")
        return "Không hợp lệ (vui lòng nhập số)"

    units = ["", "nghìn", "triệu", "tỷ"]
    digits = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]
    teens = ["mười", "mười một", "mười hai", "mười ba", "mười bốn", "mười lăm", 
             "mười sáu", "mười bảy", "mười tám", "mười chín"]
    tens = ["", "", "hai mươi", "ba mươi", "bốn mươi", "năm mươi", "sáu mươi", 
            "bảy mươi", "tám mươi", "chín mươi"]

    def convert_chunk(chunk):
        if chunk == 0:
            return "không"
        result = ""
        hundreds = chunk // 100
        tens_units = chunk % 100
        if hundreds > 0:
            result += digits[hundreds] + " trăm "
        if tens_units >= 20:
            tens_digit = tens_units // 10
            units_digit = tens_units % 10
            result += tens[tens_digit]
            if units_digit > 0:
                result += " " + digits[units_digit]
        elif tens_units >= 10:
            result += teens[tens_units - 10]
        elif tens_units > 0:
            result += digits[tens_units]
        return result.strip()

    if number == 0:
        return "không đồng"

    result = ""
    chunk_idx = 0
    while number > 0:
        chunk = number % 1000
        if chunk > 0:
            chunk_str = convert_chunk(chunk)
            if chunk_idx > 0:
                chunk_str += " " + units[chunk_idx]
            if result:
                result = chunk_str + " " + result
            else:
                result = chunk_str
        number //= 1000
        chunk_idx += 1

    return result.strip() + " đồng"

def normalize_vietnamese(input_str):
    """Chuẩn hóa chuỗi tiếng Việt thành không dấu và thay khoảng trắng bằng dấu gạch dưới."""
    vietnamese_map = {
        'à': 'a', 'á': 'a', 'ả': 'a', 'ã': 'a', 'ạ': 'a',
        'ă': 'a', 'ằ': 'a', 'ắ': 'a', 'ẳ': 'a', 'ẵ': 'a', 'ặ': 'a',
        'â': 'a', 'ầ': 'a', 'ấ': 'a', 'ẩ': 'a', 'ẫ': 'a', 'ậ': 'a',
        'è': 'e', 'é': 'e', 'ẻ': 'e', 'ẽ': 'e', 'ẹ': 'e',
        'ê': 'e', 'ề': 'e', 'ế': 'e', 'ể': 'e', 'ễ': 'e', 'ệ': 'e',
        'ì': 'i', 'í': 'i', 'ỉ': 'i', 'ĩ': 'i', 'ị': 'i',
        'ò': 'o', 'ó': 'o', 'ỏ': 'o', 'õ': 'o', 'ọ': 'o',
        'ô': 'o', 'ồ': 'o', 'ố': 'o', 'ổ': 'o', 'ỗ': 'o', 'ộ': 'o',
        'ơ': 'o', 'ờ': 'o', 'ớ': 'o', 'ở': 'o', 'ỡ': 'o', 'ợ': 'o',
        'ù': 'u', 'ú': 'u', 'ủ': 'u', 'ũ': 'u', 'ụ': 'u',
        'ư': 'u', 'ừ': 'u', 'ứ': 'u', 'ử': 'u', 'ữ': 'u', 'ự': 'u',
        'ỳ': 'y', 'ý': 'y', 'ỷ': 'y', 'ỹ': 'y', 'ỵ': 'y',
        'đ': 'd',
        'À': 'A', 'Á': 'A', 'Ả': 'A', 'Ã': 'A', 'Ạ': 'A',
        'Ă': 'A', 'Ằ': 'A', 'Ắ': 'A', 'Ẳ': 'A', 'Ẵ': 'A', 'Ặ': 'A',
        'Â': 'A', 'Ầ': 'A', 'Ấ': 'A', 'Ẩ': 'A', 'Ẫ': 'A', 'Ậ': 'A',
        'È': 'E', 'É': 'E', 'Ẻ': 'E', 'Ẽ': 'E', 'Ẹ': 'E',
        'Ê': 'E', 'Ề': 'E', 'Ế': 'E', 'Ể': 'E', 'Ễ': 'E', 'Ệ': 'E',
        'Ì': 'I', 'Í': 'I', 'Ỉ': 'I', 'Ĩ': 'I', 'Ị': 'I',
        'Ò': 'O', 'Ó': 'O', 'Ỏ': 'O', 'Õ': 'O', 'Ọ': 'O',
        'Ô': 'O', 'Ồ': 'O', 'Ố': 'O', 'Ổ': 'O', 'Ỗ': 'O', 'Ộ': 'O',
        'Ơ': 'O', 'Ờ': 'O', 'Ớ': 'O', 'Ở': 'O', 'Ỡ': 'O', 'Ợ': 'O',
        'Ù': 'U', 'Ú': 'U', 'Ủ': 'U', 'Ũ': 'U', 'Ụ': 'U',
        'Ư': 'U', 'Ừ': 'U', 'Ứ': 'U', 'Ử': 'U', 'Ữ': 'U', 'Ự': 'U',
        'Ỳ': 'Y', 'Ý': 'Y', 'Ỷ': 'Y', 'Ỹ': 'Y', 'Ỵ': 'Y',
        'Đ': 'D'
    }
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    normalized = ''.join(vietnamese_map.get(c, c) for c in nfkd_form if not unicodedata.combining(c))
    normalized = normalized.lower().replace(" ", "_").replace(",", "").replace("/", "_").replace("(", "").replace(")", "")
    return normalized
