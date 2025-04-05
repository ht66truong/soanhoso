import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import os
import logging

from modules.utils import number_to_words


class DataManager:
    def __init__(self, app):
        self.app = app

    def add_entry_data(self):
        """Tạo một bản ghi dữ liệu mới trong cơ sở dữ liệu."""
        if not self.app.config_manager.current_config_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn cấu hình trước!")
            return
        
        entry_name = simpledialog.askstring("Thêm dữ liệu", "Nhập tên cho bộ dữ liệu mới:")
        if not entry_name:
            return
            
        # Kiểm tra tên có trùng không
        if entry_name in [entry["name"] for entry in self.app.saved_entries]:
            messagebox.showwarning("Cảnh báo", "Tên này đã tồn tại, vui lòng chọn tên khác!")
            return
        
        # Tạo dữ liệu trống
        data = {field: "" for field in self.app.fields}
        data["nganh_nghe"] = []
        data["thanh_vien"] = []
        
        # Lưu vào SQLite
        self.app.config_manager.db_manager.save_entry(
            self.app.config_manager.current_config_name,
            entry_name,
            data
        )
        
        # Cập nhật danh sách đã lưu từ cơ sở dữ liệu
        self.app.saved_entries = self.app.config_manager.db_manager.get_entries(
            self.app.config_manager.current_config_name
        )
        
        # Cập nhật dropdown
        self.app.load_data_dropdown["values"] = [entry["name"] for entry in self.app.saved_entries]
        self.app.load_data_dropdown.set(entry_name)
        
        # Hiển thị thông báo thành công
        messagebox.showinfo("Thành công", f"Đã thêm dữ liệu '{entry_name}'!")
        logging.info(f"Thêm dữ liệu '{entry_name}'")

    def save_entry_data(self):
        """Lưu dữ liệu từ các trường nhập liệu."""
        selected_name = self.app.load_data_var.get()
        
        # Nếu không có tên được chọn, tạo tên mới
        if not selected_name:
            # Kiểm tra nếu người dùng thực sự nhập liệu
            current_data = {field: self.app.entries[field].get() for field in self.app.entries}
            if any(value.strip() for value in current_data.values()):  # Nếu có dữ liệu được nhập
                # Hỏi người dùng đặt tên cho dữ liệu
                new_name = simpledialog.askstring("Tên dữ liệu", "Nhập tên cho dữ liệu mới:")
                if not new_name:
                    messagebox.showwarning("Cảnh báo", "Vui lòng nhập tên cho dữ liệu!")
                    return
                    
                # Kiểm tra tên có trùng không
                for entry in self.app.saved_entries:
                    if entry["name"] == new_name:
                        messagebox.showwarning("Cảnh báo", "Tên này đã tồn tại, vui lòng chọn tên khác!")
                        return
                        
                selected_name = new_name
                self.app.load_data_var.set(selected_name)
            else:
                return  # Không có dữ liệu, không cần lưu

        if not self.app.config_manager.current_config_name:
            return

        # Thu thập dữ liệu từ tất cả các trường nhập liệu
        data = {field: self.app.entries[field].get() for field in self.app.entries}

        # Tự động thêm von_đieu_le_bang_chu vào dữ liệu
        if "von_đieu_le" in data:
            data["von_đieu_le_bang_chu"] = number_to_words(data["von_đieu_le"])

        # Tự động thêm so_tien_bang_chu vào dữ liệu
        if "so_tien" in data:
            data["so_tien_bang_chu"] = number_to_words(data["so_tien"])

         # Lấy danh sách ngành nghề từ industry_tree
        if hasattr(self.app, 'industry_tree'):
            industries = []
            for item in self.app.industry_tree.get_children():
                values = self.app.industry_tree.item(item)['values']
                if len(values) >= 3:  # Đảm bảo đủ 3 cột
                    industry = {
                        "ten_nganh": values[0],
                        "ma_nganh": values[1],
                        "la_nganh_chinh": values[2] == "Có"
                    }
                    industries.append(industry)
            data["nganh_nghe"] = industries  # Cập nhật danh sách ngành nghề
        
        # Lấy danh sách thành viên từ member_tree
        if hasattr(self.app, 'member_tree'):
            members = []
            for item in self.app.member_tree.get_children():
                values = self.app.member_tree.item(item)['values']
                member = {}
                for i, col in enumerate(self.app.displayed_columns):
                    if i < len(values):  # Đảm bảo không vượt quá số lượng giá trị
                        if col == "la_chu_tich":
                            member[col] = (values[i] == "X")
                        else:
                            member[col] = values[i]
                members.append(member)
            data["thanh_vien"] = members  # Cập nhật danh sách thành viên

        # Lưu vào SQLite thông qua DatabaseManager
        self.app.config_manager.db_manager.save_entry(
            self.app.config_manager.current_config_name,
            selected_name,
            data
        )
        
        # Cập nhật danh sách đã lưu từ cơ sở dữ liệu
        self.app.saved_entries = self.app.config_manager.db_manager.get_entries(
            self.app.config_manager.current_config_name
        )
        
        # Cập nhật dropdown
        self.app.load_data_dropdown["values"] = [entry["name"] for entry in self.app.saved_entries]
        
        # Hiển thị thông báo thành công
        if selected_name not in [entry["name"] for entry in self.app.saved_entries]:
            messagebox.showinfo("Thành công", f"Đã thêm dữ liệu '{selected_name}'!")
            logging.info(f"Thêm dữ liệu '{selected_name}'")
        else:
            logging.info(f"Lưu dữ liệu '{selected_name}'")

    def delete_entry_data(self):
        """Xóa bản ghi dữ liệu từ cơ sở dữ liệu."""
        if not self.app.config_manager.current_config_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn cấu hình trước!")
            return
            
        selected_name = self.app.load_data_var.get()
        if not selected_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn dữ liệu để xóa!")
            return
            
        if messagebox.askyesno("Xác nhận", f"Bạn có chắc muốn xóa dữ liệu '{selected_name}' không?"):
            # Xóa dữ liệu từ SQLite
            self.app.config_manager.db_manager.delete_entry(
                self.app.config_manager.current_config_name,
                selected_name
            )
            
            # Cập nhật danh sách đã lưu từ cơ sở dữ liệu
            self.app.saved_entries = self.app.config_manager.db_manager.get_entries(
                self.app.config_manager.current_config_name
            )
            
            # Cập nhật dropdown
            self.app.load_data_dropdown["values"] = [entry["name"] for entry in self.app.saved_entries]
            self.app.load_data_dropdown.set("")
            
            # Xóa dữ liệu khỏi các trường nhập liệu
            for entry in self.app.entries.values():
                entry.delete(0, tk.END)
                
            # Hiển thị thông báo thành công
            messagebox.showinfo("Thành công", f"Dữ liệu '{selected_name}' đã được xóa!")
            logging.info(f"Xóa dữ liệu '{selected_name}'")

    def rename_entry_data(self):
        """Đổi tên bản ghi dữ liệu trong cơ sở dữ liệu."""
        if not self.app.config_manager.current_config_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn cấu hình trước!")
            return
            
        selected_name = self.app.load_data_var.get()
        if not selected_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn dữ liệu để sửa tên!")
            return
            
        new_name = simpledialog.askstring("Sửa tên dữ liệu", "Nhập tên mới:", initialvalue=selected_name)
        if not new_name or new_name == selected_name:
            return
            
        # Kiểm tra tên mới có trùng không
        if new_name in [entry["name"] for entry in self.app.saved_entries]:
            messagebox.showwarning("Cảnh báo", "Tên này đã tồn tại, vui lòng chọn tên khác!")
            return
            
        # Đổi tên trong SQLite
        self.app.config_manager.db_manager.rename_entry(
            self.app.config_manager.current_config_name,
            selected_name,
            new_name
        )
        
        # Cập nhật danh sách đã lưu từ cơ sở dữ liệu
        self.app.saved_entries = self.app.config_manager.db_manager.get_entries(
            self.app.config_manager.current_config_name
        )
        
        # Cập nhật dropdown
        self.app.load_data_dropdown["values"] = [entry["name"] for entry in self.app.saved_entries]
        self.app.load_data_dropdown.set(new_name)
        
        # Hiển thị thông báo thành công
        messagebox.showinfo("Thành công", f"Đã đổi tên dữ liệu thành '{new_name}'!")
        logging.info(f"Đổi tên dữ liệu từ '{selected_name}' thành '{new_name}'")

    def clear_entries(self):
        """Xóa thông tin trong tab hiện tại."""
        # Lấy tab hiện tại
        current_tab = self.app.notebook.tab(self.app.notebook.select(), "text")
        
        # Lấy danh sách các trường trong tab hiện tại
        fields_in_current_tab = self.app.field_groups.get(current_tab, [])
        
        # Xóa dữ liệu trong các trường thuộc tab hiện tại
        for field in fields_in_current_tab:
            if field in self.app.entries:
                self.app.entries[field].delete(0, tk.END)
        
        logging.info(f"Đã xóa thông tin trong tab '{current_tab}'")

    def add_entry_data_from_import(self, entry_data):
        """Thêm dữ liệu từ file Excel vào danh sách."""
        if not self.app.config_manager.current_config_name:
            return

        entry_name = entry_data.get("ten_doanh_nghiep", f"Dữ liệu {len(self.app.saved_entries) + 1}")
        if any(entry["name"] == entry_name for entry in self.app.saved_entries):
            entry_name = f"{entry_name}_{len(self.app.saved_entries) + 1}"

        self.app.saved_entries.append({"name": entry_name, "data": entry_data})
        self.app.config_manager.configs[self.app.config_manager.current_config_name]["entries"] = self.app.saved_entries
        self.app.load_data_dropdown["values"] = [entry["name"] for entry in self.app.saved_entries]
        self.app.config_manager.save_configs()
        logging.info(f"Thêm dữ liệu từ file Excel: {entry_name}")

class TemplateManager:
    def __init__(self, app):
        self.app = app

    def update_template_tree(self):
        """Cập nhật cây template dựa trên cấu hình hiện tại."""
        self.app.template_tree.delete(*self.app.template_tree.get_children())
        if not self.app.config_manager.current_config_name:
            return
        
        templates = self.app.config_manager.configs[self.app.config_manager.current_config_name].get("templates", {})
        for template_name, template_path in templates.items():
            self.app.template_tree.insert("", "end", text=template_name, values=(template_path,))

    def drop_template_files(self, event):
        """Xử lý thả file template vào ứng dụng."""
        if not self.app.config_manager.current_config_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn cấu hình trước!")
            return
        
        # Lấy danh sách đường dẫn file từ sự kiện Drop
        files = self.app.template_frame.tk.splitlist(event.data)
        # Lọc ra chỉ các file .docx
        docx_files = [f for f in files if f.lower().endswith('.docx') and not f.startswith('~$')]
        
        if not docx_files:
            messagebox.showwarning("Cảnh báo", "Không có file .docx hợp lệ!")
            return
        
        templates = self.app.config_manager.configs[self.app.config_manager.current_config_name].setdefault("templates", {})
        
        # Thêm các template mới
        for file_path in docx_files:
            template_name = os.path.basename(file_path)
            # Sao chép file vào thư mục templates
            dest_path = os.path.join(self.app.templates_dir, template_name)
            try:
                import shutil
                shutil.copy2(file_path, dest_path)
                templates[template_name] = dest_path
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể sao chép file {template_name}: {str(e)}")
        
        # Lưu cấu hình với templates mới
        self.app.config_manager.db_manager.save_config(
            self.app.config_manager.current_config_name,
            self.app.config_manager.configs[self.app.config_manager.current_config_name]
        )
        
        # Cập nhật cây template
        self.update_template_tree()
        messagebox.showinfo("Thành công", f"Đã thêm {len(docx_files)} template!")

    def add_multiple_templates(self):
        """Thêm nhiều template, sao chép file vào thư mục templates và tránh ghi đè."""
        if not self.app.config_manager.current_config_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn cấu hình trước!")
            return
        template_paths = filedialog.askopenfilenames(filetypes=[("Word files", "*.docx")], title="Chọn nhiều template Word")
        if template_paths:
            added_count = 0
            templates = self.app.config_manager.configs[self.app.config_manager.current_config_name].get("templates", {})
            for template_path in template_paths:
                if os.path.exists(template_path):
                    template_name = os.path.basename(template_path)
                    base_name = os.path.splitext(template_name)[0]
                    extension = os.path.splitext(template_name)[1]
                    new_name = base_name + extension
                    counter = 1
                    target_path = os.path.join(self.app.templates_dir, new_name)
                    while os.path.exists(target_path):
                        new_name = f"{base_name}_{counter}{extension}"
                        target_path = os.path.join(self.app.templates_dir, new_name)
                        counter += 1
                    while new_name in templates:
                        new_name = f"{base_name}_{counter}{extension}"
                        target_path = os.path.join(self.app.templates_dir, new_name)
                        counter += 1
                    try:
                        import shutil
                        shutil.copy2(template_path, target_path)
                        templates[new_name] = new_name
                        added_count += 1
                    except Exception as e:
                        logging.error(f"Lỗi khi sao chép template {template_path}: {str(e)}")
                        messagebox.showerror("Lỗi", f"Không thể sao chép template: {str(e)}")
            if added_count > 0:
                self.app.config_manager.configs[self.app.config_manager.current_config_name]["templates"] = templates
                self.app.config_manager.save_configs()
                self.update_template_tree()
                messagebox.showinfo("Thành công", f"Đã thêm {added_count} template vào cấu hình '{self.app.config_manager.current_config_name}'!")
                logging.info(f"Thêm {added_count} template vào cấu hình '{self.app.config_manager.current_config_name}'")
   
    def delete_template(self):
        """Xóa template đã chọn."""
        selected = self.app.template_tree.selection()
        if not selected:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn template để xóa!")
            return
        
        # Lấy tên các template đã chọn
        selected_templates = [self.app.template_tree.item(item)["text"] for item in selected]
        
        # Hiển thị thông báo xác nhận dựa vào số lượng template được chọn
        confirm_message = "Bạn có chắc muốn xóa template này?" if len(selected_templates) == 1 else f"Bạn có chắc muốn xóa {len(selected_templates)} template đã chọn?"
        
        if messagebox.askyesno("Xác nhận", confirm_message):
            templates = self.app.config_manager.configs[self.app.config_manager.current_config_name].get("templates", {})
            deleted_count = 0
            
            # Xóa từng template được chọn
            for template_name in selected_templates:
                if template_name in templates:
                    del templates[template_name]
                    deleted_count += 1
            
            # Lưu cấu hình với templates mới
            self.app.config_manager.db_manager.save_config(
                self.app.config_manager.current_config_name,
                self.app.config_manager.configs[self.app.config_manager.current_config_name]
            )
            
            # Cập nhật cây template
            self.update_template_tree()
            
            # Hiển thị thông báo thành công
            if deleted_count == 1:
                messagebox.showinfo("Thành công", f"Đã xóa template '{selected_templates[0]}'!")
            else:
                messagebox.showinfo("Thành công", f"Đã xóa {deleted_count} template!")
            
            logging.info(f"Đã xóa {deleted_count} template")

    def show_template_context_menu(self, event):
        context_menu = tk.Menu(self.app.root, tearoff=0)
        context_menu.add_command(label="Thêm template", command=self.add_multiple_templates)
        context_menu.add_command(label="Xóa", command=self.delete_template)
        context_menu.post(event.x_root, event.y_root)

    # kéo thả templates
    def start_drag(self, event):
        item = self.app.template_tree.identify_row(event.y)
        if item:
            self.app.drag_item = item

    def drag_template(self, event):
        if self.app.drag_item:
            self.app.template_tree.selection_set(self.app.drag_item)

    def drop_template(self, event):
        """Xử lý khi thả template để thay đổi thứ tự."""
        if self.app.drag_item:
            target = self.app.template_tree.identify_row(event.y)
            if target and target != self.app.drag_item:
                # Lấy thông tin template
                drag_text = self.app.template_tree.item(self.app.drag_item, "text")
                drag_values = self.app.template_tree.item(self.app.drag_item, "values")
                
                # Xóa template cũ
                self.app.template_tree.delete(self.app.drag_item)
                
                # Xác định vị trí mới
                target_index = self.app.template_tree.index(target)
                
                # Chèn vào vị trí mới
                self.app.template_tree.insert("", target_index, text=drag_text, values=drag_values)
                
                # Cập nhật thứ tự trong cấu hình
                templates = {}
                for item in self.app.template_tree.get_children():
                    text = self.app.template_tree.item(item, "text")
                    values = self.app.template_tree.item(item, "values")
                    templates[text] = values[0]
                
                self.app.config_manager.configs[self.app.config_manager.current_config_name]["templates"] = templates
                
                # Lưu cấu hình
                self.app.config_manager.db_manager.save_config(
                    self.app.config_manager.current_config_name,
                    self.app.config_manager.configs[self.app.config_manager.current_config_name]
                )
            
            # Reset drag_item
            self.app.drag_item = None

class FieldManager:
    def __init__(self, app):
        self.app = app

    def add_field(self):
        if not self.app.config_manager.current_config_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn cấu hình trước!")
            return
        new_field = simpledialog.askstring("Thêm trường", "Nhập tên trường mới:")
        if new_field:
            new_field = new_field.strip()
            if not new_field:
                messagebox.showwarning("Cảnh báo", "Tên trường không được để trống!")
                return
            current_tab = self.app.notebook.tab(self.app.notebook.select(), "text")  # Lấy tab hiện tại từ notebook
            if not current_tab:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn tab trước khi thêm trường!")
                return
            if new_field in self.app.field_groups[current_tab]:
                messagebox.showwarning("Cảnh báo", "Trường này đã tồn tại trong tab hiện tại!")
                return
            self.app.field_groups[current_tab].append(new_field)
            self.app.config_manager.configs[self.app.config_manager.current_config_name]["field_groups"][current_tab] = self.app.field_groups[current_tab]
            self.app.config_manager.save_configs()
            self.app.config_manager.load_selected_config(None)
            logging.info(f"Thêm trường '{new_field}' vào tab '{current_tab}'")
            
    def delete_selected_field(self):
        if not self.app.config_manager.current_config_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn cấu hình trước!")
            return
        field = self.app.field_var.get()
        if not field:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn trường để xóa!")
            return
        current_tab = self.app.notebook.tab(self.app.notebook.select(), "text")
        if field in self.app.field_groups[current_tab]:
            self.app.field_groups[current_tab].remove(field)
            self.app.config_manager.configs[self.app.config_manager.current_config_name]["field_groups"][current_tab] = self.app.field_groups[current_tab]
            self.app.config_manager.save_configs()
            self.app.config_manager.load_selected_config(None)
            logging.info(f"Xóa trường '{field}' khỏi tab '{current_tab}'")

    def rename_selected_field(self):
        if not self.app.config_manager.current_config_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn cấu hình trước!")
            return
        old_field = self.app.field_var.get()
        if not old_field:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn trường để sửa tên!")
            return
        new_field = simpledialog.askstring("Sửa tên trường", "Nhập tên mới cho trường:", initialvalue=old_field)
        if new_field and new_field != old_field:
            current_tab = self.app.notebook.tab(self.app.notebook.select(), "text")
            if new_field in self.app.field_groups[current_tab]:
                messagebox.showwarning("Cảnh báo", "Tên trường đã tồn tại trong tab này!")
                return
            idx = self.app.field_groups[current_tab].index(old_field)
            self.app.field_groups[current_tab][idx] = new_field
            self.app.config_manager.configs[self.app.config_manager.current_config_name]["field_groups"][current_tab] = self.app.field_groups[current_tab]
            self.app.config_manager.save_configs()
            self.app.config_manager.load_selected_config(None)
            logging.info(f"Sửa tên trường từ '{old_field}' thành '{new_field}' trong tab '{current_tab}'")

    def rename_field(self, field):
        if not self.app.config_manager.current_config_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn cấu hình trước!")
            return
        old_field = field
        new_field = simpledialog.askstring("Sửa tên trường", "Nhập tên mới cho trường:", initialvalue=old_field)
        if new_field and new_field != old_field:
            current_tab = self.app.notebook.tab(self.app.notebook.select(), "text")
            if new_field in self.app.field_groups[current_tab]:
                messagebox.showwarning("Cảnh báo", "Tên trường đã tồn tại trong tab này!")
                return
            idx = self.app.field_groups[current_tab].index(old_field)
            self.app.field_groups[current_tab][idx] = new_field
            self.app.config_manager.configs[self.app.config_manager.current_config_name]["field_groups"][current_tab] = self.app.field_groups[current_tab]
            self.app.config_manager.save_configs()
            self.app.config_manager.load_selected_config(None)
            logging.info(f"Sửa tên trường từ '{old_field}' thành '{new_field}' trong tab '{current_tab}'")

    def delete_field(self, field):
        if not self.app.config_manager.current_config_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn cấu hình trước!")
            return
        current_tab = self.app.notebook.tab(self.app.notebook.select(), "text")
        if field in self.app.field_groups[current_tab]:
            self.app.field_groups[current_tab].remove(field)
            self.app.config_manager.configs[self.app.config_manager.current_config_name]["field_groups"][current_tab] = self.app.field_groups[current_tab]
            self.app.config_manager.save_configs()
            self.app.config_manager.load_selected_config(None)
            logging.info(f"Xóa trường '{field}' khỏi tab '{current_tab}'")

class TabManager:
    def __init__(self, app):
        self.app = app

    def create_tabs(self):
        """Tạo các tab dựa trên field_groups, bao gồm tab Thông tin thành viên và Ngành nghề kinh doanh."""
        tab_names = list(self.app.field_groups.keys())
        for i, tab_name in enumerate(tab_names):
            if tab_name == "Ngành nghề kinh doanh":
                self.app.industry_manager.create_industry_tab()
            elif tab_name == "Ngành bổ sung":
                self.app.industry_manager.create_additional_industry_tab()
            elif tab_name == "Ngành giảm":
                self.app.industry_manager.create_removed_industry_tab()
            elif tab_name == "Ngành điều chỉnh":
                self.app.industry_manager.create_adjusted_industry_tab()
            elif tab_name == "Thông tin thành viên":
                self.app.member_manager.create_member_tab()
            else:
                tab = ttk.Frame(self.app.notebook)
                self.app.notebook.add(tab, text=tab_name)
                canvas = tk.Canvas(tab, bg="#ffffff", highlightthickness=0)
                scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
                scrollable_frame = ttk.Frame(canvas, padding=5)
                scrollable_frame.bind("<Configure>", lambda e, c=canvas: c.configure(scrollregion=c.bbox("all")))
                canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
                canvas.configure(yscrollcommand=scrollbar.set)
                canvas.pack(side="left", fill="both", expand=True)
                scrollbar.pack(side="right", fill="y")
                
                scrollable_frame.grid_columnconfigure(0, weight=1)
                scrollable_frame.grid_columnconfigure(1, weight=2)
                
                # Đảm bảo mỗi tab có đủ trường để kiểm tra cuộn (tạm thời thêm nhiều trường giả)
                fields = self.app.field_groups[tab_name]
                for j, field in enumerate(fields):
                    display_name = field.replace('_', ' ').title()
                    label_width = max(60, len(display_name) // 2)
                    label = ttk.Label(scrollable_frame, text=f"{display_name}:", width=label_width, anchor="e", wraplength=500)
                    label.grid(row=j, column=0, padx=5, pady=2, sticky="e")
                    self.app.labels[field] = label
                    entry = ttk.Entry(scrollable_frame, width=50)
                    entry.grid(row=j, column=1, padx=5, pady=2, sticky="ew")
                    self.app.entries[field] = entry

                    # Gắn sự kiện tự động lưu
                    entry.bind("<KeyRelease>", lambda event: self.app.data_manager.save_entry_data())
                    #entry.bind("<FocusOut>", lambda event: self.app.data_manager.save_entry_data())

                    # Gọi phương thức để thêm menu ngữ cảnh cho ô nhập liệu
                    self.app.add_entry_context_menu(entry)

                    # Lấy menu ngữ cảnh hiện có của entry
                    context_menu = entry.context_menu if hasattr(entry, 'context_menu') else tk.Menu(self.app.root, tearoff=0)
                    entry.context_menu = context_menu

                    # Thêm các tùy chọn quản lý trường vào menu ngữ cảnh
                    context_menu.add_separator()
                    context_menu.add_command(label="Thêm trường", image=self.app.add_icon_img, compound="left", command=self.app.field_manager.add_field)
                    context_menu.add_command(label="Xóa trường", image=self.app.delete_icon_img, compound="left", command=lambda f=field: self.app.field_manager.delete_field(f))
                    context_menu.add_command(label="Sửa tên trường", image=self.app.edit_icon_img, compound="left", command=lambda f=field: self.app.field_manager.rename_field(f))
                    
                    # Hàm hiển thị menu ngữ cảnh
                    def show_context_menu(event, menu=context_menu):
                        menu.post(event.x_root, event.y_root)

                    # Gán menu ngữ cảnh cho cả label và entry
                    label.bind("<Button-3>", show_context_menu)
                    entry.bind("<Button-3>", show_context_menu)

                    if field == "von_đieu_le":
                        def update_von_dieu_le_bang_chu(event):
                            von_dieu_le_value = self.app.entries["von_đieu_le"].get()
                            von_dieu_le_bang_chu = number_to_words(von_dieu_le_value)
                            if "von_đieu_le_bang_chu" in self.app.entries:
                                self.app.entries["von_đieu_le_bang_chu"].delete(0, tk.END)
                                self.app.entries["von_đieu_le_bang_chu"].insert(0, von_dieu_le_bang_chu)
                        entry.bind("<KeyRelease>", update_von_dieu_le_bang_chu)

                    if field == "so_tien":
                        def update_so_tien_bang_chu(event):
                            so_tien_value = self.app.entries["so_tien"].get()
                            so_tien_bang_chu = number_to_words(so_tien_value)
                            if "so_tien_bang_chu" in self.app.entries:
                                self.app.entries["so_tien_bang_chu"].delete(0, tk.END)
                                self.app.entries["so_tien_bang_chu"].insert(0, so_tien_bang_chu)
                        entry.bind("<KeyRelease>", update_so_tien_bang_chu)

                # Thêm dropdown chọn nhân viên nếu là tab "Thông tin uỷ quyền"
                if tab_name == "Thông tin uỷ quyền" and hasattr(self.app, 'employee_manager'):
                    self.app.employee_manager.add_employee_dropdown_to_tab(scrollable_frame)

                # Thêm sự kiện cuộn chuột cho từng canvas riêng biệt
                def on_mousewheel(event, c=canvas):  # Truyền canvas cụ thể vào hàm
                    c.yview_scroll(int(-1 * (event.delta / 120)), "units")
                canvas.bind("<Enter>", lambda e, c=canvas: c.bind_all("<MouseWheel>", lambda evt: on_mousewheel(evt, c)))
                canvas.bind("<Leave>", lambda e, c=canvas: c.unbind_all("<MouseWheel>"))

        if self.app.current_tab_index < len(tab_names):
            self.app.notebook.select(self.app.current_tab_index)
            
        self.app.update_field_dropdown()  # Cập nhật danh sách trường khi tạo tab

    def add_tab(self):
        tab_name = simpledialog.askstring("Thêm tab", "Nhập tên tab mới:")
        if tab_name and tab_name not in self.app.field_groups:
            self.app.field_groups[tab_name] = []
            self.app.config_manager.configs[self.app.config_manager.current_config_name]["field_groups"] = self.app.field_groups
            self.app.config_manager.save_configs()
            self.app.clear_tabs()
            self.create_tabs()
            self.app.tab_dropdown["values"] = list(self.app.field_groups.keys())
            self.app.tab_dropdown.set(tab_name)
            messagebox.showinfo("Thành công", f"Đã thêm tab '{tab_name}'!")
            logging.info(f"Thêm tab '{tab_name}'")

    def delete_tab(self):
        selected_tab = self.app.tab_var.get()
        if not selected_tab:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn tab để xóa!")
            return
        if messagebox.askyesno("Xác nhận", f"Bạn có chắc muốn xóa tab '{selected_tab}' không?"):
            # Safely remove fields from self.app.fields if they exist
            for field in self.app.field_groups[selected_tab]:
                if hasattr(self.app, 'fields') and field in self.app.fields:
                    self.app.fields.remove(field)
            
            # Delete the tab from field_groups
            del self.app.field_groups[selected_tab]
            
            # Update configuration
            self.app.config_manager.configs[self.app.config_manager.current_config_name]["field_groups"] = self.app.field_groups
            self.app.config_manager.save_configs()
            
            # Recreate tabs
            self.clear_tabs()
            self.create_tabs()
            
            # Update dropdown
            self.app.tab_dropdown["values"] = list(self.app.field_groups.keys())
            self.app.tab_dropdown.set(list(self.app.field_groups.keys())[0] if self.app.field_groups else "")
            
            # Display success message
            messagebox.showinfo("Thành công", f"Đã xóa tab '{selected_tab}'!")
            logging.info(f"Xóa tab '{selected_tab}'")

    def rename_tab(self):
        old_tab = self.app.tab_var.get()
        if not old_tab:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn tab để sửa tên!")
            return
        new_tab = simpledialog.askstring("Sửa tên tab", "Nhập tên mới:", initialvalue=old_tab)
        if new_tab and new_tab != old_tab and new_tab not in self.app.field_groups:
            self.app.field_groups[new_tab] = self.app.field_groups.pop(old_tab)
            self.app.config_manager.configs[self.app.config_manager.current_config_name]["field_groups"] = self.app.field_groups
            self.app.config_manager.save_configs()
            self.clear_tabs()
            self.create_tabs()
            self.app.tab_dropdown["values"] = list(self.app.field_groups.keys())
            self.app.tab_dropdown.set(new_tab)
            messagebox.showinfo("Thành công", f"Đã đổi tên tab thành '{new_tab}'!")
            logging.info(f"Đổi tên tab từ '{old_tab}' thành '{new_tab}'")

    def clear_tabs(self): 
        """Xóa tất cả các tab trong notebook."""
        current_tab_index = self.app.notebook.index("current") if self.app.notebook.tabs() else 0
        for tab in self.app.notebook.tabs():
            self.app.notebook.forget(tab)
        self.app.notebook.update_idletasks()
        self.app.entries.clear()
        self.app.labels.clear()
        self.app.current_tab_index = current_tab_index

