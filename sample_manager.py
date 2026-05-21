import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
from datetime import datetime
import os
import platform

# 引入用于数据整理和 Excel 输出的专业库
try:
    import pandas as pd
except ImportError:
    pd = None

# ==========================================
# 第一步：数据库初始化（保持结构兼容）
# ==========================================
def init_db():
    conn = sqlite3.connect("coatings_samples.db")
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS sample_requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_name TEXT NOT NULL,
            contact_person TEXT,
            product_model TEXT NOT NULL,
            quantity TEXT,
            status TEXT DEFAULT '已申请',
            date_added TEXT,
            notes TEXT
        )
    ''')
    
    cursor.execute("PRAGMA table_info(sample_requests)")
    columns = [column[1] for column in cursor.fetchall()]
    if "photo_path" not in columns:
        cursor.execute("ALTER TABLE sample_requests ADD COLUMN photo_path TEXT")
        
    if "phone_number" not in columns:
        cursor.execute("ALTER TABLE sample_requests ADD COLUMN phone_number TEXT")
        
    conn.commit()
    conn.close()

# ==========================================
# 第二步：自定义组件 - 自动匹配输入框
# ==========================================
class AutoCompleteEntry(tk.Entry):
    def __init__(self, master, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.bind("<KeyRelease>", self.on_key_release)
        self.listbox = None
        self.history_companies = []

    def set_history(self, companies_list):
        self.history_companies = list(set(companies_list))

    def on_key_release(self, event):
        if event.keysym in ("Up", "Down", "Return", "Escape"):
            return
        value = self.get().strip()
        if not value:
            self.hide_listbox()
            return
        match_data = [item for item in self.history_companies if value.lower() in item.lower()]
        if match_data:
            self.show_listbox(match_data)
        else:
            self.hide_listbox()

    def show_listbox(self, match_data):
        if not self.listbox:
            self.listbox = tk.Listbox(self.master, height=5, exportselection=False)
            self.listbox.bind("<<ListboxSelect>>", self.on_select)
        self.listbox.place(x=self.winfo_x(), y=self.winfo_y() + self.winfo_height(), width=self.winfo_width())
        self.listbox.delete(0, tk.END)
        for item in match_data:
            self.listbox.insert(tk.END, item)
        self.listbox.lift()

    def hide_listbox(self):
        if self.listbox:
            self.listbox.place_forget()

    def on_select(self, event):
        if self.listbox:
            index = self.listbox.curselection()
            if index:
                selected_text = self.listbox.get(index)
                self.delete(0, tk.END)
                self.insert(0, selected_text)
                self.hide_listbox()

# ==========================================
# 第三步：核心业务逻辑
# ==========================================
current_photo_path = ""

def select_photo():
    global current_photo_path
    file_path = filedialog.askopenfilename(
        title="选择样品照片",
        filetypes=[("Image Files", "*.jpg *.jpeg *.png *.bmp")]
    )
    if file_path:
        current_photo_path = file_path
        label_photo_status.config(text="照片已就绪", fg="green")
    else:
        current_photo_path = ""
        label_photo_status.config(text="未选择照片", fg="gray")

def save_request():
    global current_photo_path
    company = entry_company.get().strip()
    model = entry_model.get().strip()
    
    if not company or not model:
        messagebox.showwarning("输入错误", "客户名称和助剂型号为必填项！")
        return
        
    conn = sqlite3.connect("coatings_samples.db")
    cursor = conn.cursor()
    current_date = datetime.now().strftime("%Y-%m-%d %H:%M")
    
    cursor.execute('''
        INSERT INTO sample_requests (company_name, contact_person, phone_number, product_model, quantity, date_added, notes, photo_path)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ''', (company, entry_contact.get(), entry_phone.get(), model, entry_qty.get(), current_date, entry_notes.get(), current_photo_path))
    
    conn.commit()
    conn.close()
    messagebox.showinfo("成功", "样品申请已成功记录！")
    
    current_photo_path = ""
    label_photo_status.config(text="未选择照片", fg="gray")
    clear_inputs()
    show_all_records()

def on_tree_double_click(event):
    selected_item = tree.selection()
    if not selected_item:
        return
    values = tree.item(selected_item)['values']
    if len(values) < 10:
        messagebox.showinfo("提示", "该条历史记录不包含完整的照片数据结构。")
        return
    photo_path = values[9]
    if photo_path and photo_path != "无" and os.path.exists(str(photo_path)):
        if platform.system() == "Windows":
            os.startfile(photo_path)
        elif platform.system() == "Darwin":
            os.system(f"open '{photo_path}'")
        else:
            os.system(f"xdg-open '{photo_path}'")
    else:
        messagebox.showinfo("提示", f"无法打开图片。\n可能原因：该记录未上传照片，或本地图片文件已被移动/删除。\n路径: {photo_path}")

# 【通用化修改】重构 Excel 导出接口的表头与文件名
def export_to_excel_gui():
    if pd is None:
        messagebox.showerror("依赖缺失", "未检测到必要组件！\n请先在终端运行以下命令安装：\npip install pandas openpyxl")
        return
        
    db_path = "coatings_samples.db"
    try:
        conn = sqlite3.connect(db_path)
        query = """
            SELECT id, company_name, contact_person, phone_number, product_model, 
                   quantity, status, date_added, notes, photo_path 
            FROM sample_requests
        """
        df = pd.read_sql_query(query, conn)
        conn.close()
        
        if df.empty:
            messagebox.showwarning("提示", "当前数据库中没有任何数据记录，无法导出！")
            return
            
        # 【通用化修改】将英文字段转换为通用的化工产品助剂表头
        excel_headers = {
            'id': '申请编号',
            'company_name': '客户公司名称',
            'contact_person': '联系人',
            'phone_number': '联系电话',
            'product_model': '产品助剂型号', # 去除“耐磨”字样
            'quantity': '索样数量',
            'status': '当前状态',
            'date_added': '录入时间',
            'notes': '备注/客户反馈',
            'photo_path': '留档照片本地路径'
        }
        df.rename(columns=excel_headers, inplace=True)
        
        # 执行精准的日期排行排序（降序）
        df['_temp_datetime'] = pd.to_datetime(df['录入时间'], errors='coerce')
        df.sort_values(by='_temp_datetime', ascending=False, inplace=True)
        df.drop(columns=['_temp_datetime'], inplace=True)
        
        # 【通用化修改】定义更符合单机版系统特征的默认文件名
        current_date_str = datetime.now().strftime('%Y%m%d')
        default_filename = f"化工产品助剂样品申请表_{current_date_str}.xlsx"
        
        save_file_path = filedialog.asksaveasfilename(
            title="选择 Excel 报表保存位置",
            defaultextension=".xlsx",
            filetypes=[("Excel 电子表格", "*.xlsx")],
            initialfile=default_filename
        )
        
        if save_file_path:
            df.to_excel(save_file_path, index=False, engine='openpyxl')
            messagebox.showinfo("导出成功", f"数据已按日期排行整理完毕！\nExcel 报表已成功保存至：\n{save_file_path}")
            
    except Exception as e:
        messagebox.showerror("导出失败", f"在导出数据转换为 Excel 时遇到了内部错误：\n{str(e)}")

def refresh_autocomplete_source():
    conn = sqlite3.connect("coatings_samples.db")
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT company_name FROM sample_requests")
    companies = [row[0] for row in cursor.fetchall()]
    conn.close()
    entry_company.set_history(companies)

def search_records():
    search_key = entry_search.get().strip()
    conn = sqlite3.connect("coatings_samples.db")
    cursor = conn.cursor()
    if search_key:
        cursor.execute("SELECT * FROM sample_requests WHERE company_name LIKE ?", (f'%{search_key}%',))
    else:
        cursor.execute("SELECT * FROM sample_requests")
    rows = cursor.fetchall()
    conn.close()
    update_treeview(rows)

def update_status():
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("提示", "请先在列表中选择一条记录！")
        return
    new_status = combo_status.get()
    record_id = tree.item(selected_item)['values'][0]
    
    conn = sqlite3.connect("coatings_samples.db")
    cursor = conn.cursor()
    cursor.execute("UPDATE sample_requests SET status = ? WHERE id = ?", (new_status, record_id))
    conn.commit()
    conn.close()
    
    messagebox.showinfo("成功", f"状态已更新为: {new_status}")
    show_all_records()

def show_all_records():
    conn = sqlite3.connect("coatings_samples.db")
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM sample_requests ORDER BY id DESC")
    rows = cursor.fetchall()
    conn.close()
    update_treeview(rows)
    refresh_autocomplete_source()

def update_treeview(rows):
    for item in tree.get_children():
        tree.delete(item)
    for row in rows:
        conn = sqlite3.connect("coatings_samples.db")
        cursor = conn.cursor()
        cursor.execute("SELECT id, company_name, contact_person, phone_number, product_model, quantity, status, date_added, notes, photo_path FROM sample_requests WHERE id=?", (row[0],))
        standard_row = cursor.fetchone()
        conn.close()
        
        if not standard_row:
            continue

        row_list = list(standard_row)
        tree.insert("", "end", values=row_list)

def clear_inputs():
    entry_company.delete(0, tk.END)
    entry_contact.delete(0, tk.END)
    entry_phone.delete(0, tk.END)
    entry_model.delete(0, tk.END)
    entry_qty.delete(0, tk.END)
    entry_notes.delete(0, tk.END)
    entry_company.hide_listbox()

# ==========================================
# 第四步：GUI 界面搭建（通用化 UI 重构）
# ==========================================
root = tk.Tk()
# 【通用化修改】更改软件主窗口标题
root.title("化工产品助剂单机样品申请信息系统")
root.geometry("1050x650")

frame_input = tk.LabelFrame(root, text=" 填写样品申请信息 ", padx=10, pady=10)
frame_input.pack(fill="x", padx=15, pady=10)

tk.Label(frame_input, text="客户公司*:").grid(row=0, column=0, sticky="w", pady=5)
entry_company = AutoCompleteEntry(frame_input, width=22)
entry_company.grid(row=0, column=1, padx=5, pady=5)

tk.Label(frame_input, text="联系人:").grid(row=0, column=2, sticky="w", pady=5)
entry_contact = tk.Entry(frame_input, width=15)
entry_contact.grid(row=0, column=3, padx=5, pady=5)

tk.Label(frame_input, text="联系电话:").grid(row=0, column=4, sticky="w", pady=5)
entry_phone = tk.Entry(frame_input, width=18)
entry_phone.grid(row=0, column=5, padx=5, pady=5)

# 【通用化修改】去除标签中的“耐磨”字样
tk.Label(frame_input, text="助剂型号*:").grid(row=1, column=0, sticky="w", pady=5)
entry_model = tk.Entry(frame_input, width=22)
entry_model.grid(row=1, column=1, padx=5, pady=5)

tk.Label(frame_input, text="索样数量:").grid(row=1, column=2, sticky="w", pady=5)
entry_qty = tk.Entry(frame_input, width=15)
entry_qty.grid(row=1, column=3, padx=5, pady=5)

tk.Label(frame_input, text="样品照片:").grid(row=1, column=4, sticky="w", pady=5)
btn_photo = tk.Button(frame_input, text="上传样品照片", command=select_photo, bg="#E0E0E0")
btn_photo.grid(row=1, column=5, sticky="w", padx=5, pady=5)
label_photo_status = tk.Label(frame_input, text="未选择照片", fg="gray")
label_photo_status.grid(row=1, column=5, padx=95, pady=5, sticky="w")

tk.Label(frame_input, text="备注/要求:").grid(row=2, column=0, sticky="w", pady=5)
entry_notes = tk.Entry(frame_input, width=68)
entry_notes.grid(row=2, column=1, columnspan=5, sticky="w", padx=5, pady=5)

btn_save = tk.Button(frame_input, text="提交申请", command=save_request, bg="#4CAF50", fg="white", width=12)
btn_save.grid(row=2, column=5, padx=10, pady=5, sticky="e")

frame_action = tk.Frame(root, padx=15)
frame_action.pack(fill="x", pady=5)

tk.Label(frame_action, text="查询客户:").pack(side="left")
entry_search = tk.Entry(frame_action, width=15)
entry_search.pack(side="left", padx=5)
btn_search = tk.Button(frame_action, text="搜索", command=search_records, width=6)
btn_search.pack(side="left", padx=5)
btn_refresh = tk.Button(frame_action, text="显示全部", command=show_all_records, width=8)
btn_refresh.pack(side="left", padx=5)

tk.Label(frame_action, text="  更改状态:").pack(side="left")
combo_status = ttk.Combobox(frame_action, values=["已申请", "已寄出", "测试中", "测试通过", "测试未通过"], width=10, state="readonly")
combo_status.set("已寄出")
combo_status.pack(side="left", padx=5)
btn_update = tk.Button(frame_action, text="更新状态", command=update_status, bg="#2196F3", fg="white", width=8)
btn_update.pack(side="left", padx=5)

btn_export = tk.Button(frame_action, text="导出 Excel 报表", command=export_to_excel_gui, bg="#FF9800", fg="white", width=14, font=("Helvetica", 9, "bold"))
btn_export.pack(side="right", padx=5)

frame_table = tk.Frame(root, padx=15, pady=10)
frame_table.pack(fill="both", expand=True)

columns = ("id", "company", "contact", "phone", "model", "qty", "status", "date", "notes", "photo")
tree = ttk.Treeview(frame_table, columns=columns, show="headings")

tree.heading("id", text="编号")
tree.heading("company", text="客户公司")
tree.heading("contact", text="联系人")
tree.heading("phone", text="联系电话")
# 【通用化修改】底部的表格列头重命名
tree.heading("model", text="助剂型号") 
tree.heading("qty", text="数量")
tree.heading("status", text="状态")
tree.heading("date", text="录入时间")
tree.heading("notes", text="备注/反馈")
tree.heading("photo", text="照片路径")

tree.column("id", width=40, anchor="center")
tree.column("company", width=130)
tree.column("contact", width=65)
tree.column("phone", width=95)
tree.column("model", width=95)
tree.column("qty", width=45, anchor="center")
tree.column("status", width=65, anchor="center")
tree.column("date", width=110, anchor="center")
tree.column("notes", width=150)
tree.column("photo", width=140)

tree.pack(fill="both", expand=True)
tree.bind("<Double-1>", on_tree_double_click)

# ==========================================
# 第五步：启动入口
# ==========================================
if __name__ == "__main__":
    try:
        current_dir = os.path.dirname(os.path.abspath(__file__))
        if current_dir:
            os.chdir(current_dir)
        
        init_db()
        show_all_records()
        root.mainloop()
    except Exception as e:
        import tkinter.messagebox as msg
        root_err = tk.Tk()
        root_err.withdraw()
        msg.showerror("启动错误", f"程序运行发生错误:\n{str(e)}")