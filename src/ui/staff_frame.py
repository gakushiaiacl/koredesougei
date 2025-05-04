import tkinter as tk
from tkinter import ttk, messagebox
import uuid
from models import Staff

class StaffFrame(ttk.Frame):
    """職員管理画面"""
    
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.create_widgets()
        self.load_staff_list()
    
    def create_widgets(self):
        """ウィジェットの作成"""
        # 左側：入力フォーム
        input_frame = ttk.LabelFrame(self, text="職員情報入力", padding=10)
        input_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 名前の入力
        ttk.Label(input_frame, text="名前:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.name_var = tk.StringVar()
        self.name_entry = ttk.Entry(input_frame, textvariable=self.name_var, width=30)
        self.name_entry.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # 運転可能かどうか
        ttk.Label(input_frame, text="運転可能:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.can_drive_var = tk.BooleanVar()
        ttk.Checkbutton(input_frame, variable=self.can_drive_var).grid(row=1, column=1, sticky=tk.W, pady=5)
        
        # 勤務曜日
        ttk.Label(input_frame, text="勤務曜日:").grid(row=2, column=0, sticky=tk.W, pady=5)
        workday_frame = ttk.Frame(input_frame)
        workday_frame.grid(row=2, column=1, sticky=tk.W, pady=5)
        
        self.workday_vars = {}
        weekdays = ["月", "火", "水", "木", "金", "土"]
        
        for i, day in enumerate(weekdays):
            var = tk.BooleanVar()
            self.workday_vars[day] = var
            ttk.Checkbutton(workday_frame, text=day, variable=var).grid(row=0, column=i, padx=5)
        
        # ボタンフレーム
        button_frame = ttk.Frame(input_frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="追加", command=self.add_staff).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="更新", command=self.update_staff).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="削除", command=self.delete_staff).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="クリア", command=self.clear_form).pack(side=tk.LEFT, padx=5)
        
        # Excelインポート/エクスポートボタン
        excel_frame = ttk.Frame(input_frame)
        excel_frame.grid(row=4, column=0, columnspan=2, pady=10)
        
        ttk.Button(excel_frame, text="Excelエクスポート", command=self.export_to_excel).pack(side=tk.LEFT, padx=5)
        ttk.Button(excel_frame, text="Excelインポート", command=self.import_from_excel).pack(side=tk.LEFT, padx=5)
        
        # 右側：職員リスト
        list_frame = ttk.LabelFrame(self, text="職員リスト", padding=10)
        list_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # ツリービュー
        columns = ("id", "name", "can_drive", "workdays")
        self.staff_tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="browse")
        
        # 列の設定
        self.staff_tree.heading("id", text="ID")
        self.staff_tree.heading("name", text="名前")
        self.staff_tree.heading("can_drive", text="運転可能")
        self.staff_tree.heading("workdays", text="勤務曜日")
        
        self.staff_tree.column("id", width=50)
        self.staff_tree.column("name", width=100)
        self.staff_tree.column("can_drive", width=80)
        self.staff_tree.column("workdays", width=250)
        
        self.staff_tree.pack(fill=tk.BOTH, expand=True)
        
        # スクロールバー
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.staff_tree.yview)
        self.staff_tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 選択時のイベント
        self.staff_tree.bind("<<TreeviewSelect>>", self.on_staff_select)
    
    def load_staff_list(self):
        """職員リストの読み込み"""
        # ツリービューの中身をクリア
        for item in self.staff_tree.get_children():
            self.staff_tree.delete(item)
        
        # 職員リストをツリービューに追加
        for staff in self.app.staff_list:
            workdays_str = ", ".join(staff.workdays)
            can_drive_str = "はい" if staff.can_drive else "いいえ"
            
            self.staff_tree.insert("", tk.END, values=(
                staff.id,
                staff.name,
                can_drive_str,
                workdays_str
            ))
    
    def on_staff_select(self, event):
        """職員選択時の処理"""
        selected_items = self.staff_tree.selection()
        if not selected_items:
            return
        
        # 選択された職員の情報を取得
        item = selected_items[0]
        staff_id = self.staff_tree.item(item, "values")[0]
        
        # 職員IDに一致する職員を検索
        for staff in self.app.staff_list:
            if str(staff.id) == str(staff_id):
                # フォームに値を設定
                self.name_var.set(staff.name)
                self.can_drive_var.set(staff.can_drive)
                
                # 勤務曜日チェックボックスの設定
                for day, var in self.workday_vars.items():
                    var.set(day in staff.workdays)
                
                break
    
    def get_selected_workdays(self):
        """選択された勤務曜日を取得"""
        return [day for day, var in self.workday_vars.items() if var.get()]
    
    def add_staff(self):
        """職員の追加"""
        name = self.name_var.get().strip()
        
        if not name:
            messagebox.showerror("エラー", "名前を入力してください。")
            return
        
        # 新しい職員オブジェクトを作成
        new_staff = Staff(
            id=str(uuid.uuid4()),
            name=name,
            can_drive=self.can_drive_var.get(),
            workdays=self.get_selected_workdays()
        )
        
        # 職員リストに追加
        self.app.staff_list.append(new_staff)
        
        # リストを更新
        self.load_staff_list()
        
        # フォームをクリア
        self.clear_form()
        
        messagebox.showinfo("成功", f"職員「{name}」が追加されました。")
    
    def update_staff(self):
        """職員情報の更新"""
        selected_items = self.staff_tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "更新する職員を選択してください。")
            return
        
        # 選択された職員のID
        item = selected_items[0]
        staff_id = self.staff_tree.item(item, "values")[0]
        
        # 入力値の取得
        name = self.name_var.get().strip()
        
        if not name:
            messagebox.showerror("エラー", "名前を入力してください。")
            return
        
        # 職員の更新
        for i, staff in enumerate(self.app.staff_list):
            if str(staff.id) == str(staff_id):
                # 職員情報を更新
                staff.name = name
                staff.can_drive = self.can_drive_var.get()
                staff.workdays = self.get_selected_workdays()
                
                # リストを更新
                self.load_staff_list()
                
                messagebox.showinfo("成功", f"職員「{name}」の情報が更新されました。")
                break
    
    def delete_staff(self):
        """職員の削除"""
        selected_items = self.staff_tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "削除する職員を選択してください。")
            return
        
        # 選択された職員のID
        item = selected_items[0]
        staff_id = self.staff_tree.item(item, "values")[0]
        staff_name = self.staff_tree.item(item, "values")[1]
        
        # 確認ダイアログ
        if not messagebox.askyesno("確認", f"職員「{staff_name}」を削除しますか？"):
            return
        
        # 職員の削除
        for i, staff in enumerate(self.app.staff_list):
            if str(staff.id) == str(staff_id):
                del self.app.staff_list[i]
                break
        
        # リストを更新
        self.load_staff_list()
        
        # フォームをクリア
        self.clear_form()
        
        messagebox.showinfo("成功", f"職員「{staff_name}」が削除されました。")
    
    def clear_form(self):
        """入力フォームのクリア"""
        self.name_var.set("")
        self.can_drive_var.set(False)
        
        for var in self.workday_vars.values():
            var.set(False)
        
        # 選択を解除
        for item in self.staff_tree.selection():
            self.staff_tree.selection_remove(item)
            
    def export_to_excel(self):
        """職員データをExcelにエクスポートする"""
        filepath = self.app.export_manager.export_data_to_excel("staff")
        if filepath:
            messagebox.showinfo("エクスポート完了", f"職員データが以下の場所にエクスポートされました：\n{filepath}")
    
    def import_from_excel(self):
        """Excelファイルからデータをインポートする"""
        if messagebox.askyesno("確認", "Excelからデータをインポートすると、現在のデータが上書きされます。続行しますか？"):
            self.app.export_manager.import_data_from_excel("staff") 