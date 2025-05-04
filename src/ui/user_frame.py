import tkinter as tk
from tkinter import ttk, messagebox
import uuid
from models import User

class UserFrame(ttk.Frame):
    """利用者管理画面"""
    
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.create_widgets()
        self.load_user_list()
    
    def create_widgets(self):
        """ウィジェットの作成"""
        # 左側：入力フォーム
        input_frame = ttk.LabelFrame(self, text="利用者情報入力", padding=10)
        input_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 名前の入力
        ttk.Label(input_frame, text="名前:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.name_var = tk.StringVar()
        self.name_entry = ttk.Entry(input_frame, textvariable=self.name_var, width=30)
        self.name_entry.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # 住所の入力
        ttk.Label(input_frame, text="住所:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.address_var = tk.StringVar()
        self.address_entry = ttk.Entry(input_frame, textvariable=self.address_var, width=30)
        self.address_entry.grid(row=1, column=1, sticky=tk.W, pady=5)
        
        # 朝の送迎時間
        ttk.Label(input_frame, text="朝の送迎時間:").grid(row=2, column=0, sticky=tk.W, pady=5)
        time_frame_morning = ttk.Frame(input_frame)
        time_frame_morning.grid(row=2, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(time_frame_morning, text="迎え:").pack(side=tk.LEFT)
        self.pickup_time_morning_var = tk.StringVar()
        ttk.Entry(time_frame_morning, textvariable=self.pickup_time_morning_var, width=10).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(time_frame_morning, text="送り:").pack(side=tk.LEFT, padx=(10, 0))
        self.dropoff_time_morning_var = tk.StringVar()
        ttk.Entry(time_frame_morning, textvariable=self.dropoff_time_morning_var, width=10).pack(side=tk.LEFT, padx=5)
        
        # 夕方の送迎時間
        ttk.Label(input_frame, text="夕方の送迎時間:").grid(row=3, column=0, sticky=tk.W, pady=5)
        time_frame_evening = ttk.Frame(input_frame)
        time_frame_evening.grid(row=3, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(time_frame_evening, text="迎え:").pack(side=tk.LEFT)
        self.pickup_time_evening_var = tk.StringVar()
        ttk.Entry(time_frame_evening, textvariable=self.pickup_time_evening_var, width=10).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(time_frame_evening, text="送り:").pack(side=tk.LEFT, padx=(10, 0))
        self.dropoff_time_evening_var = tk.StringVar()
        ttk.Entry(time_frame_evening, textvariable=self.dropoff_time_evening_var, width=10).pack(side=tk.LEFT, padx=5)
        
        # 制約条件
        ttk.Label(input_frame, text="制約条件:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.constraints_var = tk.StringVar()
        self.constraints_entry = ttk.Entry(input_frame, textvariable=self.constraints_var, width=30)
        self.constraints_entry.grid(row=4, column=1, sticky=tk.W, pady=5)
        
        # 通所曜日
        ttk.Label(input_frame, text="通所曜日:").grid(row=5, column=0, sticky=tk.W, pady=5)
        attendance_frame = ttk.Frame(input_frame)
        attendance_frame.grid(row=5, column=1, sticky=tk.W, pady=5)
        
        self.attendance_vars = {}
        weekdays = ["月", "火", "水", "木", "金", "土"]
        
        for i, day in enumerate(weekdays):
            var = tk.BooleanVar()
            self.attendance_vars[day] = var
            ttk.Checkbutton(attendance_frame, text=day, variable=var).grid(row=0, column=i, padx=5)
        
        # ボタンフレーム
        button_frame = ttk.Frame(input_frame)
        button_frame.grid(row=6, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="追加", command=self.add_user).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="更新", command=self.update_user).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="削除", command=self.delete_user).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="クリア", command=self.clear_form).pack(side=tk.LEFT, padx=5)

        # Excelインポート/エクスポートボタン
        excel_frame = ttk.Frame(input_frame)
        excel_frame.grid(row=7, column=0, columnspan=2, pady=10)
        
        ttk.Button(excel_frame, text="Excelエクスポート", command=self.export_to_excel).pack(side=tk.LEFT, padx=5)
        ttk.Button(excel_frame, text="Excelインポート", command=self.import_from_excel).pack(side=tk.LEFT, padx=5)
        
        # 右側：利用者リスト
        list_frame = ttk.LabelFrame(self, text="利用者リスト", padding=10)
        list_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # ツリービュー
        columns = ("id", "name", "address", "attendance_days")
        self.user_tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="browse")
        
        # 列の設定
        self.user_tree.heading("id", text="ID")
        self.user_tree.heading("name", text="名前")
        self.user_tree.heading("address", text="住所")
        self.user_tree.heading("attendance_days", text="通所曜日")
        
        self.user_tree.column("id", width=50)
        self.user_tree.column("name", width=100)
        self.user_tree.column("address", width=150)
        self.user_tree.column("attendance_days", width=150)
        
        self.user_tree.pack(fill=tk.BOTH, expand=True)
        
        # スクロールバー
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.user_tree.yview)
        self.user_tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 選択時のイベント
        self.user_tree.bind("<<TreeviewSelect>>", self.on_user_select)
    
    def load_user_list(self):
        """利用者リストの読み込み"""
        # ツリービューの中身をクリア
        for item in self.user_tree.get_children():
            self.user_tree.delete(item)
        
        # 利用者リストをツリービューに追加
        for user in self.app.user_list:
            attendance_days_str = ", ".join(user.attendance_days)
            
            self.user_tree.insert("", tk.END, values=(
                user.id,
                user.name,
                user.address,
                attendance_days_str
            ))
    
    def on_user_select(self, event):
        """利用者選択時の処理"""
        selected_items = self.user_tree.selection()
        if not selected_items:
            return
        
        # 選択された利用者の情報を取得
        item = selected_items[0]
        user_id = self.user_tree.item(item, "values")[0]
        
        # 利用者IDに一致する利用者を検索
        for user in self.app.user_list:
            if str(user.id) == str(user_id):
                # フォームに値を設定
                self.name_var.set(user.name)
                self.address_var.set(user.address)
                self.pickup_time_morning_var.set(user.pickup_time_morning)
                self.dropoff_time_morning_var.set(user.dropoff_time_morning)
                self.pickup_time_evening_var.set(user.pickup_time_evening)
                self.dropoff_time_evening_var.set(user.dropoff_time_evening)
                self.constraints_var.set(user.constraints)
                
                # 通所曜日チェックボックスの設定
                for day, var in self.attendance_vars.items():
                    var.set(day in user.attendance_days)
                
                break
    
    def get_selected_attendance_days(self):
        """選択された通所曜日を取得"""
        return [day for day, var in self.attendance_vars.items() if var.get()]
    
    def add_user(self):
        """利用者の追加"""
        name = self.name_var.get().strip()
        address = self.address_var.get().strip()
        
        if not name:
            messagebox.showerror("エラー", "名前を入力してください。")
            return
        
        if not address:
            messagebox.showerror("エラー", "住所を入力してください。")
            return
        
        # 新しい利用者オブジェクトを作成
        new_user = User(
            id=str(uuid.uuid4()),
            name=name,
            address=address,
            pickup_time_morning=self.pickup_time_morning_var.get(),
            dropoff_time_morning=self.dropoff_time_morning_var.get(),
            pickup_time_evening=self.pickup_time_evening_var.get(),
            dropoff_time_evening=self.dropoff_time_evening_var.get(),
            constraints=self.constraints_var.get(),
            attendance_days=self.get_selected_attendance_days()
        )
        
        # 利用者リストに追加
        self.app.user_list.append(new_user)
        
        # リストを更新
        self.load_user_list()
        
        # フォームをクリア
        self.clear_form()
        
        messagebox.showinfo("成功", f"利用者「{name}」が追加されました。")
    
    def update_user(self):
        """利用者情報の更新"""
        selected_items = self.user_tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "更新する利用者を選択してください。")
            return
        
        # 選択された利用者のID
        item = selected_items[0]
        user_id = self.user_tree.item(item, "values")[0]
        
        # 入力値の取得
        name = self.name_var.get().strip()
        address = self.address_var.get().strip()
        
        if not name:
            messagebox.showerror("エラー", "名前を入力してください。")
            return
        
        if not address:
            messagebox.showerror("エラー", "住所を入力してください。")
            return
        
        # 利用者の更新
        for i, user in enumerate(self.app.user_list):
            if str(user.id) == str(user_id):
                # 利用者情報を更新
                user.name = name
                user.address = address
                user.pickup_time_morning = self.pickup_time_morning_var.get()
                user.dropoff_time_morning = self.dropoff_time_morning_var.get()
                user.pickup_time_evening = self.pickup_time_evening_var.get()
                user.dropoff_time_evening = self.dropoff_time_evening_var.get()
                user.constraints = self.constraints_var.get()
                user.attendance_days = self.get_selected_attendance_days()
                
                # リストを更新
                self.load_user_list()
                
                messagebox.showinfo("成功", f"利用者「{name}」の情報が更新されました。")
                break
    
    def delete_user(self):
        """利用者の削除"""
        selected_items = self.user_tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "削除する利用者を選択してください。")
            return
        
        # 選択された利用者のID
        item = selected_items[0]
        user_id = self.user_tree.item(item, "values")[0]
        user_name = self.user_tree.item(item, "values")[1]
        
        # 確認ダイアログ
        if not messagebox.askyesno("確認", f"利用者「{user_name}」を削除しますか？"):
            return
        
        # 利用者の削除
        for i, user in enumerate(self.app.user_list):
            if str(user.id) == str(user_id):
                del self.app.user_list[i]
                break
        
        # リストを更新
        self.load_user_list()
        
        # フォームをクリア
        self.clear_form()
        
        messagebox.showinfo("成功", f"利用者「{user_name}」が削除されました。")
    
    def clear_form(self):
        """入力フォームのクリア"""
        self.name_var.set("")
        self.address_var.set("")
        self.pickup_time_morning_var.set("")
        self.dropoff_time_morning_var.set("")
        self.pickup_time_evening_var.set("")
        self.dropoff_time_evening_var.set("")
        self.constraints_var.set("")
        
        for var in self.attendance_vars.values():
            var.set(False)
        
        # 選択を解除
        for item in self.user_tree.selection():
            self.user_tree.selection_remove(item)
            
    def export_to_excel(self):
        """利用者データをExcelにエクスポートする"""
        filepath = self.app.export_manager.export_data_to_excel("user")
        if filepath:
            messagebox.showinfo("エクスポート完了", f"利用者データが以下の場所にエクスポートされました：\n{filepath}")
    
    def import_from_excel(self):
        """Excelファイルからデータをインポートする"""
        if messagebox.askyesno("確認", "Excelからデータをインポートすると、現在のデータが上書きされます。続行しますか？"):
            self.app.export_manager.import_data_from_excel("user") 