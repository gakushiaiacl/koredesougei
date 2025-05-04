import tkinter as tk
from tkinter import ttk, messagebox
import uuid
from models import Vehicle

class VehicleFrame(ttk.Frame):
    """車両管理画面"""
    
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.create_widgets()
        self.load_vehicle_list()
    
    def create_widgets(self):
        """ウィジェットの作成"""
        # 左側：入力フォーム
        input_frame = ttk.LabelFrame(self, text="車両情報入力", padding=10)
        input_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 名前の入力
        ttk.Label(input_frame, text="車両名:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.name_var = tk.StringVar()
        self.name_entry = ttk.Entry(input_frame, textvariable=self.name_var, width=30)
        self.name_entry.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # 乗車可能人数
        ttk.Label(input_frame, text="乗車可能人数:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.capacity_var = tk.StringVar()
        capacity_spinner = ttk.Spinbox(input_frame, from_=1, to=10, textvariable=self.capacity_var, width=5)
        capacity_spinner.grid(row=1, column=1, sticky=tk.W, pady=5)
        self.capacity_var.set("3")  # デフォルト値
        
        # ボタンフレーム
        button_frame = ttk.Frame(input_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="追加", command=self.add_vehicle).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="更新", command=self.update_vehicle).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="削除", command=self.delete_vehicle).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="クリア", command=self.clear_form).pack(side=tk.LEFT, padx=5)
        
        # Excelインポート/エクスポートボタン
        excel_frame = ttk.Frame(input_frame)
        excel_frame.grid(row=3, column=0, columnspan=2, pady=10)
        
        ttk.Button(excel_frame, text="Excelエクスポート", command=self.export_to_excel).pack(side=tk.LEFT, padx=5)
        ttk.Button(excel_frame, text="Excelインポート", command=self.import_from_excel).pack(side=tk.LEFT, padx=5)
        
        # 右側：車両リスト
        list_frame = ttk.LabelFrame(self, text="車両リスト", padding=10)
        list_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # ツリービュー
        columns = ("id", "name", "capacity")
        self.vehicle_tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="browse")
        
        # 列の設定
        self.vehicle_tree.heading("id", text="ID")
        self.vehicle_tree.heading("name", text="車両名")
        self.vehicle_tree.heading("capacity", text="乗車可能人数")
        
        self.vehicle_tree.column("id", width=50)
        self.vehicle_tree.column("name", width=150)
        self.vehicle_tree.column("capacity", width=100)
        
        self.vehicle_tree.pack(fill=tk.BOTH, expand=True)
        
        # スクロールバー
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.vehicle_tree.yview)
        self.vehicle_tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 選択時のイベント
        self.vehicle_tree.bind("<<TreeviewSelect>>", self.on_vehicle_select)
    
    def load_vehicle_list(self):
        """車両リストの読み込み"""
        # ツリービューの中身をクリア
        for item in self.vehicle_tree.get_children():
            self.vehicle_tree.delete(item)
        
        # 車両リストをツリービューに追加
        for vehicle in self.app.vehicle_list:
            self.vehicle_tree.insert("", tk.END, values=(
                vehicle.id,
                vehicle.name,
                vehicle.capacity
            ))
    
    def on_vehicle_select(self, event):
        """車両選択時の処理"""
        selected_items = self.vehicle_tree.selection()
        if not selected_items:
            return
        
        # 選択された車両の情報を取得
        item = selected_items[0]
        vehicle_id = self.vehicle_tree.item(item, "values")[0]
        
        # 車両IDに一致する車両を検索
        for vehicle in self.app.vehicle_list:
            if str(vehicle.id) == str(vehicle_id):
                # フォームに値を設定
                self.name_var.set(vehicle.name)
                self.capacity_var.set(vehicle.capacity)
                break
    
    def add_vehicle(self):
        """車両の追加"""
        name = self.name_var.get().strip()
        
        if not name:
            messagebox.showerror("エラー", "車両名を入力してください。")
            return
        
        try:
            capacity = int(self.capacity_var.get())
            if capacity <= 0:
                raise ValueError()
        except ValueError:
            messagebox.showerror("エラー", "乗車可能人数は1以上の数値を入力してください。")
            return
        
        # 新しい車両オブジェクトを作成
        new_vehicle = Vehicle(
            id=str(uuid.uuid4()),
            name=name,
            capacity=capacity
        )
        
        # 車両リストに追加
        self.app.vehicle_list.append(new_vehicle)
        
        # リストを更新
        self.load_vehicle_list()
        
        # フォームをクリア
        self.clear_form()
        
        messagebox.showinfo("成功", f"車両「{name}」が追加されました。")
    
    def update_vehicle(self):
        """車両情報の更新"""
        selected_items = self.vehicle_tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "更新する車両を選択してください。")
            return
        
        # 選択された車両のID
        item = selected_items[0]
        vehicle_id = self.vehicle_tree.item(item, "values")[0]
        
        # 入力値の取得
        name = self.name_var.get().strip()
        
        if not name:
            messagebox.showerror("エラー", "車両名を入力してください。")
            return
        
        try:
            capacity = int(self.capacity_var.get())
            if capacity <= 0:
                raise ValueError()
        except ValueError:
            messagebox.showerror("エラー", "乗車可能人数は1以上の数値を入力してください。")
            return
        
        # 車両の更新
        for i, vehicle in enumerate(self.app.vehicle_list):
            if str(vehicle.id) == str(vehicle_id):
                # 車両情報を更新
                vehicle.name = name
                vehicle.capacity = capacity
                
                # リストを更新
                self.load_vehicle_list()
                
                messagebox.showinfo("成功", f"車両「{name}」の情報が更新されました。")
                break
    
    def delete_vehicle(self):
        """車両の削除"""
        selected_items = self.vehicle_tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "削除する車両を選択してください。")
            return
        
        # 選択された車両のID
        item = selected_items[0]
        vehicle_id = self.vehicle_tree.item(item, "values")[0]
        vehicle_name = self.vehicle_tree.item(item, "values")[1]
        
        # 確認ダイアログ
        if not messagebox.askyesno("確認", f"車両「{vehicle_name}」を削除しますか？"):
            return
        
        # 車両の削除
        for i, vehicle in enumerate(self.app.vehicle_list):
            if str(vehicle.id) == str(vehicle_id):
                del self.app.vehicle_list[i]
                break
        
        # リストを更新
        self.load_vehicle_list()
        
        # フォームをクリア
        self.clear_form()
        
        messagebox.showinfo("成功", f"車両「{vehicle_name}」が削除されました。")
    
    def clear_form(self):
        """入力フォームのクリア"""
        self.name_var.set("")
        self.capacity_var.set("3")  # デフォルト値にリセット
        
        # 選択を解除
        for item in self.vehicle_tree.selection():
            self.vehicle_tree.selection_remove(item)
            
    def export_to_excel(self):
        """車両データをExcelにエクスポートする"""
        filepath = self.app.export_manager.export_data_to_excel("vehicle")
        if filepath:
            messagebox.showinfo("エクスポート完了", f"車両データが以下の場所にエクスポートされました：\n{filepath}")
    
    def import_from_excel(self):
        """Excelファイルからデータをインポートする"""
        if messagebox.askyesno("確認", "Excelからデータをインポートすると、現在のデータが上書きされます。続行しますか？"):
            self.app.export_manager.import_data_from_excel("vehicle") 