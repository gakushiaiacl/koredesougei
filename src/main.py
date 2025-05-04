import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
import uuid
from models import Staff, User, Vehicle
from ui.staff_frame import StaffFrame
from ui.user_frame import UserFrame
from ui.vehicle_frame import VehicleFrame
from ui.schedule_frame import ScheduleFrame
from ui.settings_frame import SettingsFrame
from ui.export_manager import ExportManager

class TransportApp:
    """送迎スケジューリングアプリケーション"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("これで送迎")
        self.root.geometry("1000x700")
        
        # データを保存するディレクトリ
        self.data_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data")
        os.makedirs(self.data_dir, exist_ok=True)
        
        # データファイルのパス
        self.staff_file = os.path.join(self.data_dir, "staff.json")
        self.users_file = os.path.join(self.data_dir, "users.json")
        self.vehicles_file = os.path.join(self.data_dir, "vehicles.json")
        self.settings_file = os.path.join(self.data_dir, "settings.json")
        
        # データの初期化
        self.staff_list = []
        self.user_list = []
        self.vehicle_list = []
        self.settings = {
            "workdays": ["月", "火", "水", "木", "金", "土"],
            "api_key": "",
            "facility_address": ""
        }
        
        # 保存されたデータを読み込む
        self.load_data()
        
        # エクスポートマネージャーの初期化
        self.export_manager = ExportManager(self)
        
        # タブコントロール
        self.tab_control = ttk.Notebook(root)
        
        # 各タブのフレーム
        self.staff_frame = StaffFrame(self.tab_control, self)
        self.user_frame = UserFrame(self.tab_control, self)
        self.vehicle_frame = VehicleFrame(self.tab_control, self)
        self.schedule_frame = ScheduleFrame(self.tab_control, self)
        self.settings_frame = SettingsFrame(self.tab_control, self)
        
        # タブの追加
        self.tab_control.add(self.staff_frame, text="職員")
        self.tab_control.add(self.user_frame, text="利用者")
        self.tab_control.add(self.vehicle_frame, text="車両")
        self.tab_control.add(self.schedule_frame, text="送迎スケジュール")
        self.tab_control.add(self.settings_frame, text="設定")
        
        self.tab_control.pack(expand=1, fill="both")
        
        # 保存ボタン
        save_button = ttk.Button(root, text="すべてのデータを保存", command=self.save_all_data)
        save_button.pack(pady=10)
        
        # 終了時にデータを保存
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def load_data(self):
        """保存されたデータを読み込む"""
        # 職員データ
        if os.path.exists(self.staff_file):
            try:
                with open(self.staff_file, 'r', encoding='utf-8') as f:
                    staff_data = json.load(f)
                    self.staff_list = [Staff(**data) for data in staff_data]
            except Exception as e:
                messagebox.showwarning("警告", f"職員データの読み込みに失敗しました: {str(e)}")
        
        # 利用者データ
        if os.path.exists(self.users_file):
            try:
                with open(self.users_file, 'r', encoding='utf-8') as f:
                    user_data = json.load(f)
                    self.user_list = [User(**data) for data in user_data]
            except Exception as e:
                messagebox.showwarning("警告", f"利用者データの読み込みに失敗しました: {str(e)}")
        
        # 車両データ
        if os.path.exists(self.vehicles_file):
            try:
                with open(self.vehicles_file, 'r', encoding='utf-8') as f:
                    vehicle_data = json.load(f)
                    self.vehicle_list = [Vehicle(**data) for data in vehicle_data]
            except Exception as e:
                messagebox.showwarning("警告", f"車両データの読み込みに失敗しました: {str(e)}")
        
        # 設定データ
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    self.settings = json.load(f)
            except Exception as e:
                messagebox.showwarning("警告", f"設定データの読み込みに失敗しました: {str(e)}")
    
    def save_all_data(self):
        """すべてのデータを保存"""
        try:
            # 職員データ
            with open(self.staff_file, 'w', encoding='utf-8') as f:
                json.dump([staff.to_dict() for staff in self.staff_list], f, ensure_ascii=False, indent=2)
            
            # 利用者データ
            with open(self.users_file, 'w', encoding='utf-8') as f:
                json.dump([user.to_dict() for user in self.user_list], f, ensure_ascii=False, indent=2)
            
            # 車両データ
            with open(self.vehicles_file, 'w', encoding='utf-8') as f:
                json.dump([vehicle.to_dict() for vehicle in self.vehicle_list], f, ensure_ascii=False, indent=2)
            
            # 設定データ
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=2)
            
            messagebox.showinfo("保存完了", "すべてのデータが保存されました。")
        except Exception as e:
            messagebox.showerror("エラー", f"データの保存中にエラーが発生しました: {str(e)}")
    
    def on_closing(self):
        """アプリケーション終了時の処理"""
        if messagebox.askokcancel("終了確認", "変更を保存して終了しますか？"):
            self.save_all_data()
            self.root.destroy()

def create_sample_data(app):
    """サンプルデータの作成"""
    # すでにデータがある場合は作成しない
    if app.staff_list or app.user_list or app.vehicle_list:
        return
    
    # 職員データ
    app.staff_list = [
        Staff(id=str(uuid.uuid4()), name="山田太郎", can_drive=True, workdays=["月", "火", "水", "木", "金"]),
        Staff(id=str(uuid.uuid4()), name="佐藤花子", can_drive=True, workdays=["月", "水", "金"]),
        Staff(id=str(uuid.uuid4()), name="鈴木一郎", can_drive=True, workdays=["火", "木", "土"]),
        Staff(id=str(uuid.uuid4()), name="高橋次郎", can_drive=False, workdays=["月", "火", "水", "木", "金"]),
        Staff(id=str(uuid.uuid4()), name="田中三郎", can_drive=False, workdays=["月", "火", "水", "木", "金", "土"]),
    ]
    
    # 車両データ
    app.vehicle_list = [
        Vehicle(id=str(uuid.uuid4()), name="ハイエース", capacity=3),
        Vehicle(id=str(uuid.uuid4()), name="キャラバン", capacity=4),
        Vehicle(id=str(uuid.uuid4()), name="エブリイ", capacity=2),
    ]
    
    # 利用者データ
    app.user_list = [
        User(id=str(uuid.uuid4()), name="佐々木幸子", address="東京都新宿区西新宿1-1-1", 
             pickup_time_morning="8:00", dropoff_time_morning="9:00", 
             pickup_time_evening="16:00", dropoff_time_evening="17:00", 
             attendance_days=["月", "水", "金"]),
        User(id=str(uuid.uuid4()), name="伊藤健太", address="東京都渋谷区道玄坂1-1-1", 
             pickup_time_morning="8:15", dropoff_time_morning="9:15", 
             pickup_time_evening="16:15", dropoff_time_evening="17:15", 
             attendance_days=["月", "火", "木"]),
        User(id=str(uuid.uuid4()), name="小林京子", address="東京都中野区中野1-1-1", 
             pickup_time_morning="8:30", dropoff_time_morning="9:30", 
             pickup_time_evening="16:30", dropoff_time_evening="17:30", 
             attendance_days=["火", "木", "土"]),
        User(id=str(uuid.uuid4()), name="吉田智也", address="東京都杉並区荻窪1-1-1", 
             pickup_time_morning="8:45", dropoff_time_morning="9:45", 
             pickup_time_evening="16:45", dropoff_time_evening="17:45", 
             attendance_days=["月", "水", "金"]),
        User(id=str(uuid.uuid4()), name="加藤裕子", address="東京都練馬区光が丘1-1-1", 
             pickup_time_morning="8:10", dropoff_time_morning="9:10", 
             pickup_time_evening="16:10", dropoff_time_evening="17:10", 
             attendance_days=["月", "火", "水", "木", "金"]),
    ]
    
    # 設定データ
    app.settings = {
        "workdays": ["月", "火", "水", "木", "金", "土"],
        "api_key": "",
        "facility_address": "東京都豊島区東池袋1-1-1"
    }
    
    # データを保存
    app.save_all_data()
    
    messagebox.showinfo("サンプルデータ", "サンプルデータが作成されました。")

if __name__ == "__main__":
    root = tk.Tk()
    app = TransportApp(root)
    
    # コマンドライン引数でサンプルデータ作成を指定可能
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "--sample":
        create_sample_data(app)
    
    root.mainloop() 