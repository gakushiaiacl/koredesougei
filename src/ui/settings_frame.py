import tkinter as tk
from tkinter import ttk, messagebox

class SettingsFrame(ttk.Frame):
    """設定画面"""
    
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.create_widgets()
        self.load_settings()
    
    def create_widgets(self):
        """ウィジェットの作成"""
        settings_frame = ttk.LabelFrame(self, text="システム設定", padding=20)
        settings_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 施設の住所
        ttk.Label(settings_frame, text="施設の住所:", font=("", 12)).grid(row=0, column=0, sticky=tk.W, pady=10)
        self.facility_address_var = tk.StringVar()
        ttk.Entry(settings_frame, textvariable=self.facility_address_var, width=50).grid(row=0, column=1, sticky=tk.W, pady=10)
        
        # Google Maps APIキー
        ttk.Label(settings_frame, text="Google Maps APIキー:", font=("", 12)).grid(row=1, column=0, sticky=tk.W, pady=10)
        self.api_key_var = tk.StringVar()
        ttk.Entry(settings_frame, textvariable=self.api_key_var, width=50).grid(row=1, column=1, sticky=tk.W, pady=10)
        
        # APIキーの説明
        api_key_info = (
            "Google Maps APIキーは距離計算、地図表示、ルート表示に使用されます。\n"
            "APIキーがない場合は、ダミーデータを使用して計算しますが地図表示はエラーになります。\n"
            "正確な計算とマップ表示のためには、Google Cloud Platformで以下のAPIを有効にしたAPIキーを取得してください：\n"
            "・Google Maps JavaScript API（地図表示）\n"
            "・Google Maps Directions API（ルート表示）\n"
            "・Google Maps Distance Matrix API（距離/時間計算）\n\n"
            "【APIキー取得方法】\n"
            "1. Google Cloud Platform Consoleにアクセス\n"
            "2. プロジェクトを作成\n"
            "3. 上記の各APIを有効化\n"
            "4. 認証情報からAPIキーを発行"
        )
        ttk.Label(settings_frame, text=api_key_info, foreground="gray", justify=tk.LEFT, wraplength=400).grid(row=2, column=1, sticky=tk.W, pady=5)
        
        # 勤務曜日設定
        ttk.Label(settings_frame, text="勤務曜日:", font=("", 12)).grid(row=3, column=0, sticky=tk.W, pady=10)
        workday_frame = ttk.Frame(settings_frame)
        workday_frame.grid(row=3, column=1, sticky=tk.W, pady=10)
        
        self.workday_vars = {}
        weekdays = ["月", "火", "水", "木", "金", "土", "日"]
        
        for i, day in enumerate(weekdays):
            var = tk.BooleanVar()
            self.workday_vars[day] = var
            ttk.Checkbutton(workday_frame, text=day, variable=var).grid(row=0, column=i, padx=5)
        
        # 保存ボタン
        ttk.Button(settings_frame, text="設定を保存", command=self.save_settings).grid(row=4, column=0, columnspan=2, pady=20)
        
        # ヘルプ情報
        help_frame = ttk.LabelFrame(self, text="ヘルプ", padding=20)
        help_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        help_text = (
            "【使い方】\n"
            "1. 「職員」タブで、送迎に関わる職員を登録してください。運転可能な職員を必ず指定してください。\n"
            "2. 「利用者」タブで、送迎が必要な利用者とその住所を登録してください。\n"
            "3. 「車両」タブで、使用する車両と乗車可能人数を登録してください。\n"
            "4. このタブで施設の住所を必ず入力してください。\n"
            "5. 「送迎スケジュール」タブで「送迎ルート最適化」ボタンを押すと、最適な送迎ルートが計算されます。\n"
            "6. 地図表示ボタンでルートを地図で確認できます（APIキーが必要）。\n"
            "7. ExcelとChatGPTチェック用テキストとしてエクスポートできます。\n\n"
            "※データはアプリケーションを終了する際に自動的に保存されます。"
        )
        
        ttk.Label(help_frame, text=help_text, justify=tk.LEFT, wraplength=600).pack(fill=tk.BOTH)
    
    def load_settings(self):
        """設定の読み込み"""
        # 施設の住所
        self.facility_address_var.set(self.app.settings.get("facility_address", ""))
        
        # APIキー
        self.api_key_var.set(self.app.settings.get("api_key", ""))
        
        # 勤務曜日
        workdays = self.app.settings.get("workdays", ["月", "火", "水", "木", "金", "土"])
        for day, var in self.workday_vars.items():
            var.set(day in workdays)
    
    def save_settings(self):
        """設定の保存"""
        facility_address = self.facility_address_var.get().strip()
        
        if not facility_address:
            messagebox.showerror("エラー", "施設の住所を入力してください。")
            return
        
        # 設定を更新
        self.app.settings["facility_address"] = facility_address
        self.app.settings["api_key"] = self.api_key_var.get().strip()
        self.app.settings["workdays"] = [day for day, var in self.workday_vars.items() if var.get()]
        
        # 設定を保存
        self.app.save_all_data()
        
        messagebox.showinfo("成功", "設定が保存されました。") 