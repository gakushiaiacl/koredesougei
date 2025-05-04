import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import datetime
from models import Route
from ui.map_view import MapView
from ui.export_manager import ExportManager

class ScheduleFrame(ttk.Frame):
    """送迎スケジュール画面"""
    
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.routes = []  # 計算されたルート
        self.weekdays = ["月", "火", "水", "木", "金", "土"]
        self.weekday_index = 0  # 現在表示している曜日のインデックス
        self.selected_day = self.weekdays[self.weekday_index]  # 現在選択している曜日
        self.is_morning = True  # 朝か夕方か
        
        # 地図表示用オブジェクト
        self.map_view = None
        
        # エクスポート管理用オブジェクト
        self.export_manager = None
        
        self.create_widgets()
    
    def create_widgets(self):
        """ウィジェットの作成"""
        # トップフレーム：曜日選択と最適化ボタン
        top_frame = ttk.Frame(self, padding=10)
        top_frame.pack(fill=tk.X)
        
        # 曜日選択ボタン
        day_frame = ttk.LabelFrame(top_frame, text="曜日選択", padding=5)
        day_frame.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(day_frame, text="<", command=self.prev_day).pack(side=tk.LEFT, padx=5)
        self.day_label = ttk.Label(day_frame, text=self.selected_day, font=("", 14))
        self.day_label.pack(side=tk.LEFT, padx=10)
        ttk.Button(day_frame, text=">", command=self.next_day).pack(side=tk.LEFT, padx=5)
        
        # 朝/夕方切り替えボタン
        time_frame = ttk.LabelFrame(top_frame, text="時間帯", padding=5)
        time_frame.pack(side=tk.LEFT, padx=5)
        
        self.time_var = tk.StringVar(value="朝")
        ttk.Radiobutton(time_frame, text="朝", variable=self.time_var, value="朝", 
                        command=self.on_time_change).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(time_frame, text="夕方", variable=self.time_var, value="夕方", 
                        command=self.on_time_change).pack(side=tk.LEFT, padx=5)
        
        # 最適化ボタン
        optimize_frame = ttk.Frame(top_frame)
        optimize_frame.pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(optimize_frame, text="送迎ルート最適化", command=self.optimize_routes).pack(padx=5, pady=5)
        
        # 機能ボタンフレーム
        action_frame = ttk.Frame(self, padding=5)
        action_frame.pack(fill=tk.X, padx=5, pady=(0, 5))
        
        # 地図表示ボタン
        ttk.Button(action_frame, text="地図表示", command=self.show_map).pack(side=tk.LEFT, padx=5)
        
        # エクスポートボタンフレーム
        export_frame = ttk.LabelFrame(action_frame, text="エクスポート", padding=5)
        export_frame.pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(export_frame, text="Excel", command=self.export_to_excel).pack(side=tk.LEFT, padx=5)
        ttk.Button(export_frame, text="ChatGPTチェック用", command=self.export_for_chat_gpt).pack(side=tk.LEFT, padx=5)
        
        # 結果表示エリア
        result_frame = ttk.LabelFrame(self, text="送迎スケジュール", padding=10)
        result_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # ツリービュー（車両ごとのタブを作成）
        self.notebook = ttk.Notebook(result_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # デフォルトのタブ（未計算時）
        self.default_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.default_tab, text="スケジュールなし")
        
        ttk.Label(self.default_tab, text="「送迎ルート最適化」ボタンをクリックして、\n送迎ルートを計算してください。", 
                  font=("", 12), justify=tk.CENTER).pack(expand=True)
    
    def prev_day(self):
        """前の曜日に移動"""
        self.weekday_index = (self.weekday_index - 1) % len(self.weekdays)
        self.selected_day = self.weekdays[self.weekday_index]
        self.day_label.config(text=self.selected_day)
        self.update_schedule_display()
    
    def next_day(self):
        """次の曜日に移動"""
        self.weekday_index = (self.weekday_index + 1) % len(self.weekdays)
        self.selected_day = self.weekdays[self.weekday_index]
        self.day_label.config(text=self.selected_day)
        self.update_schedule_display()
    
    def on_time_change(self):
        """時間帯の切り替え"""
        self.is_morning = (self.time_var.get() == "朝")
        self.update_schedule_display()
    
    def optimize_routes(self):
        """送迎ルートの最適化"""
        # 入力データのチェック
        if not self.app.staff_list:
            messagebox.showerror("エラー", "職員が登録されていません。")
            return
            
        if not self.app.user_list:
            messagebox.showerror("エラー", "利用者が登録されていません。")
            return
            
        if not self.app.vehicle_list:
            messagebox.showerror("エラー", "車両が登録されていません。")
            return
        
        # 運転可能な職員のチェック
        drivers = [s for s in self.app.staff_list if s.can_drive]
        if not drivers:
            messagebox.showerror("エラー", "運転可能な職員が登録されていません。")
            return
        
        # 設定のチェック
        if not self.app.settings.get("facility_address"):
            messagebox.showerror("エラー", "施設の住所が設定されていません。設定タブで施設の住所を入力してください。")
            return
        
        # APIキーのチェック（実際には使用しなくてもOK - ダミーデータを使用）
        api_key = self.app.settings.get("api_key") or "dummy_key"
        
        # 進捗ダイアログの表示
        progress = tk.Toplevel(self)
        progress.title("計算中")
        progress.geometry("300x100")
        progress.transient(self)
        progress.grab_set()
        
        ttk.Label(progress, text="送迎ルートを計算中です...", font=("", 12)).pack(pady=10)
        progress_bar = ttk.Progressbar(progress, mode="indeterminate")
        progress_bar.pack(fill=tk.X, padx=20)
        progress_bar.start()
        
        # 更新して表示を確実に
        self.update_idletasks()
        
        try:
            # モジュールのインポートを試みる
            try:
                # 各曜日と時間帯ごとに最適化
                from src.optimizer import TransportOptimizer
                optimizer = TransportOptimizer(api_key, self.app.settings["facility_address"])
            except ImportError as e:
                # インポートエラーの詳細を表示
                messagebox.showerror("インポートエラー", 
                                    f"最適化モジュールのインポートに失敗しました: {str(e)}\n"
                                    f"エラータイプ: {type(e).__name__}\n"
                                    f"モジュール名: {getattr(e, 'name', 'unknown')}")
                raise
            
            self.routes = []  # 既存のルートをクリア
            
            for day in self.weekdays:
                # その日の利用者をフィルタリング
                day_users = [u for u in self.app.user_list if day in u.attendance_days]
                
                if day_users:
                    # 朝の送迎
                    morning_routes = optimizer.optimize_routes(
                        day_users, self.app.vehicle_list, self.app.staff_list, day, is_morning=True
                    )
                    self.routes.extend(morning_routes)
                    
                    # 夕方の送迎
                    evening_routes = optimizer.optimize_routes(
                        day_users, self.app.vehicle_list, self.app.staff_list, day, is_morning=False
                    )
                    self.routes.extend(evening_routes)
            
            # マップビューの初期化
            self.map_view = MapView(self.app.settings.get("facility_address", ""))
            
            # エクスポートマネージャーの初期化
            self.export_manager = ExportManager(self.app)
            
            # ルートの表示更新
            self.update_schedule_display()
            
            messagebox.showinfo("成功", "送迎ルートの最適化が完了しました。")
            
        except ImportError as e:
            # すでに処理済み
            pass
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            messagebox.showerror("エラー", f"計算中にエラーが発生しました: {str(e)}\n\n詳細:\n{error_details}")
        
        finally:
            # 進捗ダイアログを閉じる
            progress_bar.stop()
            progress.destroy()
    
    def update_schedule_display(self):
        """スケジュール表示の更新"""
        # タブをクリアして再作成
        for tab in self.notebook.tabs():
            self.notebook.forget(tab)
        
        # 現在の曜日と時間帯のルートをフィルタリング
        current_routes = [r for r in self.routes 
                          if r.date == self.selected_day and r.is_morning == self.is_morning]
        
        if not current_routes:
            # ルートがない場合はデフォルトタブを表示
            self.notebook.add(self.default_tab, text="スケジュールなし")
            return
        
        # 車両ごとのタブを作成
        for route in current_routes:
            tab = ttk.Frame(self.notebook)
            self.notebook.add(tab, text=f"{route.vehicle.name}")
            
            # タブ内のコンテンツ
            content_frame = ttk.Frame(tab, padding=10)
            content_frame.pack(fill=tk.BOTH, expand=True)
            
            # 車両情報
            vehicle_info = f"車両: {route.vehicle.name} (乗車可能人数: {route.vehicle.capacity}人)"
            ttk.Label(content_frame, text=vehicle_info, font=("", 12, "bold")).pack(anchor=tk.W, pady=(0, 10))
            
            # ドライバー情報
            driver_info = f"運転手: {route.driver.name}" if route.driver else "運転手: 未割り当て"
            ttk.Label(content_frame, text=driver_info).pack(anchor=tk.W)
            
            # 同乗スタッフ情報
            assistant_info = f"同乗スタッフ: {route.assistant.name}" if route.assistant else "同乗スタッフ: なし"
            ttk.Label(content_frame, text=assistant_info).pack(anchor=tk.W, pady=(0, 10))
            
            # 地図表示ボタン
            map_button = ttk.Button(content_frame, text="このルートを地図で表示", 
                                    command=lambda r=route: self.show_map_for_route(r))
            map_button.pack(anchor=tk.W, pady=(0, 10))
            
            # ルート詳細のツリービュー
            columns = ("order", "time", "user", "address", "is_pickup")
            route_tree = ttk.Treeview(content_frame, columns=columns, show="headings", height=10)
            
            route_tree.heading("order", text="順番")
            route_tree.heading("time", text="時間")
            route_tree.heading("user", text="利用者")
            route_tree.heading("address", text="住所")
            route_tree.heading("is_pickup", text="種別")
            
            route_tree.column("order", width=50)
            route_tree.column("time", width=80)
            route_tree.column("user", width=100)
            route_tree.column("address", width=200)
            route_tree.column("is_pickup", width=80)
            
            route_tree.pack(fill=tk.BOTH, expand=True)
            
            # ルートのデータを追加
            for i, stop in enumerate(route.stops):
                pickup_str = "迎え" if stop.is_pickup else "送り"
                
                route_tree.insert("", tk.END, values=(
                    i + 1,
                    stop.time,
                    stop.user.name if stop.user else "施設",
                    stop.user.address if stop.user else self.app.settings.get("facility_address", ""),
                    pickup_str
                ))
            
            # スクロールバー
            scrollbar = ttk.Scrollbar(content_frame, orient=tk.VERTICAL, command=route_tree.yview)
            route_tree.configure(yscroll=scrollbar.set)
            scrollbar.place(relx=1, rely=0, relheight=1, anchor=tk.NE)
    
    def show_map(self):
        """現在表示中の曜日・時間帯のルートを地図表示"""
        if not self.routes or not self.map_view:
            messagebox.showinfo("情報", "表示するルートがありません。まず「送迎ルート最適化」を実行してください。")
            return
        
        current_routes = [r for r in self.routes 
                          if r.date == self.selected_day and r.is_morning == self.is_morning]
        
        if not current_routes:
            messagebox.showinfo("情報", f"{self.selected_day}曜日 {self.time_var.get()}の送迎ルートはありません。")
            return
        
        # 現在表示中のタブのルートを取得
        current_tab = self.notebook.index("current")
        if current_tab < 0 or current_tab >= len(current_routes):
            # タブが選択されていない場合は最初のルートを表示
            route = current_routes[0]
        else:
            route = current_routes[current_tab]
        
        self.show_map_for_route(route)
    
    def show_map_for_route(self, route):
        """特定のルートを地図表示"""
        if not self.map_view:
            self.map_view = MapView(self.app.settings.get("facility_address", ""))
        
        # APIキーの取得
        api_key = self.app.settings.get("api_key", "")
        
        # APIキーの警告表示
        if not api_key:
            if not messagebox.askyesno("警告", 
                "Google Maps APIキーが設定されていません。\n"
                "地図表示は機能しない可能性があります。\n\n"
                "続行しますか？"):
                return
        else:
            # APIキーの形式チェック
            if not api_key.startswith("AIza"):
                messagebox.showwarning("警告", 
                    "Google Maps APIキーの形式が正しくない可能性があります。\n"
                    "APIキーは通常「AIza」で始まります。\n"
                    "また、以下の設定が必要です：\n"
                    "1. Google Maps JavaScript API が有効\n"
                    "2. Google Maps Geocoding API が有効\n"
                    "3. Google Maps Directions API が有効\n"
                    "4. APIキー制限が正しく設定されていること")
        
        try:
            filepath = self.map_view.create_map_for_route(route, api_key)
            if filepath:
                messagebox.showinfo("地図表示", "ブラウザでルートマップを開きました。\n"
                                           f"ファイルは {filepath} に保存されました。\n\n"
                                           "注意: 地図が正しく表示されない場合は、以下を確認してください：\n"
                                           "- APIキーが正しく設定されているか\n"
                                           "- APIキーに必要な権限があるか\n"
                                           "- ブラウザでJavaScriptが有効か")
            else:
                messagebox.showwarning("警告", "ルートマップの作成に失敗しました。")
        except Exception as e:
            messagebox.showerror("エラー", f"地図表示中にエラーが発生しました: {str(e)}")
    
    def export_to_excel(self):
        """送迎スケジュールをExcelにエクスポート"""
        if not self.routes:
            messagebox.showinfo("情報", "エクスポートするルートがありません。まず「送迎ルート最適化」を実行してください。")
            return
        
        if not self.export_manager:
            self.export_manager = ExportManager(self.app)
        
        try:
            filepath = self.export_manager.export_to_excel(self.routes)
            if filepath:
                if messagebox.askyesno("エクスポート完了", 
                                      f"Excelファイルが保存されました。\n\nファイル: {filepath}\n\nファイルを開きますか？"):
                    import os
                    os.startfile(filepath)
        except Exception as e:
            messagebox.showerror("エラー", f"Excelエクスポート中にエラーが発生しました: {str(e)}")
    
    def export_for_chat_gpt(self):
        """ChatGPTチェック用にテキスト形式でエクスポート"""
        if not self.routes:
            messagebox.showinfo("情報", "エクスポートするルートがありません。まず「送迎ルート最適化」を実行してください。")
            return
        
        if not self.export_manager:
            self.export_manager = ExportManager(self.app)
        
        try:
            filepath = self.export_manager.export_to_text(self.routes)
            if filepath:
                if messagebox.askyesno("エクスポート完了", 
                                      f"チェック用テキストファイルが保存されました。\n\nファイル: {filepath}\n\nファイルを開きますか？"):
                    import os
                    os.startfile(filepath)
        except Exception as e:
            messagebox.showerror("エラー", f"テキストエクスポート中にエラーが発生しました: {str(e)}")