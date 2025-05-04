import numpy as np
import requests
import os
import json
import sys
import traceback
from datetime import datetime, timedelta

# ORToolsをインポート
try:
    from ortools.constraint_solver import routing_enums_pb2
    from ortools.constraint_solver import pywrapcp
except ImportError as e:
    print(f"ORToolsインポートエラー: {e}")
    print(f"Python検索パス: {sys.path}")
    print(f"現在のディレクトリ: {os.getcwd()}")
    traceback.print_exc()
    # アプリケーションを終了せずにダミーのクラスとして定義
    class routing_enums_pb2:
        class FirstSolutionStrategy:
            PATH_CHEAPEST_ARC = 0
    
    class pywrapcp:
        @staticmethod
        def RoutingIndexManager(*args, **kwargs):
            raise ImportError("ORToolsをインストールしてください: pip install ortools")
        
        @staticmethod
        def RoutingModel(*args, **kwargs):
            raise ImportError("ORToolsをインストールしてください: pip install ortools")
        
        @staticmethod
        def DefaultRoutingSearchParameters():
            return None

# モデルをインポート
try:
    from src.models import Route, RouteStop
except ImportError:
    try:
        from models import Route, RouteStop
    except ImportError as e:
        print(f"モデルインポートエラー: {e}")
        traceback.print_exc()

class TransportOptimizer:
    """送迎ルートの最適化を行うクラス"""
    
    def __init__(self, api_key, facility_address):
        """
        初期化
        
        Args:
            api_key: GoogleマップまたはOpenRouteServiceのAPIキー
            facility_address: 施設の住所
        """
        self.api_key = api_key
        self.facility_address = facility_address
        self.distance_matrix = None
        self.users = []
    
    def calculate_distance_matrix(self, users):
        """
        距離行列を計算
        
        Args:
            users: 利用者のリスト
        
        Returns:
            距離行列（施設と利用者間の移動時間を表す行列）
        """
        self.users = users
        addresses = [self.facility_address] + [user.address for user in users]
        n = len(addresses)
        
        # キャッシュファイルのパス
        cache_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data")
        cache_file = os.path.join(cache_dir, "distance_matrix_cache.json")
        
        # キャッシュがあれば読み込む
        if os.path.exists(cache_file):
            with open(cache_file, 'r') as f:
                cache = json.load(f)
                # キャッシュに必要な住所がすべて含まれているかチェック
                if all(addr in cache for addr in addresses):
                    matrix = np.zeros((n, n))
                    for i, from_addr in enumerate(addresses):
                        for j, to_addr in enumerate(addresses):
                            if i != j:  # 同じ場所の場合は0
                                matrix[i][j] = cache[from_addr].get(to_addr, 0)
                    return matrix
        
        # キャッシュがない場合はAPIで計算
        try:
            # Google Maps Distance Matrix API
            matrix = np.zeros((n, n))
            cache = {}
            
            for i, from_addr in enumerate(addresses):
                if from_addr not in cache:
                    cache[from_addr] = {}
                
                for j in range(n):
                    if i == j:
                        continue  # 同じ場所の場合は0
                    
                    to_addr = addresses[j]
                    
                    # すでに計算済みならスキップ
                    if to_addr in cache[from_addr]:
                        matrix[i][j] = cache[from_addr][to_addr]
                        continue
                    
                    url = f"https://maps.googleapis.com/maps/api/distancematrix/json"
                    params = {
                        "origins": from_addr,
                        "destinations": to_addr,
                        "key": self.api_key
                    }
                    
                    response = requests.get(url, params=params)
                    data = response.json()
                    
                    if data['status'] == 'OK':
                        # 秒単位の所要時間を取得
                        duration = data['rows'][0]['elements'][0]['duration']['value']
                        matrix[i][j] = duration
                        cache[from_addr][to_addr] = duration
            
            # キャッシュを保存
            os.makedirs(cache_dir, exist_ok=True)
            with open(cache_file, 'w') as f:
                json.dump(cache, f)
            
            return matrix
            
        except Exception as e:
            print(f"距離行列の計算に失敗しました: {e}")
            # エラーの場合はダミーデータを生成
            return np.random.randint(5, 30, size=(n, n)) * 60  # 5〜30分をランダムに設定
    
    def optimize_routes(self, users, vehicles, staff, day, is_morning=True):
        """
        指定された日の送迎ルートを最適化
        
        Args:
            users: その日に送迎が必要な利用者のリスト
            vehicles: 利用可能な車両のリスト
            staff: 利用可能なスタッフのリスト
            day: 曜日
            is_morning: 朝の送迎か夕方の送迎か
        
        Returns:
            最適化されたルートのリスト
        """
        if not users or not vehicles or not staff:
            return []
        
        # 利用可能なドライバーのフィルタリング
        available_drivers = [s for s in staff if s.can_drive and day in s.workdays]
        if not available_drivers:
            return []
        
        # 距離行列の計算
        self.distance_matrix = self.calculate_distance_matrix(users)
        
        # 車両ごとに最適化
        routes = []
        remaining_users = users.copy()
        drivers_assigned = []
        
        for vehicle in vehicles:
            if not remaining_users or not available_drivers:
                break
                
            # 車両の容量以下の利用者を選択
            vehicle_users = remaining_users[:min(len(remaining_users), vehicle.capacity)]
            
            # ドライバーの割り当て
            available_drivers_for_vehicle = [d for d in available_drivers if d not in drivers_assigned]
            if not available_drivers_for_vehicle:
                break
            
            driver = available_drivers_for_vehicle[0]
            drivers_assigned.append(driver)
            
            # 同乗スタッフの割り当て（同乗スタッフはドライバー以外から選ぶ）
            assistant = None
            available_assistants = [s for s in staff if s != driver and day in s.workdays and s not in drivers_assigned]
            if available_assistants:
                assistant = available_assistants[0]
                drivers_assigned.append(assistant)
            
            # この車両用のサブ問題を解く
            sub_matrix = np.zeros((len(vehicle_users) + 1, len(vehicle_users) + 1))
            for i in range(len(vehicle_users) + 1):
                for j in range(len(vehicle_users) + 1):
                    if i == 0:  # 施設からの経路
                        if j == 0:
                            sub_matrix[i][j] = 0
                        else:
                            sub_matrix[i][j] = self.distance_matrix[0][users.index(vehicle_users[j-1]) + 1]
                    elif j == 0:  # 施設への経路
                        sub_matrix[i][j] = self.distance_matrix[users.index(vehicle_users[i-1]) + 1][0]
                    else:  # 利用者間の経路
                        sub_matrix[i][j] = self.distance_matrix[users.index(vehicle_users[i-1]) + 1][users.index(vehicle_users[j-1]) + 1]
            
            # OR-Tools を使ったルート最適化
            optimal_route_indices = self._solve_vehicle_routing_problem(sub_matrix, vehicle_users)
            
            # ルートを構築
            route = Route(
                id=len(routes) + 1,
                vehicle=vehicle,
                driver=driver,
                assistant=assistant,
                date=day,
                is_morning=is_morning,
                stops=[]
            )
            
            # 施設の出発時間または到着時間を設定
            facility_time = "8:30" if is_morning else "16:00"  # デフォルト値
            facility_time_dt = datetime.strptime(facility_time, "%H:%M")
            
            # 時間順序が正しくなるように調整
            if is_morning:
                # 朝は施設から出発し、順番に利用者宅へ
                pickup_times = []
                current_time = facility_time_dt
                
                # まず施設を追加
                route.stops.append(RouteStop(
                    user=None,
                    is_pickup=False,  # 施設からの出発
                    time=current_time.strftime("%H:%M")
                ))
                
                last_idx = 0
                for i, idx in enumerate(optimal_route_indices):
                    if idx == 0:  # 施設は既に追加済み
                        continue
                        
                    user = vehicle_users[idx - 1]
                    
                    # 前の停留所からの移動時間計算
                    travel_time_sec = sub_matrix[last_idx][idx]
                    travel_time_min = travel_time_sec / 60
                    
                    # 時間を更新
                    current_time = current_time + timedelta(minutes=travel_time_min)
                    
                    # ルート停留所の追加
                    route.stops.append(RouteStop(
                        user=user,
                        is_pickup=True,  # 朝は迎え
                        time=current_time.strftime("%H:%M")
                    ))
                    
                    last_idx = idx
            else:
                # 夕方は利用者宅から施設へ
                # 夕方ルートではユーザー宅を先に訪問し、最後に施設に到着
                pickup_times = []
                # 施設到着時間から逆算
                current_time = facility_time_dt
                
                # 最後に施設を追加するため、一時保存
                last_stop = RouteStop(
                    user=None,
                    is_pickup=True,  # 施設への到着
                    time=current_time.strftime("%H:%M")
                )
                
                # ルートを逆順に処理（施設に向かう方向）
                reversed_indices = list(reversed(optimal_route_indices))
                last_idx = 0
                
                for i, idx in enumerate(reversed_indices):
                    if idx == 0:  # 施設は最後に追加
                        continue
                        
                    user = vehicle_users[idx - 1]
                    
                    # 次の停留所への移動時間計算
                    if i+1 < len(reversed_indices):
                        next_idx = reversed_indices[i+1]
                        travel_time_sec = sub_matrix[idx][next_idx]
                        travel_time_min = travel_time_sec / 60
                        
                        # 時間を逆算
                        pickup_time = current_time - timedelta(minutes=travel_time_min)
                    else:
                        # 最初の停留所の場合、施設からの出発時間を設定
                        pickup_time = current_time - timedelta(minutes=15)  # 仮の移動時間
                    
                    current_time = pickup_time
                    
                    # ルート停留所の追加
                    route.stops.append(RouteStop(
                        user=user,
                        is_pickup=False,  # 夕方は送り
                        time=current_time.strftime("%H:%M")
                    ))
                    
                    last_idx = idx
                
                # 施設を最後に追加
                route.stops.append(last_stop)
                
                # 停留所を時間順に並べ替え
                route.stops = sorted(route.stops, key=lambda x: x.time)
            
            routes.append(route)
            
            # 割り当てた利用者を残りのリストから削除
            for user in vehicle_users:
                remaining_users.remove(user)
        
        return routes
    
    def _solve_vehicle_routing_problem(self, distance_matrix, users):
        """
        OR-Toolsを使用して車両ルーティング問題を解く
        
        Args:
            distance_matrix: 距離行列
            users: 利用者のリスト
            
        Returns:
            最適なルートのインデックスリスト
        """
        # インデックス0は施設（デポ）
        manager = pywrapcp.RoutingIndexManager(len(distance_matrix), 1, 0)
        routing = pywrapcp.RoutingModel(manager)
        
        def distance_callback(from_index, to_index):
            from_node = manager.IndexToNode(from_index)
            to_node = manager.IndexToNode(to_index)
            return int(distance_matrix[from_node][to_node])
        
        transit_callback_index = routing.RegisterTransitCallback(distance_callback)
        routing.SetArcCostEvaluatorOfAllVehicles(transit_callback_index)
        
        # 解法のパラメータを設定
        search_parameters = pywrapcp.DefaultRoutingSearchParameters()
        search_parameters.first_solution_strategy = (
            routing_enums_pb2.FirstSolutionStrategy.PATH_CHEAPEST_ARC)
        search_parameters.time_limit.seconds = 10  # 計算時間制限
        
        # 問題を解く
        solution = routing.SolveWithParameters(search_parameters)
        
        if solution:
            route_indices = []
            index = routing.Start(0)
            while not routing.IsEnd(index):
                route_indices.append(manager.IndexToNode(index))
                index = solution.Value(routing.NextVar(index))
            route_indices.append(manager.IndexToNode(index))  # デポに戻る
            return route_indices
        else:
            # 解が見つからない場合、デフォルトルートを返す
            return list(range(len(distance_matrix))) 