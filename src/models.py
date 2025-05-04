class Staff:
    """職員を表すクラス"""
    def __init__(self, id=None, name="", can_drive=False, workdays=None, **kwargs):
        self.id = id  # 自動生成されるID
        self.name = name  # 名前
        self.can_drive = can_drive  # 運転可能かどうか
        self.workdays = workdays or []  # 勤務日 (例: ["月", "火", "水", "木", "金"])
    
    def to_dict(self):
        """辞書形式に変換"""
        return {
            "id": self.id,
            "name": self.name,
            "can_drive": self.can_drive,
            "workdays": self.workdays
        }

class User:
    """利用者を表すクラス"""
    def __init__(self, id=None, name="", address="", pickup_time_morning="", 
                 dropoff_time_morning="", pickup_time_evening="", 
                 dropoff_time_evening="", constraints="", attendance_days=None, **kwargs):
        self.id = id  # 自動生成されるID
        self.name = name  # 名前
        self.address = address  # 住所
        self.pickup_time_morning = pickup_time_morning  # 朝の送迎時間
        self.dropoff_time_morning = dropoff_time_morning  # 朝の到着時間
        self.pickup_time_evening = pickup_time_evening  # 夕方の送迎時間
        self.dropoff_time_evening = dropoff_time_evening  # 夕方の到着時間
        self.constraints = constraints  # 制約条件（特記事項）
        self.attendance_days = attendance_days or []  # 通所日
    
    def to_dict(self):
        """辞書形式に変換"""
        return {
            "id": self.id,
            "name": self.name,
            "address": self.address,
            "pickup_time_morning": self.pickup_time_morning,
            "dropoff_time_morning": self.dropoff_time_morning,
            "pickup_time_evening": self.pickup_time_evening,
            "dropoff_time_evening": self.dropoff_time_evening,
            "constraints": self.constraints,
            "attendance_days": self.attendance_days
        }

class Vehicle:
    """車両を表すクラス"""
    def __init__(self, id=None, name="", capacity=0, **kwargs):
        self.id = id  # 自動生成されるID
        self.name = name  # 車名
        self.capacity = capacity  # 乗車可能人数（利用者）
    
    def to_dict(self):
        """辞書形式に変換"""
        return {
            "id": self.id,
            "name": self.name,
            "capacity": self.capacity
        }

class RouteStop:
    """ルート上の停車地点を表すクラス"""
    def __init__(self, user=None, is_pickup=True, time=""):
        self.user = user  # 利用者
        self.is_pickup = is_pickup  # 迎えか送りか
        self.time = time  # 時間
    
    def to_dict(self):
        """辞書形式に変換"""
        return {
            "user_id": self.user.id if self.user else None,
            "is_pickup": self.is_pickup,
            "time": self.time
        }

class Route:
    """送迎ルートを表すクラス"""
    def __init__(self, id=None, vehicle=None, driver=None, assistant=None, 
                 stops=None, date=None, is_morning=True):
        self.id = id  # 自動生成されるID
        self.vehicle = vehicle  # 車両
        self.driver = driver  # 運転手
        self.assistant = assistant  # 同乗スタッフ
        self.stops = stops or []  # 停車地点のリスト
        self.date = date  # 日付
        self.is_morning = is_morning  # 朝か夕方か
    
    def to_dict(self):
        """辞書形式に変換"""
        return {
            "id": self.id,
            "vehicle_id": self.vehicle.id if self.vehicle else None,
            "driver_id": self.driver.id if self.driver else None,
            "assistant_id": self.assistant.id if self.assistant else None,
            "stops": [stop.to_dict() for stop in self.stops],
            "date": self.date,
            "is_morning": self.is_morning
        } 