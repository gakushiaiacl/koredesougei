import tkinter as tk
from tkinter import ttk
import webbrowser
import os
import json

class MapView:
    """Google Mapsを使用してルートを表示するクラス"""
    
    def __init__(self, facility_address):
        self.facility_address = facility_address
        self.html_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "data", "maps")
        os.makedirs(self.html_dir, exist_ok=True)
    
    def create_map_for_route(self, route, api_key=None):
        """ルートのマップを生成してブラウザで表示する"""
        if not route.stops:
            return None
            
        # HTMLファイルの作成
        filename = f"{route.date}_{route.is_morning}_{route.vehicle.name}.html".replace(" ", "_")
        filepath = os.path.join(self.html_dir, filename)
        
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(self._generate_map_html(route, api_key))
        
        # ブラウザでHTMLを開く
        webbrowser.open('file://' + os.path.realpath(filepath))
        return filepath
    
    def _generate_map_html(self, route, api_key=None):
        """Google Mapsを使用したHTMLを生成"""
        # ルートの表示名（日付と時間帯）
        time_of_day = "朝" if route.is_morning else "夕方"
        route_title = f"{route.date}曜日 {time_of_day} - {route.vehicle.name}"
        
        # マーカーとウェイポイントの作成
        waypoints = []
        markers = []
        
        # ルートの停留所を時間で並べ替え
        sorted_stops = sorted(route.stops, key=lambda x: x.time)
        
        for i, stop in enumerate(sorted_stops):
            # 施設または利用者の住所
            address = stop.user.address if stop.user else self.facility_address
            label = stop.user.name if stop.user else "施設"
            is_facility = stop.user is None
            
            if is_facility:
                # 施設は始点と終点になる
                icon_color = "red"
            else:
                # 利用者の場合
                icon_color = "blue"
            
            waypoints.append(address)
            
            # マーカー情報 - ここでは位置情報にインデックスではなく仮の座標を入れる
            marker = {
                "address": address,  # 住所をそのまま保持
                "label": f"{i+1}",
                "title": label,
                "icon": f"https://maps.google.com/mapfiles/ms/icons/{icon_color}-dot.png"
            }
            markers.append(marker)
        
        # APIキーが設定されているかどうかのメッセージ
        api_key_message = ""
        if not api_key:
            api_key_message = """
            <div style="background-color: #ffe0e0; padding: 10px; margin-bottom: 10px; border-radius: 5px;">
                <strong>注意:</strong> Google Maps APIキーが設定されていません。「設定」タブでAPIキーを設定してください。
            </div>
            """
        
        # HTMLテンプレート
        html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <title>{route_title}</title>
            <style>
                #map {{
                    height: 800px;
                    width: 100%;
                }}
                #info {{
                    padding: 10px;
                    background-color: #f5f5f5;
                    margin-bottom: 10px;
                }}
            </style>
        </head>
        <body>
            <div id="info">
                <h2>{route_title}</h2>
                <p>運転手: {route.driver.name if route.driver else '未割り当て'}</p>
                <p>同乗スタッフ: {route.assistant.name if route.assistant else 'なし'}</p>
            </div>
            {api_key_message}
            <div id="map"></div>
            
            <script>
                // Google Maps APIのコールバック関数
                function initMap() {{
                    const directionsService = new google.maps.DirectionsService();
                    const directionsRenderer = new google.maps.DirectionsRenderer({{
                        suppressMarkers: true
                    }});
                    
                    // マップの初期化
                    const map = new google.maps.Map(document.getElementById("map"), {{
                        zoom: 12,
                        center: {{ lat: 35.6895, lng: 139.6917 }}, // 東京
                    }});
                    
                    directionsRenderer.setMap(map);
                    
                    // マーカー情報とアドレス
                    const markerInfo = {json.dumps(markers)};
                    const addresses = markerInfo.map(m => m.address);
                    
                    // ジオコーディングを行う関数
                    const geocoder = new google.maps.Geocoder();
                    const geocodePromises = [];
                    
                    // 各アドレスのジオコーディングを非同期で実行
                    const geocodeAddress = (address, index) => {{
                        return new Promise((resolve, reject) => {{
                            geocoder.geocode({{ 'address': address }}, (results, status) => {{
                                if (status === 'OK') {{
                                    resolve({{ index: index, location: results[0].geometry.location }});
                                }} else {{
                                    console.error('Geocode failed: ' + status + ' for address: ' + address);
                                    // エラーの場合は東京駅の座標をデフォルトとして使用
                                    resolve({{ index: index, location: {{ lat: 35.6812, lng: 139.7671 }} }});
                                }}
                            }});
                        }});
                    }};
                    
                    // 全てのアドレスをジオコーディング
                    addresses.forEach((address, index) => {{
                        geocodePromises.push(geocodeAddress(address, index));
                    }});
                    
                    // 全てのジオコーディングが完了したら
                    Promise.all(geocodePromises).then(locations => {{
                        // マーカーの配置
                        locations.forEach(loc => {{
                            const idx = loc.index;
                            const markerData = markerInfo[idx];
                            // ここで正しい位置情報を使ってマーカーを作成
                            const marker = new google.maps.Marker({{
                                position: loc.location,  // ジオコードされた位置情報
                                map: map,
                                label: markerData.label,
                                title: markerData.title,
                                icon: markerData.icon
                            }});
                        }});
                        
                        // ルートの計算と表示
                        if (locations.length > 1) {{
                            const origin = locations[0].location;
                            const destination = locations[locations.length - 1].location;
                            
                            // 始点と終点以外のウェイポイント
                            const waypointsForRoute = locations.slice(1, locations.length - 1).map(loc => ({{
                                location: loc.location,
                                stopover: true
                            }}));
                            
                            directionsService.route(
                                {{
                                    origin: origin,
                                    destination: destination,
                                    waypoints: waypointsForRoute,
                                    optimizeWaypoints: false,
                                    travelMode: google.maps.TravelMode.DRIVING,
                                }},
                                (response, status) => {{
                                    if (status === "OK") {{
                                        directionsRenderer.setDirections(response);
                                    }} else {{
                                        window.alert("ルートの取得に失敗しました: " + status);
                                    }}
                                }}
                            );
                        }}
                    }}).catch(error => {{
                        console.error('Error in geocoding promises:', error);
                    }});
                }}
            </script>
            <script async defer
                src="https://maps.googleapis.com/maps/api/js?key={api_key or ''}&callback=initMap">
            </script>
        </body>
        </html>
        """
        
        return html 