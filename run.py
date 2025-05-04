#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
これで送迎 - 送迎スケジューリングアプリケーション
起動用スクリプト
"""

import os
import sys
import tkinter as tk
import traceback
import logging

# ロギングの設定
log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
os.makedirs(log_dir, exist_ok=True)
log_file = os.path.join(log_dir, "app.log")

logging.basicConfig(
    filename=log_file,
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

def exception_handler(exc_type, exc_value, exc_traceback):
    """未処理の例外をログに記録"""
    logging.error("未処理の例外:", exc_info=(exc_type, exc_value, exc_traceback))
    # 標準のエラーハンドラーも呼び出す
    sys.__excepthook__(exc_type, exc_value, exc_traceback)

# 未処理の例外ハンドラーを設定
sys.excepthook = exception_handler

def main():
    """メイン関数"""
    try:
        # 現在のディレクトリをPATHに追加
        current_dir = os.path.dirname(os.path.abspath(__file__))
        logging.info(f"現在のディレクトリ: {current_dir}")
        if current_dir not in sys.path:
            sys.path.insert(0, current_dir)
            logging.info(f"PATHに追加: {current_dir}")
        
        # src ディレクトリへのパスを追加
        src_dir = os.path.join(current_dir, "src")
        logging.info(f"ソースディレクトリ: {src_dir}")
        if src_dir not in sys.path:
            sys.path.insert(0, src_dir)
            logging.info(f"PATHに追加: {src_dir}")
        
        # Python検索パスをログに記録
        logging.info("Python検索パス:")
        for path in sys.path:
            logging.info(f"  {path}")
        
        # ORToolsのインポートチェック
        try:
            import ortools
            logging.info(f"ORTools version: {ortools.__version__}")
        except ImportError as e:
            logging.error(f"ORToolsのインポートに失敗: {e}")
            raise
        
        # サンプルデータ作成フラグの処理
        create_sample = "--sample" in sys.argv
        
        # アプリケーションのインポート
        logging.info("アプリケーションモジュールをインポート")
        from src.main import TransportApp, create_sample_data
        
        # アプリケーションを起動
        logging.info("アプリケーションの起動")
        root = tk.Tk()
        app = TransportApp(root)
        
        # サンプルデータを作成
        if create_sample:
            logging.info("サンプルデータの作成")
            create_sample_data(app)
        
        logging.info("メインループの開始")
        root.mainloop()
        logging.info("アプリケーションの終了")
        
    except Exception as e:
        logging.error(f"アプリケーションの起動エラー: {e}")
        traceback.print_exc(file=open(log_file, 'a'))
        raise

if __name__ == "__main__":
    main() 