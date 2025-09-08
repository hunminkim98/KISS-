#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
KISS 연구비 처리 웹 애플리케이션 - pywebview 메인 앱
작성자: 차세대지원팀 데이터 김훈민
"""

import webview
import threading
import logging
from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS
import os
import sys
from pathlib import Path

# 기존 코어 로직 임포트
from research_core import ExcelFileLoader, DataClassifier, ExcelExporter
from config import WINDOW_TITLE

class KISSWebApp:
    def __init__(self):
        self.app = Flask(__name__, static_folder='web/build', static_url_path='')
        CORS(self.app)
        
        # 기존 로직 인스턴스
        self.file_loader = ExcelFileLoader()
        self.classifier = DataClassifier()
        self.exporter = ExcelExporter()
        
        self.setup_routes()
        
    def setup_routes(self):
        """Flask 라우트 설정"""
        
        @self.app.route('/')
        def serve_react():
            return send_from_directory(self.app.static_folder, 'index.html')
            
        @self.app.route('/api/health')
        def health_check():
            return jsonify({"status": "healthy", "message": "KISS 연구비 처리 시스템"})
            
        @self.app.route('/api/upload', methods=['POST'])
        def upload_file():
            try:
                if 'file' not in request.files:
                    return jsonify({"error": "파일이 없습니다"}), 400
                    
                file = request.files['file']
                if file.filename == '':
                    return jsonify({"error": "파일이 선택되지 않았습니다"}), 400
                
                # 임시 파일 저장
                temp_path = f"./temp_{file.filename}"
                file.save(temp_path)
                
                # 파일 로드
                if self.file_loader.load_file(temp_path):
                    file_info = self.file_loader.get_data_info()
                    return jsonify({
                        "success": True,
                        "message": "파일 업로드 성공",
                        "data": file_info
                    })
                else:
                    return jsonify({"error": "파일 처리 중 오류가 발생했습니다"}), 500
                    
            except Exception as e:
                logging.error(f"파일 업로드 오류: {str(e)}")
                return jsonify({"error": str(e)}), 500
                
        @self.app.route('/api/classify', methods=['POST'])
        def classify_data():
            try:
                if self.file_loader.data is None:
                    return jsonify({"error": "먼저 파일을 업로드하세요"}), 400
                    
                # 데이터 분류
                result = self.classifier.classify_data(self.file_loader.data)
                
                # 분류 통계
                stats = self.classifier.get_classification_stats()
                
                return jsonify({
                    "success": True,
                    "message": "데이터 분류 완료",
                    "stats": stats,
                    "business_count": len(result['business']) if result['business'] is not None else 0,
                    "research_count": len(result['research']) if result['research'] is not None else 0
                })
                
            except Exception as e:
                logging.error(f"데이터 분류 오류: {str(e)}")
                return jsonify({"error": str(e)}), 500
                
        @self.app.route('/api/data/<data_type>')
        def get_data(data_type):
            try:
                if data_type == 'business':
                    data = self.classifier.business_data
                elif data_type == 'research':
                    data = self.classifier.research_data
                elif data_type == 'unclassified':
                    data = self.classifier.unclassified_data
                else:
                    return jsonify({"error": "잘못된 데이터 타입"}), 400
                    
                if data is None or data.empty:
                    return jsonify({"data": [], "total": 0})
                    
                # 데이터를 JSON으로 변환
                records = data.head(100).to_dict('records')  # 처음 100개만
                
                return jsonify({
                    "data": records,
                    "total": len(data),
                    "columns": list(data.columns)
                })
                
            except Exception as e:
                logging.error(f"데이터 조회 오류: {str(e)}")
                return jsonify({"error": str(e)}), 500
                
        @self.app.route('/api/export', methods=['POST'])
        def export_excel():
            try:
                if self.classifier.business_data is None and self.classifier.research_data is None:
                    return jsonify({"error": "분류된 데이터가 없습니다"}), 400
                    
                # Excel 내보내기
                output_path = "./출력_연구비_집행관리.xlsx"
                if self.exporter.export_to_excel(
                    self.classifier.business_data,
                    self.classifier.research_data,
                    output_path
                ):
                    return jsonify({
                        "success": True,
                        "message": "Excel 파일 생성 완료",
                        "path": output_path
                    })
                else:
                    return jsonify({"error": "Excel 파일 생성 실패"}), 500
                    
            except Exception as e:
                logging.error(f"Excel 내보내기 오류: {str(e)}")
                return jsonify({"error": str(e)}), 500

    def run_flask(self):
        """Flask 서버 실행"""
        self.app.run(host='127.0.0.1', port=5000, debug=False, threaded=True)

def main():
    """메인 함수"""
    # 로깅 설정
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler('kiss_webapp.log', encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    
    # 웹 앱 인스턴스 생성
    webapp = KISSWebApp()
    
    # Flask 서버를 별도 스레드에서 실행
    flask_thread = threading.Thread(target=webapp.run_flask)
    flask_thread.daemon = True
    flask_thread.start()
    
    # pywebview 윈도우 생성
    webview.create_window(
        title=WINDOW_TITLE + " - 웹 버전",
        url='http://127.0.0.1:5000',
        width=1400,
        height=900,
        min_size=(1200, 800),
        resizable=True
    )
    
    # 윈도우 시작
    webview.start(debug=False)

if __name__ == '__main__':
    main()