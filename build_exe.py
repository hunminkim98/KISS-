#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
연구비 처리 자동화 프로그램 - 실행 파일 빌드 스크립트

작성자: 차세대지원팀 데이터 김훈민
작성일자: 2025-07-22
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

def clean_build_directories():
    """이전 빌드 결과물 정리"""
    print("🧹 이전 빌드 결과물 정리 중...")
    
    dirs_to_clean = ['build', 'dist', '__pycache__']
    files_to_clean = ['*.spec']
    
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"   ✅ {dir_name} 폴더 삭제 완료")
    
    # .spec 파일들 삭제
    for spec_file in Path('.').glob('*.spec'):
        spec_file.unlink()
        print(f"   ✅ {spec_file} 파일 삭제 완료")

def build_executable():
    """PyInstaller로 실행 파일 생성"""
    print("🚀 실행 파일 생성 시작...")
    
    # PyInstaller 명령어 구성
    cmd = [
        'pyinstaller',
        '--onefile',                    # 단일 실행 파일 생성
        '--windowed',                   # GUI 프로그램 (콘솔 창 숨김)
        '--name=연구비처리도우미',        # 실행 파일 이름
        '--add-data=config.py:.',       # config.py 포함
        '--add-data=research_core.py:.', # research_core.py 포함
        '--add-data=research_gui.py:.',  # research_gui.py 포함
        '--hidden-import=pandas',       # pandas 명시적 포함
        '--hidden-import=openpyxl',     # openpyxl 명시적 포함
        '--hidden-import=tkinter',      # tkinter 명시적 포함
        '--hidden-import=numpy',        # numpy 명시적 포함
        '--hidden-import=colorlog',     # colorlog 명시적 포함
        '--hidden-import=psutil',       # psutil 명시적 포함
        '--hidden-import=pillow',       # pillow 명시적 포함
        '--hidden-import=xlsxwriter',   # xlsxwriter 명시적 포함
        'main.py'                       # 메인 스크립트
    ]
    
    print(f"📋 실행 명령어: {' '.join(cmd)}")
    
    try:
        # PyInstaller 실행
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("✅ 실행 파일 생성 완료!")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"❌ 실행 파일 생성 실패:")
        print(f"   오류 코드: {e.returncode}")
        print(f"   오류 메시지: {e.stderr}")
        return False

def create_portable_version():
    """포터블 버전 생성 (폴더 형태)"""
    print("📦 포터블 버전 생성 시작...")
    
    cmd = [
        'pyinstaller',
        '--onedir',                     # 폴더 형태로 생성
        '--windowed',                   # GUI 프로그램
        '--name=연구비처리도우미_포터블',  # 폴더 이름
        '--add-data=config.py:.',
        '--add-data=research_core.py:.',
        '--add-data=research_gui.py:.',
        '--add-data=test:test',         # 테스트 폴더도 포함
        '--hidden-import=pandas',
        '--hidden-import=openpyxl',
        '--hidden-import=tkinter',
        '--hidden-import=numpy',
        '--hidden-import=colorlog',
        '--hidden-import=psutil',
        '--hidden-import=pillow',
        '--hidden-import=xlsxwriter',
        'main.py'
    ]
    
    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("✅ 포터블 버전 생성 완료!")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"❌ 포터블 버전 생성 실패:")
        print(f"   오류 코드: {e.returncode}")
        print(f"   오류 메시지: {e.stderr}")
        return False

def show_results():
    """빌드 결과 표시"""
    print("\n🎉 빌드 완료!")
    print("=" * 50)
    
    if os.path.exists('dist'):
        print("📁 생성된 파일들:")
        for item in os.listdir('dist'):
            item_path = os.path.join('dist', item)
            if os.path.isfile(item_path):
                size = os.path.getsize(item_path) / (1024 * 1024)  # MB
                print(f"   📄 {item} ({size:.1f} MB)")
            else:
                print(f"   📁 {item}/")
        
        print(f"\n📍 결과물 위치: {os.path.abspath('dist')}")
    else:
        print("❌ dist 폴더를 찾을 수 없습니다.")

def main():
    """메인 함수"""
    print("🐍 연구비 처리 자동화 프로그램 - 실행 파일 빌드")
    print("=" * 50)
    
    # 1. 이전 빌드 정리
    clean_build_directories()
    
    # 2. 단일 실행 파일 생성
    success1 = build_executable()
    
    # 3. 포터블 버전 생성
    success2 = create_portable_version()
    
    # 4. 결과 표시
    if success1 or success2:
        show_results()
        
        print("\n💡 사용 방법:")
        print("   • 단일 파일: dist/연구비처리도우미 실행")
        print("   • 포터블: dist/연구비처리도우미_포터블/ 폴더 내의 실행 파일 사용")
        print("   • 포터블 버전은 테스트 파일도 포함되어 있습니다.")
    else:
        print("\n❌ 빌드에 실패했습니다. 오류 메시지를 확인해주세요.")
        sys.exit(1)

if __name__ == "__main__":
    main()
