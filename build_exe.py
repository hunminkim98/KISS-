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
    
    # PyInstaller 실행 파일 경로 찾기
    import site
    user_base = site.USER_BASE
    pyinstaller_exe = os.path.join(user_base, 'Python311', 'Scripts', 'pyinstaller.exe')
    # PyInstaller를 현재 python 실행환경에서 호출하도록 변경 (python -m PyInstaller)
    # 이렇게 하면 활성화된 가상환경/패키지 설치 경로를 일관되게 사용합니다.
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--onefile',                    # 단일 실행 파일 생성
        '--windowed',                   # GUI 프로그래밍 (콘솔 창 숨김)
        '--name=연구비처리도우미',        # 실행 파일 이름
        '--add-data=config.py:.',       # config.py 포함
        '--add-data=research_core.py:.', # research_core.py 포함
        '--add-data=research_gui.py:.',  # research_gui.py 포함
        # 일반적으로 필요한 hidden-imports
        '--hidden-import=pandas',
        '--hidden-import=openpyxl',
        '--hidden-import=tkinter',
        '--hidden-import=numpy',
        '--hidden-import=colorlog',
        '--hidden-import=psutil',
        '--hidden-import=pillow',
        '--hidden-import=xlsxwriter',
        # Excel 상호작용을 위해 xlwings 및 pywin32 관련 모듈을 명시적으로 포함
        '--hidden-import=xlwings',
        '--hidden-import=xlwings.server',
        '--hidden-import=xlwings._xlwindows',
        '--hidden-import=win32com',
        '--hidden-import=pythoncom',
        '--hidden-import=pywintypes',
        'main.py'
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
    
    # PyInstaller 실행 파일 경로 찾기
    import site
    user_base = site.USER_BASE
    pyinstaller_exe = os.path.join(user_base, 'Python311', 'Scripts', 'pyinstaller.exe')
    # 현재 python 환경에서 PyInstaller를 호출하도록 변경
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--onedir',                     # 폴더 형태로 생성
        '--windowed',                   # GUI 프로그래밍
        '--name=연구비처리도우미_포터블',  # 포터블 이름
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
        '--hidden-import=xlwings',
        '--hidden-import=xlwings.server',
        '--hidden-import=xlwings._xlwindows',
        '--hidden-import=win32com',
        '--hidden-import=pythoncom',
        '--hidden-import=pywintypes',
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
