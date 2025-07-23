#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
연구비 처리 자동화 프로그램 설정 파일

작성자: 차세대지원팀 데이터 김훈민
작성일자: 2025-07-22
"""

# 데이터 분류 기준
BUSINESS_PREFIX = '25 차세대'
RESEARCH_PREFIX = '25 심층연구'

# GUI 설정
WINDOW_TITLE = "KISS 연구비 처리 도우미 v0.1"
WINDOW_SIZE = "800x600"
MAIN_FONT = ("Arial", 16, "bold")
NORMAL_FONT = ("Arial", 10)
BUTTON_FONT = ("Arial", 10, "bold")

# 색상 설정
BUTTON_COLOR = "#4CAF50"
BUTTON_TEXT_COLOR = "black"

# 파일 관련 설정
SUPPORTED_EXTENSIONS = ['.xlsx', '.xls']
FILE_DIALOG_TYPES = [
    ("Excel files", "*.xlsx *.xls"),
    ("All files", "*.*")
]

# 로깅 설정
LOG_FILE = '연구비_처리.log'
LOG_FORMAT = '%(asctime)s - %(levelname)s - %(message)s'

# 데이터 처리 설정
SUMMARY_COLUMN = '적요'
UNCLASSIFIED_WARNING_THRESHOLD = 0.7  # 70% 이상 미분류 시 경고

# GUI 레이아웃 설정
TEXT_AREA_HEIGHT = 10
TEXT_AREA_WIDTH = 90
BUTTON_WIDTH = 12
CLASSIFY_BUTTON_WIDTH = 20
CLASSIFY_BUTTON_HEIGHT = 2
FILE_PATH_LABEL_WIDTH = 50

# 샘플 데이터 표시 개수
SAMPLE_DATA_COUNT = 3
INFO_SAMPLE_COUNT = 5

# Excel 출력 관련 설정
OUTPUT_SHEET_NAMES = {
    'business': '집행관리(사업비)',
    'research': '집행관리(연구비)'
}

# 출력할 컬럼 매핑 (원본 컬럼명 -> 출력 컬럼명)
OUTPUT_COLUMNS = [
    '결의서',      # 1st
    '발의일자',    # 2nd
    '번호',        # 3rd
    '적요',        # 4th
    '작성자',      # 5th
    '총지급액',    # 6th
    '예산과목'     # 14th
]

# 연구비 시트 전용 추가 컬럼
RESEARCH_ADDITIONAL_COLUMNS = ['연구자', '반영일']

# 파일 저장 관련 설정
SAVE_FILE_TYPES = [
    ("Excel files", "*.xlsx"),
    ("All files", "*.*")
]
DEFAULT_OUTPUT_FILENAME = "연구비_집행관리"

# Excel 스타일링 설정
EXCEL_STYLING = {
    # 컬럼 너비 설정
    'column_widths': {
        '결의서': 12,
        '발의일자': 15,
        '번호': 10,
        '적요': 50,      # 적요는 내용이 길어서 넓게
        '작성자': 12,
        '총지급액': 15,
        '예산과목': 15,
        '연구자': 15,
        '반영일': 15
    },

    # 헤더 스타일
    'header_style': {
        'fill_color': '366092',      # 파란색 배경
        'font_color': 'FFFFFF',      # 흰색 글자
        'bold': True
    }
}

# 사업비 요약 시트 관련 설정
# 예산과목을 예산목-세목-예산과목 계층으로 매핑하는 설정
BUDGET_CLASSIFICATION = {
    # 예산목 (대분류) 매핑
    'budget_categories': {
        '인건비': {
            'subcategories': {
                '일용임금': ['일용임금']
            }
        },
        '민간이전': {
            'subcategories': {
                '고용부담금': ['일용직 고용부담금']
            }
        },
        '운영비': {
            'subcategories': {
                '일반수용비': ['지급수수료', '도서인쇄비', '소모품비', '홍보비'],
                '공공요금및제세': ['제세공과금', '통신비', '보험료'],
                '피복비': ['피복비'],
                '임차료': ['임차료'],
                '유류비': ['유류비'],
                '시설장비유지비': ['시설장비유지비'],
                '복리후생비': ['급여성복리후생비(일용직)', '일용직 비급여성비용'],
                '일반용역비': ['행사비', '일반용역비']
            }
        },
        '여비': {
            'subcategories': {
                '국내여비': ['국내여비'],
                '국외여비': ['국외업무여비']
            }
        },
        '업무추진비': {
            'subcategories': {
                '사업추진비': ['사업추진비', '회의비']
            }
        }
    },

    # 기본 예산금액 설정 (모두 0으로 설정)
    'default_budget_amounts': {
        '일용임금': 0,
        '일용직 고용부담금': 0,
        '지급수수료': 0,
        '도서인쇄비': 0,
        '소모품비': 0,
        '홍보비': 0,
        '제세공과금': 0,
        '통신비': 0,
        '보험료': 0,
        '피복비': 0,
        '임차료': 0,
        '유류비': 0,
        '시설장비유지비': 0,
        '급여성복리후생비(일용직)': 0,
        '일용직 비급여성비용': 0,
        '행사비': 0,
        '일반용역비': 0,
        '국내여비': 0,
        '국외업무여비': 0,
        '사업추진비': 0,
        '회의비': 0
    }
}

# 사업비 요약 시트 컬럼 설정
SUMMARY_SHEET_COLUMNS = [
    '예산목',
    '세목',
    '예산과목',
    '예산금액',
    '지출액',
    '예산잔액',
    '집행률'
]

# 총액 시트 컬럼 설정 (센터와 심층연구로 지출액 분리)
TOTAL_SHEET_COLUMNS = [
    '예산목',
    '세목',
    '예산과목',
    '예산금액',
    '센터',
    '심층연구',
    '예산잔액',
    '집행률'
]

