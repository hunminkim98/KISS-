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
    # 2025년 기본 예산금액 설정 (나중에 추가됨)
    '2025_budget_amounts': {}
}


# 기본 예산금액 설정
BUDGET2025 = {
    '2025_budget_amounts': {
        '일용임금': 1256007740,
        '일용직 고용부담금': 277035006,
        '지급수수료': 82900000,
        '도서인쇄비': 37000000,
        '소모품비': 29347216,
        '홍보비': 0,
        '제세공과금': 6460038,
        '통신비': 700000,
        '보험료': 0,
        '피복비': 23400000,
        '임차료': 387000000,
        '유류비': 17400000,
        '시설장비유지비': 5000000,
        '급여성복리후생비(일용직)': 19800000,
        '일용직 비급여성비용': 11550000,
        '행사비': 0,
        '일반용역비': 281000000,
        '국내여비': 220000000,
        '국외업무여비': 36000000,
        '사업추진비': 29400000,
        '회의비': 0,
        '연구개발비': 90000000
    }
}

BUDGET2024 = {
    '2024_budget_amounts': {
        '일용임금': 689170381,
        '일용직 고용부담금': 11757858,
        '지급수수료': 70705927,
        '도서인쇄비': 2828080,
        '소모품비': 9722100,
        '홍보비': 2998250,
        '제세공과금': 3313080,
        '통신비': 194150,
        '보험료': 0,
        '피복비': 15600000,
        '임차료': 122480003,
        '유류비': 205000,
        '시설장비유지비': 0,
        '급여성복리후생비(일용직)': 6168154,
        '일용직 비급여성비용': 9982100,
        '행사비': 27653000,
        '일반용역비': 4617346,
        '국내여비': 94145553,
        '국외업무여비': 87572962,
        '사업추진비': 33032440,
        '회의비': 12898420,
        '연구개발비': 14239394,
        '자산취득비': 122480003
    }
}

BUDGET2023 = {
    '2023_budget_amounts': {
        '일용임금': 691030061,
        '일용직 고용부담금': 71814130,
        '지급수수료': 83413923,
        '도서인쇄비': 48650071,
        '소모품비': 14501530,
        '홍보비': 10992000,
        '제세공과금': 3188963,
        '통신비': 330000,
        '보험료': 0,
        '피복비': 14357100,
        '임차료': 118801805,
        '유류비': 3759858,
        '시설장비유지비': 0,
        '급여성복리후생비(일용직)': 252000,
        '일용직 비급여성비용': 10029600,
        '행사비': 56760000,
        '일반용역비': 78800000,
        '국내여비': 78758700,
        '국외업무여비': 23350489,
        '사업추진비': 9526058,
        '회의비': 14662950,
        '연구개발비': 230150000,
        '자산취득비': 9990000
    }
}

BUDGET2022 = {
    '2022_budget_amounts': {
        '일용임금': 704306201,
        '일용직 고용부담금': 117701339,
        '지급수수료': 85540875,
        '도서인쇄비': 55113610,
        '소모품비': 17729290,
        '홍보비': 12987690,
        '제세공과금': 0,
        '통신비': 198000,
        '보험료': 510250,
        '피복비': 11998000,
        '임차료': 109215965,
        '유류비': 5354443,
        '시설장비유지비': 0,
        '급여성복리후생비(일용직)': 8400660,
        '일용직 비급여성비용': 9639600,
        '행사비': 10108000,
        '일반용역비': 149965000,
        '국내여비': 56850424,
        '국외업무여비': 5377178,
        '사업추진비': 13198367,
        '회의비': 17803560,
        '연구개발비': 220163220,
        '재료비': 18295327
    }
}

# BUDGET_CLASSIFICATION의 2025_budget_amounts 업데이트
BUDGET_CLASSIFICATION['2025_budget_amounts'] = BUDGET2025['2025_budget_amounts']

# 연도별 예산 데이터 통합 딕셔너리
YEARLY_BUDGET_DATA = {
    '2022': BUDGET2022['2022_budget_amounts'],
    '2023': BUDGET2023['2023_budget_amounts'],
    '2024': BUDGET2024['2024_budget_amounts'],
    '2025': BUDGET2025['2025_budget_amounts']
}

def create_yearly_pivot_data():
    """
    연도별 예산 데이터를 피벗 테이블용 세로형 데이터로 변환합니다.
    
    Returns:
        list: [{'연도': year, '예산과목': item, '예산금액': amount}, ...] 형태의 데이터
    """
    pivot_data = []
    
    for year, budget_data in YEARLY_BUDGET_DATA.items():
        for budget_item, amount in budget_data.items():
            pivot_data.append({
                '연도': year,
                '예산과목': budget_item,
                '예산금액': amount
            })
    
    return pivot_data

def get_all_budget_items():
    """
    모든 연도의 예산과목을 통합한 리스트를 반환합니다.
    
    Returns:
        list: 모든 예산과목의 고유 리스트 (정렬됨)
    """
    all_items = set()
    
    for budget_data in YEARLY_BUDGET_DATA.values():
        all_items.update(budget_data.keys())
    
    return sorted(list(all_items))

def get_yearly_budget_summary():
    """
    연도별 예산 총액 요약을 반환합니다.
    
    Returns:
        dict: {year: total_amount} 형태의 연도별 총 예산
    """
    summary = {}
    
    for year, budget_data in YEARLY_BUDGET_DATA.items():
        summary[year] = sum(budget_data.values())
    
    return summary

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

# xlwings 대화형 피벗 테이블 설정
ENABLE_INTERACTIVE_PIVOT = True  # 대화형 피벗 테이블 활성화 여부
PIVOT_SHEET_NAME = '예산분석'
PIVOT_CHART_TITLE = '예산 분석 차트'

# 피벗 테이블 구조 설정
PIVOT_CONFIG = {
    'row_fields': ['예산과목'],                          # 행: 예산과목
    'data_fields': ['예산금액', '예산잔액', '집행률'],      # 원본 값들을 그대로 표시
    'slicer_fields': ['예산과목'],  # 예산과목 슬라이서만 추가 (값 필드 슬라이서는 별도 처리)
    'chart_position': {
        'top': 50,
        'left': 600,
        'width': 400,
        'height': 300
    },
    'slicer_positions': {
        'budget_item': {  # 예산과목 슬라이서
            'top': 50,
            'left': 1050,
            'width': 200,
            'height': 300
        },
        'metric': {  # 측정항목 슬라이서
            'top': 370,
            'left': 1050,
            'width': 200,
            'height': 150
        }
    }
}

