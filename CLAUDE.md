# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

KISS (Korean Institute for Strategic Studies) Research Fund Processing Assistant - A Python desktop application for automating research fund budget classification and reporting for a Korean research institution.

**Author**: 차세대지원팀 데이터 김훈민 (Next Generation Support Team Data Kim Hoon-min)  
**Version**: v0.1  
**Language**: Python 3 with Korean language support (UTF-8)

## Development Commands

### Running the Application
```bash
# Run the GUI application
python main.py

# Install dependencies
pip install -r requirements.txt
```

### Building Executable
```bash
# Build both single-file and portable versions
python build_exe.py

# Manual PyInstaller command for single executable
pyinstaller --onefile --windowed --name=연구비처리도우미 main.py
```

### Testing
- No automated test framework configured
- Test data available in `test/` directory with sample Excel files
- Manual testing through GUI with provided Excel samples

## Core Architecture

### Main Components

**Entry Point**: `main.py`
- Sets up logging configuration
- Initializes and runs the GUI application
- Handles top-level exception management

**GUI Layer**: `research_gui.py` 
- `ResearchFundGUI` class: Main Tkinter-based desktop interface
- File selection, data preview, classification controls
- Progress feedback and result display

**Core Logic**: `research_core.py` (2700+ lines)
- `ExcelFileLoader`: Excel file reading and validation
- `DataClassifier`: Automatic classification of budget items into business vs research categories
- `ExcelExporter`: Multi-sheet Excel output generation with styling
- `SummarySheetGenerator`: Budget summary calculations and formatting  
- `DashboardGenerator`: Visual dashboard creation with charts
- `TotalSheetGenerator`: Consolidated budget totals across categories

**Configuration**: `config.py`
- All application settings, UI parameters, colors, fonts
- Budget classification mappings and hierarchies
- Historical budget data for 2022-2025 (Korean won amounts)
- Excel styling and column width specifications

### Data Flow Architecture

1. **Input**: Excel files containing financial transaction data
2. **Classification**: Automatic categorization based on description prefixes:
   - Business funds: entries starting with "25 차세대" 
   - Research funds: entries starting with "25 심층연구"
3. **Processing**: Multi-sheet Excel generation with:
   - Separate sheets for business vs research expenditures
   - Budget summary with actuals vs planned amounts
   - Dashboard with charts and KPI visualizations
   - Total summary across all categories

### Budget Classification System

**Hierarchical Structure**: 예산목 (Budget Category) → 세목 (Sub-category) → 예산과목 (Budget Item)

**Major Categories**:
- 인건비 (Personnel): Daily wages, employment contributions
- 민간이전 (Private Transfer): Employment burden fees  
- 운영비 (Operating): General supplies, utilities, facilities, welfare
- 여비 (Travel): Domestic and international travel
- 업무추진비 (Business Promotion): Meeting and event expenses

**Key Features**:
- Multi-year budget comparison (2022-2025)
- Execution rate calculations
- Automated budget remaining calculations
- Korean currency formatting (원)

## Korean Language Considerations

- All user-facing text in Korean
- UTF-8 encoding throughout
- Korean currency formatting (원)
- Date formatting in Korean style
- Excel sheet names in Korean
- Log files with Korean characters: `연구비_처리.log`

## Dependencies

**Core Libraries**:
- `pandas`: Excel data processing and manipulation
- `openpyxl`: Excel file generation with advanced formatting
- `tkinter`: Desktop GUI framework
- `logging`: Application logging with Korean text support

**Optional/Build**:
- `pyinstaller`: Executable creation
- `xlsxwriter`: Additional Excel features
- `pillow`: Image processing for charts
- `colorlog`: Enhanced logging output

## File Structure Patterns

**Input Files**: Excel files (.xlsx, .xls) with financial transaction data
**Output Files**: Multi-sheet Excel workbooks with Korean sheet names
**Logs**: UTF-8 encoded log files with Korean messages
**Test Data**: Sample Excel files in `test/` directory

## Configuration Management

All settings centralized in `config.py`:
- UI layout parameters (fonts, colors, dimensions)
- Budget classification mappings  
- Historical budget amounts by year
- Excel formatting and styling rules
- Korean text constants and messages

Critical configuration changes require restarting the application.