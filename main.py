import os
import json
import re
import shutil
import sys
import tempfile
import threading
import unicodedata
import zipfile
import urllib.error
import urllib.request
import ctypes
import ctypes.wintypes
from collections import Counter, defaultdict
from copy import copy, deepcopy
import tkinter as tk
from tkinter import filedialog, messagebox

pythoncom = None
win32 = None
com_error = Exception

PLACEHOLDER_SHEET_BASENAME = '__XRF_TEMP__'
OPTIMIZABLE_EXTENSIONS = {'.xlsx', '.xlsm', '.xltx', '.xltm'}
EXCEL_FILE_TYPES = [
    ('Excel files', '*.xlsx *.xls *.xlsm *.xltx *.xltm'),
    ('All files', '*.*'),
]
ELEMENT_SYMBOLS = ('Cd', 'Pb', 'Hg', 'Br', 'Cr')
ELEMENT_KEYS = {symbol: symbol.upper() for symbol in ELEMENT_SYMBOLS}
HEADER_SCAN_LIMIT = 40
APP_VERSION = '1.1.0'
OPTIMIZE_COMPRESSLEVEL = 4
GITHUB_REPO_OWNER = 'WooDoGwon'
GITHUB_REPO_NAME = 'Program'
GITHUB_RELEASE_API_URL = f'https://api.github.com/repos/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/releases/latest'
GITHUB_RELEASE_PAGE_URL = f'https://github.com/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/releases'
UPDATE_ASSET_NAME_PREFIX = 'XRF_Report_Auto_Input_Setup_v'
UPDATE_DOWNLOAD_DIR_NAME = 'updates'
UPDATE_CHECK_TIMEOUT_SECONDS = 8
LOGO_FILE_NAME = 'Logo.png'
TOP_BAR_LOGO_HEIGHT = 44
TOP_BAR_LOGO_MAX_WIDTH = 170
ABOUT_LOGO_HEIGHT = 58
ABOUT_LOGO_MAX_WIDTH = 210
VERSION_HISTORY = [
    ('1.1.0', '\ubc84\uc804 \ud45c\uae30 1.1.0 \uc801\uc6a9', '2026-04-22'),
    ('1.0.10', '\u0047\u0069\u0074\u0048\u0075\u0062 \uc790\ub3d9 \uc5c5\ub370\uc774\ud2b8 \uae30\ub2a5 \ucd94\uac00', '2026-04-22'),
    ('1.0.9', '버전 업데이트 / 노트북 화면 자동 맞춤 개선', '2026-04-21'),
    ('1.0.8', '버전 업데이트 / 기존 버전 파일 정리 / Excel Value 접근 개선 / 해상도 프리셋 확장', '2026-04-20'),
    ('1.0.7', '정식 버전 / 화면 모드 전환 / 최소 창 크기 고정', '2026-04-20'),
    ('1.0.6', '로고/Help 메뉴/버전 정보 UI 개선', '2026-04-16'),
    ('1.0.5', '업체 TOTAL 기준 순번 표시', '2026-03-31'),
    ('1.0.0', '초기 배포 버전', '-'),
]
VERSION_INFO_MIN_ROWS = 4
DEFAULT_WINDOW_SIZE = (1360, 840)
DEFAULT_MIN_WINDOW_SIZE = (1220, 760)
MIN_WINDOW_SIZE = (800, 600)
UI_SCALE_MIN = 0.55
UI_SCALE_MAX = 2.0
SCREEN_FIT_PADDING = (24, 32)
UI_SCALE_PRESETS = (
    ('100%', 1.0),
    ('110%', 1.1),
    ('125%', 1.25),
    ('150%', 1.5),
)
RESOLUTION_PRESETS = (
    ('SVGA (800 x 600)', (800, 600)),
    ('XGA (1024 x 768)', (1024, 768)),
    ('XGA+ (1152 x 864)', (1152, 864)),
    ('HD (1280 x 720)', (1280, 720)),
    ('WXGA (1280 x 800)', (1280, 800)),
    ('HD+ (1360 x 768)', (1360, 768)),
    ('기본 (1360 x 840)', (1360, 840)),
    ('HD+ (1366 x 768)', (1366, 768)),
    ('WXGA+ (1440 x 900)', (1440, 900)),
    ('HD++ (1600 x 900)', (1600, 900)),
    ('WSXGA+ (1680 x 1050)', (1680, 1050)),
    ('FHD (1920 x 1080)', (1920, 1080)),
)

DISPLAY_SETTINGS_DIR_NAME = 'XRF_Report_Auto_Input_System'
DISPLAY_SETTINGS_FILE_NAME = 'display_settings.json'

TARGET_INSPECTION_CACHE = {}
TOTAL_LAYOUT_CACHE = {}
TOTAL_RECORDS_CACHE = {}
EXCEL_APP = None
EXCEL_COM_INITIALIZED = False


def get_file_cache_key(path):
    try:
        stat = os.stat(path)
    except OSError:
        return None
    return (normalized_path(path), stat.st_mtime_ns, stat.st_size)


def clear_cached_path_entries(cache, path):
    path_key = normalized_path(path)
    stale_keys = [key for key in cache if key[0] == path_key]
    for key in stale_keys:
        cache.pop(key, None)


def clone_layout_metadata(layout):
    if not layout:
        return None

    cloned = {}
    for key, value in layout.items():
        if key == 'sheet':
            continue
        if key == 'columns' and value:
            cloned[key] = dict(value)
        elif key == 'elements' and value:
            cloned[key] = {symbol: dict(columns) for symbol, columns in value.items()}
        else:
            cloned[key] = value
    return cloned


def get_cached_target_inspection(path):
    cache_key = get_file_cache_key(path)
    if cache_key is None:
        return None

    cached = TARGET_INSPECTION_CACHE.get(cache_key)
    if not cached:
        return None

    return {
        'path': os.path.abspath(path),
        'sheet_names': list(cached['sheet_names']),
        'layout': clone_layout_metadata(cached['layout']),
    }


def set_cached_target_inspection(path, sheet_names, layout):
    cache_key = get_file_cache_key(path)
    if cache_key is None:
        return

    clear_cached_path_entries(TARGET_INSPECTION_CACHE, path)
    TARGET_INSPECTION_CACHE[cache_key] = {
        'sheet_names': list(sheet_names),
        'layout': clone_layout_metadata(layout),
    }


def get_cached_total_layout(path):
    cache_key = get_file_cache_key(path)
    if cache_key is None:
        return None

    cached = TOTAL_LAYOUT_CACHE.get(cache_key)
    return clone_layout_metadata(cached) if cached else None


def set_cached_total_layout(path, layout):
    cache_key = get_file_cache_key(path)
    if cache_key is None:
        return

    clear_cached_path_entries(TOTAL_LAYOUT_CACHE, path)
    TOTAL_LAYOUT_CACHE[cache_key] = clone_layout_metadata(layout)


def peek_cached_total_records(path):
    cache_key = get_file_cache_key(path)
    if cache_key is None:
        return None

    return TOTAL_RECORDS_CACHE.get(cache_key)


def has_cached_total_records(path):
    return peek_cached_total_records(path) is not None


def get_cached_total_records(path):
    cached = peek_cached_total_records(path)
    return deepcopy(cached) if cached else None


def set_cached_total_records(path, records):
    cache_key = get_file_cache_key(path)
    if cache_key is None:
        return

    clear_cached_path_entries(TOTAL_RECORDS_CACHE, path)
    TOTAL_RECORDS_CACHE[cache_key] = deepcopy(records)

SOURCE_HEADER_ALIASES = {
    'measurement': {'NO', '측정번호', '번호', '순번'},
    'recipe': {'MODE', '모드', '레시피명', '레시피', '구분', '분류', '재질구분'},
    'date': {'검사일', '검사일자', '측정일자', '측정일', '일자'},
    'company': {'업체명', '회사명', '협력업체', '업체', '업체코드', '업체명코드'},
    'vendor': {'AZENTEK NO.', 'AZENTEK NO', 'AZENTEKNO', 'AZENTEKNUMBER', '업체번호'},
    'product': {'품명', '제품명', '샘플명', '품번', '품 번'},
    'material': {'재질', '재료', '원자재', '재질명'},
    'judgment': {'종합판정', '판정', '결과', '적합여부', '합부'},
    'spot': {'SPOT TEST', 'SPOTTEST', '발색여부'},
}

TARGET_HEADER_ALIASES = {
    'measurement': {'측정번호', '번호', 'NO', '순번'},
    'recipe': {'레시피명', 'MODE', '레시피', '구분', '분류'},
    'date': {'측정일자', '검사일', '검사일자', '측정일', '일자'},
    'company': {'업체명', '회사명', '협력업체', '업체', '업체코드', '업체명코드'},
    'vendor': {'업체번호', 'AZENTEK NO.', 'AZENTEK NO', 'AZENTEKNO'},
    'product': {'품명', '제품명', '샘플명', '품번', '품 번'},
    'material': {'재질', '재료', '원자재', '재질명'},
    'judgment': {'종합판정', '판정', '결과', '적합여부', '합부'},
    'spot': {'SPOT TEST', 'SPOTTEST', '발색여부'},
}


def ensure_excel_available():
    global pythoncom, win32, com_error
    if pythoncom is not None and win32 is not None:
        return
    try:
        import pythoncom as imported_pythoncom
        import win32com.client as imported_win32
        from pywintypes import com_error as imported_com_error
    except ImportError as exc:
        raise RuntimeError('이 기능을 사용하려면 Microsoft Excel 과 pywin32 환경이 필요합니다.') from exc
    pythoncom = imported_pythoncom
    win32 = imported_win32
    com_error = imported_com_error


def ensure_excel_application():
    global EXCEL_APP, EXCEL_COM_INITIALIZED
    ensure_excel_available()

    if not EXCEL_COM_INITIALIZED:
        pythoncom.CoInitialize()
        EXCEL_COM_INITIALIZED = True

    if EXCEL_APP is not None:
        try:
            _ = EXCEL_APP.Workbooks.Count
            configure_excel_app(EXCEL_APP)
            return EXCEL_APP
        except Exception:
            try:
                EXCEL_APP.Quit()
            except Exception:
                pass
            EXCEL_APP = None

    EXCEL_APP = win32.DispatchEx('Excel.Application')
    configure_excel_app(EXCEL_APP)
    return EXCEL_APP


def shutdown_excel_application():
    global EXCEL_APP, EXCEL_COM_INITIALIZED

    if EXCEL_APP is not None:
        try:
            EXCEL_APP.Quit()
        except Exception:
            pass
        EXCEL_APP = None

    if EXCEL_COM_INITIALIZED and pythoncom is not None:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        EXCEL_COM_INITIALIZED = False


def build_file_labels(paths):
    base_names = [os.path.basename(path) for path in paths]
    duplicates = Counter(base_names)
    used_labels = Counter()
    labels = []
    for path, base_name in zip(paths, base_names):
        if duplicates[base_name] > 1:
            parent_name = os.path.basename(os.path.dirname(path)) or os.path.dirname(path)
            label = f'{base_name} [{parent_name}]'
        else:
            label = base_name
        used_labels[label] += 1
        if used_labels[label] > 1:
            label = f'{label} ({used_labels[label]})'
        labels.append(label)
    return labels


def configure_excel_app(excel):
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.EnableEvents = False
    try:
        excel.DisplayStatusBar = False
    except Exception:
        pass
    try:
        excel.AskToUpdateLinks = False
    except Exception:
        pass
    try:
        excel.Calculation = -4135
    except Exception:
        pass
    try:
        excel.AutomationSecurity = 3
    except Exception:
        pass


def open_excel_workbook(excel, path, read_only):
    return excel.Workbooks.Open(
        Filename=os.path.abspath(path),
        UpdateLinks=0,
        ReadOnly=read_only,
        IgnoreReadOnlyRecommended=True,
        Notify=False,
        AddToMru=False,
    )


def normalized_path(path):
    return os.path.normcase(os.path.abspath(path))


def resource_path(relative_name):
    base_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_dir, relative_name)

def application_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def display_settings_dir():
    base_dir = os.environ.get('LOCALAPPDATA') or os.environ.get('APPDATA') or application_dir()
    return os.path.join(base_dir, DISPLAY_SETTINGS_DIR_NAME)


def display_settings_path():
    return os.path.join(display_settings_dir(), DISPLAY_SETTINGS_FILE_NAME)


def update_download_dir():
    return os.path.join(display_settings_dir(), UPDATE_DOWNLOAD_DIR_NAME)


def parse_version_parts(version_text):
    parts = []
    for part in re.findall(r'\d+', str(version_text or '')):
        try:
            parts.append(int(part))
        except ValueError:
            parts.append(0)
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts[:4])


def normalize_release_version(version_text):
    text = str(version_text or '').strip()
    if text.lower().startswith('v'):
        text = text[1:]
    match = re.search(r'\d+(?:\.\d+){1,3}', text)
    return match.group(0) if match else text


def is_newer_version(candidate_version, current_version=APP_VERSION):
    return parse_version_parts(candidate_version) > parse_version_parts(current_version)


def choose_update_asset(assets):
    setup_assets = []
    fallback_assets = []
    for asset in assets or []:
        name = str(asset.get('name') or '')
        download_url = asset.get('browser_download_url')
        if not download_url or not name.lower().endswith('.exe'):
            continue
        if name.startswith(UPDATE_ASSET_NAME_PREFIX):
            setup_assets.append(asset)
        elif 'setup' in name.lower():
            fallback_assets.append(asset)
    return (setup_assets or fallback_assets or [None])[0]


def fetch_latest_release_info():
    request = urllib.request.Request(
        GITHUB_RELEASE_API_URL,
        headers={
            'Accept': 'application/vnd.github+json',
            'User-Agent': 'XRF-Report-Auto-Input-Updater',
        },
    )
    with urllib.request.urlopen(request, timeout=UPDATE_CHECK_TIMEOUT_SECONDS) as response:
        data = json.loads(response.read().decode('utf-8'))

    version = normalize_release_version(data.get('tag_name') or data.get('name') or '')
    asset = choose_update_asset(data.get('assets') or [])
    return {
        'version': version,
        'tag_name': data.get('tag_name') or '',
        'name': data.get('name') or '',
        'body': data.get('body') or '',
        'html_url': data.get('html_url') or GITHUB_RELEASE_PAGE_URL,
        'asset_name': asset.get('name') if asset else '',
        'asset_url': asset.get('browser_download_url') if asset else '',
    }


def download_update_installer(download_url, asset_name):
    if not download_url:
        raise RuntimeError('\uc5c5\ub370\uc774\ud2b8 \uc124\uce58\ud30c\uc77c \ub2e4\uc6b4\ub85c\ub4dc \uc8fc\uc18c\uac00 \uc5c6\uc2b5\ub2c8\ub2e4.')

    safe_name = os.path.basename(asset_name or f'{UPDATE_ASSET_NAME_PREFIX}{normalize_release_version(APP_VERSION)}.exe')
    target_dir = update_download_dir()
    os.makedirs(target_dir, exist_ok=True)
    target_path = os.path.join(target_dir, safe_name)
    temp_path = target_path + '.download'

    request = urllib.request.Request(
        download_url,
        headers={'User-Agent': 'XRF-Report-Auto-Input-Updater'},
    )
    with urllib.request.urlopen(request, timeout=UPDATE_CHECK_TIMEOUT_SECONDS) as response, open(temp_path, 'wb') as file:
        shutil.copyfileobj(response, file)

    if os.path.exists(target_path):
        os.remove(target_path)
    os.replace(temp_path, target_path)
    return target_path


def load_logo_image(max_width, max_height):
    logo_path = resource_path(LOGO_FILE_NAME)
    if not os.path.isfile(logo_path):
        return None

    try:
        image = tk.PhotoImage(file=logo_path)
    except tk.TclError:
        return None

    width = image.width()
    height = image.height()
    if width <= 0 or height <= 0:
        return None

    width_factor = max(1, (width + max_width - 1) // max_width)
    height_factor = max(1, (height + max_height - 1) // max_height)
    scale_factor = max(width_factor, height_factor)
    if scale_factor > 1:
        image = image.subsample(scale_factor, scale_factor)

    return image


def load_top_bar_logo_image():
    return load_logo_image(TOP_BAR_LOGO_MAX_WIDTH, TOP_BAR_LOGO_HEIGHT)


def load_about_logo_image():
    return load_logo_image(ABOUT_LOGO_MAX_WIDTH, ABOUT_LOGO_HEIGHT)
def make_placeholder_sheet_name(existing_names):
    candidate = PLACEHOLDER_SHEET_BASENAME
    index = 1
    while candidate in existing_names:
        candidate = f'{PLACEHOLDER_SHEET_BASENAME}_{index}'
        index += 1
    return candidate


def format_file_size(size_in_bytes):
    size = float(size_in_bytes)
    for unit in ('B', 'KB', 'MB', 'GB'):
        if size < 1024 or unit == 'GB':
            if unit == 'B':
                return f'{int(size)}{unit}'
            return f'{size:.1f}{unit}'
        size /= 1024
    return f'{size_in_bytes}B'


def optimize_excel_file(path):
    extension = os.path.splitext(path)[1].lower()
    if extension not in OPTIMIZABLE_EXTENSIONS:
        return {'applied': False, 'reason': 'unsupported_extension', 'extension': extension}
    if not zipfile.is_zipfile(path):
        return {'applied': False, 'reason': 'not_zip_container', 'extension': extension}

    before_size = os.path.getsize(path)
    file_dir = os.path.dirname(path) or os.getcwd()
    file_ext = os.path.splitext(path)[1]
    fd, temp_path = tempfile.mkstemp(prefix='xrf_opt_', suffix=file_ext, dir=file_dir)
    os.close(fd)

    try:
        with zipfile.ZipFile(path, 'r') as source_zip, zipfile.ZipFile(
            temp_path,
            'w',
            compression=zipfile.ZIP_DEFLATED,
            compresslevel=OPTIMIZE_COMPRESSLEVEL,
            allowZip64=True,
        ) as target_zip:
            for info in source_zip.infolist():
                data = b'' if info.is_dir() else source_zip.read(info.filename)
                new_info = copy(info)
                new_info.compress_type = zipfile.ZIP_STORED if info.is_dir() else zipfile.ZIP_DEFLATED
                target_zip.writestr(
                    new_info,
                    data,
                    compress_type=new_info.compress_type,
                    compresslevel=OPTIMIZE_COMPRESSLEVEL,
                )
        shutil.copystat(path, temp_path)
        os.replace(temp_path, path)
    except Exception as exc:
        if os.path.exists(temp_path):
            os.remove(temp_path)
        raise RuntimeError('파일 최적화 중 오류가 발생했습니다.') from exc

    after_size = os.path.getsize(path)
    return {
        'applied': True,
        'before_size': before_size,
        'after_size': after_size,
        'saved_bytes': before_size - after_size,
    }


def normalize_header_text(value):
    if value is None:
        return ''
    text = str(value).strip().upper()
    text = text.replace('\n', '').replace('\r', '').replace('\t', '')
    text = text.replace(' ', '')
    return re.sub(r'[^0-9A-Z가-힣]+', '', text)


def normalize_match_text(value):
    if value is None:
        return ''

    text = unicodedata.normalize('NFKC', str(value)).strip().upper()
    if not text:
        return ''

    text = re.sub(r'\.(XLSX?|XLSM|XLTX|XLTM)$', '', text)
    text = text.replace('&', 'AND')
    return re.sub('[^0-9A-Zㄱ-ㆎ가-힣]+', '', text)



def normalize_material_fallback_key(value):
    if value is None:
        return ''

    text = unicodedata.normalize('NFKC', str(value)).strip().upper()
    if not text:
        return ''

    text = re.sub(r'\.(XLSX?|XLSM|XLTX|XLTM)$', '', text)
    text = text.replace('&', 'AND')
    return re.sub(r'[^0-9A-Z]+', '', text)



def is_styler_xrf_path(path):
    return 'STYLER' in normalize_match_text(os.path.basename(path))


def normalize_date_match_text(value):
    if value in (None, ''):
        return ''
    text = str(value).strip()
    if not text:
        return ''

    parts = re.findall(r'\d+', text)
    if len(parts) >= 3 and len(parts[0]) == 4:
        year = parts[0]
        month = parts[1].zfill(2)
        day = parts[2].zfill(2)
        return f'{year}{month}{day}'

    digits = ''.join(parts)
    if len(digits) >= 8:
        return digits[:8]
    return digits


def excel_cell_text(sheet, row_index, col_index):
    if not col_index:
        return ''
    cell = sheet.Cells(row_index, col_index)
    return str(cell.Text).strip()


def excel_cell_value(sheet, row_index, col_index):
    if not col_index:
        return None
    cell = sheet.Cells(row_index, col_index)
    try:
        return cell.Value2
    except Exception:
        return cell.Value


def excel_set_cell_value(sheet, row_index, col_index, value):
    if not col_index:
        return
    cell = sheet.Cells(row_index, col_index)
    try:
        cell.Value2 = value
    except Exception:
        cell.Value = value


def excel_cell_display_or_value(sheet, row_index, col_index):
    text = excel_cell_text(sheet, row_index, col_index)
    if text:
        return text
    return excel_cell_value(sheet, row_index, col_index)
def get_sheet_bounds(sheet):
    used_range = sheet.UsedRange
    first_row = used_range.Row
    first_col = used_range.Column
    last_row = first_row + used_range.Rows.Count - 1
    last_col = first_col + used_range.Columns.Count - 1
    return first_row, first_col, last_row, last_col


def coerce_range_values_to_matrix(values, row_count, col_count):
    matrix = [[None for _ in range(col_count)] for _ in range(row_count)]
    if row_count <= 0 or col_count <= 0 or values is None:
        return matrix

    if row_count == 1 and col_count == 1:
        matrix[0][0] = values
        return matrix

    if row_count == 1:
        source_row = values
        if isinstance(source_row, tuple) and source_row and isinstance(source_row[0], tuple):
            source_row = source_row[0]
        elif not isinstance(source_row, (tuple, list)):
            source_row = (source_row,)
        for col_offset, value in enumerate(source_row[:col_count]):
            matrix[0][col_offset] = value
        return matrix

    if col_count == 1:
        if not isinstance(values, (tuple, list)):
            values = (values,)
        for row_offset, row_value in enumerate(values[:row_count]):
            if isinstance(row_value, (tuple, list)):
                matrix[row_offset][0] = row_value[0] if row_value else None
            else:
                matrix[row_offset][0] = row_value
        return matrix

    if not isinstance(values, (tuple, list)):
        values = (values,)
    for row_offset, row_values in enumerate(values[:row_count]):
        if isinstance(row_values, (tuple, list)):
            for col_offset, value in enumerate(row_values[:col_count]):
                matrix[row_offset][col_offset] = value
        else:
            matrix[row_offset][0] = row_values
    return matrix


def get_normalized_sheet_range(sheet, first_row, last_row, first_col, last_col):
    row_count = last_row - first_row + 1
    col_count = last_col - first_col + 1
    if row_count <= 0 or col_count <= 0:
        return []

    values = sheet.Range(sheet.Cells(first_row, first_col), sheet.Cells(last_row, last_col)).Value2
    matrix = coerce_range_values_to_matrix(values, row_count, col_count)
    return [[normalize_header_text(value) for value in row_values] for row_values in matrix]


def row_contains_subheaders_in_values(row_values):
    return any(value in {'농도', '오차', '발색여부'} for value in row_values)


def find_header_column_in_rows(row_values_by_row, aliases, first_col):
    normalized_aliases = {normalize_header_text(alias) for alias in aliases}
    if not row_values_by_row:
        return None

    col_count = len(row_values_by_row[0])
    for col_offset in range(col_count):
        for row_values in row_values_by_row:
            if row_values[col_offset] in normalized_aliases:
                return first_col + col_offset
    return None


def find_element_columns_in_rows(header_values, subheader_values, first_col):
    element_columns = {}
    col_count = len(header_values)
    for symbol, symbol_key in ELEMENT_KEYS.items():
        header_key = normalize_header_text(symbol)
        for col_offset in range(max(0, col_count - 1)):
            if header_values[col_offset] != header_key:
                continue
            if subheader_values[col_offset] == '농도' and subheader_values[col_offset + 1] == '오차':
                element_columns[symbol_key] = {'conc': first_col + col_offset, 'err': first_col + col_offset + 1}
                break
    return element_columns


def find_spot_test_column_in_rows(header_values, subheader_values, first_col, aliases):
    normalized_aliases = {normalize_header_text(alias) for alias in aliases}
    for col_offset, header_value in enumerate(header_values):
        sub_value = subheader_values[col_offset]
        if header_value in normalized_aliases or sub_value in normalized_aliases:
            return first_col + col_offset
    return None


def detect_source_layout_in_sheet(sheet):
    first_row, first_col, last_row, last_col = get_sheet_bounds(sheet)
    limit_row = min(last_row - 1, first_row + HEADER_SCAN_LIMIT - 1)
    if limit_row < first_row:
        return None

    normalized_rows = get_normalized_sheet_range(sheet, first_row, limit_row + 1, first_col, last_col)
    if len(normalized_rows) < 2:
        return None

    best_layout = None
    sheet_name_key = normalize_header_text(sheet.Name)
    for row_offset, header_row in enumerate(range(first_row, limit_row + 1)):
        sub_offset = row_offset + 1
        if sub_offset >= len(normalized_rows):
            break

        header_values = normalized_rows[row_offset]
        subheader_values = normalized_rows[sub_offset]
        if not row_contains_subheaders_in_values(subheader_values):
            continue

        columns = {
            key: find_header_column_in_rows((header_values, subheader_values), aliases, first_col)
            for key, aliases in SOURCE_HEADER_ALIASES.items()
            if key != 'spot'
        }
        element_columns = find_element_columns_in_rows(header_values, subheader_values, first_col)
        spot_column = find_spot_test_column_in_rows(
            header_values,
            subheader_values,
            first_col,
            SOURCE_HEADER_ALIASES['spot'],
        )
        if (columns.get('material') is None and columns.get('product') is None) or len(element_columns) < 3:
            continue

        score = len(element_columns) * 10
        score += sum(2 for col in columns.values() if col is not None)
        if columns.get('material') is None and columns.get('product') is not None:
            score += 3
        if spot_column is not None:
            score += 1
        if 'TOTAL' in sheet_name_key:
            score += 4

        layout = {
            'sheet': sheet,
            'sheet_name': sheet.Name,
            'header_row': header_row,
            'subheader_row': header_row + 1,
            'data_start_row': header_row + 2,
            'columns': columns,
            'spot_column': spot_column,
            'elements': element_columns,
            'score': score,
        }
        if best_layout is None or layout['score'] > best_layout['score']:
            best_layout = layout

    return best_layout


def detect_target_layout_in_sheet(sheet):
    first_row, first_col, last_row, last_col = get_sheet_bounds(sheet)
    limit_row = min(last_row - 1, first_row + HEADER_SCAN_LIMIT - 1)
    if limit_row < first_row:
        return None

    normalized_rows = get_normalized_sheet_range(sheet, first_row, limit_row + 1, first_col, last_col)
    if len(normalized_rows) < 2:
        return None

    best_layout = None
    sheet_name_text = str(sheet.Name)
    for row_offset, header_row in enumerate(range(first_row, limit_row + 1)):
        sub_offset = row_offset + 1
        if sub_offset >= len(normalized_rows):
            break

        header_values = normalized_rows[row_offset]
        subheader_values = normalized_rows[sub_offset]
        if not row_contains_subheaders_in_values(subheader_values):
            continue

        columns = {
            key: find_header_column_in_rows((header_values, subheader_values), aliases, first_col)
            for key, aliases in TARGET_HEADER_ALIASES.items()
            if key != 'spot'
        }
        element_columns = find_element_columns_in_rows(header_values, subheader_values, first_col)
        spot_column = find_spot_test_column_in_rows(
            header_values,
            subheader_values,
            first_col,
            TARGET_HEADER_ALIASES['spot'],
        )
        if columns.get('material') is None or len(element_columns) < 3:
            continue

        score = len(element_columns) * 10
        score += sum(2 for col in columns.values() if col is not None)
        if spot_column is not None:
            score += 1
        if '유해' in sheet_name_text or '분석' in sheet_name_text:
            score += 6

        layout = {
            'sheet': sheet,
            'sheet_name': sheet.Name,
            'header_row': header_row,
            'subheader_row': header_row + 1,
            'data_start_row': header_row + 2,
            'columns': columns,
            'spot_column': spot_column,
            'elements': element_columns,
            'score': score,
        }
        if best_layout is None or layout['score'] > best_layout['score']:
            best_layout = layout

    return best_layout


def iter_priority_workbook_sheets(workbook, keywords):
    sheets = list(workbook.Worksheets)
    if len(sheets) <= 1:
        return sheets

    normalized_keywords = [normalize_header_text(keyword) for keyword in keywords]

    def sheet_priority(sheet):
        name_key = normalize_header_text(sheet.Name)
        return sum(1 for keyword in normalized_keywords if keyword and keyword in name_key)

    return sorted(sheets, key=sheet_priority, reverse=True)


def detect_total_layout_in_workbook(workbook):
    best_layout = None
    for sheet in iter_priority_workbook_sheets(workbook, ('TOTAL', '업체TOTAL', 'TOTALDATA')):
        layout = detect_source_layout_in_sheet(sheet)
        if layout and (best_layout is None or layout['score'] > best_layout['score']):
            best_layout = layout
            if 'TOTAL' in normalize_header_text(sheet.Name) and layout['score'] >= 40:
                return best_layout
    return best_layout


def detect_target_layout_in_workbook(workbook):
    best_layout = None
    for sheet in iter_priority_workbook_sheets(workbook, ('유해', '분석', 'HAZARD')):
        layout = detect_target_layout_in_sheet(sheet)
        if layout and (best_layout is None or layout['score'] > best_layout['score']):
            best_layout = layout
            sheet_name_key = normalize_header_text(sheet.Name)
            if layout['score'] >= 40 and ('유해' in sheet_name_key or '분석' in sheet_name_key or 'HAZARD' in sheet_name_key):
                return best_layout
    return best_layout


def inspect_target_workbook_with_excel(excel, path):

    workbook = None
    try:
        workbook = open_excel_workbook(excel, path, read_only=True)
        sheet_names = [sheet.Name for sheet in workbook.Worksheets]
        layout = detect_target_layout_in_workbook(workbook)
        return sheet_names, layout
    finally:
        if workbook is not None:
            workbook.Close(False)


def inspect_target_workbook(path):
    try:
        excel = ensure_excel_application()
        return inspect_target_workbook_with_excel(excel, path)
    except com_error as exc:
        raise RuntimeError('기존 저장 파일을 읽는 중 오류가 발생했습니다.') from exc


def inspect_total_workbook(path):
    abs_path = os.path.abspath(path)
    cached_layout = get_cached_total_layout(abs_path)
    if cached_layout is not None and has_cached_total_records(abs_path):
        return cached_layout

    workbook = None
    try:
        excel = ensure_excel_application()
        workbook = open_excel_workbook(excel, abs_path, read_only=True)
        layout = detect_total_layout_in_workbook(workbook)
        set_cached_total_layout(abs_path, layout)
        cache_total_records_from_layout(abs_path, workbook, layout)
        return clone_layout_metadata(layout)
    except com_error as exc:
        raise RuntimeError('업체 TOTAL 파일을 읽는 중 오류가 발생했습니다.') from exc
    finally:
        if workbook is not None:
            workbook.Close(False)

def list_workbook_sheet_names(path):
    sheet_names, _ = inspect_target_workbook(path)
    return sheet_names


def intersect_sheet_names_in_order(sheet_name_lists):
    if not sheet_name_lists:
        return []
    common_names = set(sheet_name_lists[0])
    for sheet_names in sheet_name_lists[1:]:
        common_names &= set(sheet_names)
    return [name for name in sheet_name_lists[0] if name in common_names]


def inspect_target_workbooks(paths):
    inspections = []
    failures = []
    seen = set()
    ordered_paths = []

    for path in paths:
        abs_path = os.path.abspath(path)
        path_key = normalized_path(abs_path)
        if path_key in seen:
            continue
        seen.add(path_key)
        ordered_paths.append(abs_path)

    uncached_paths = [path for path in ordered_paths if get_cached_target_inspection(path) is None]

    if uncached_paths:
        excel = ensure_excel_application()
        for abs_path in uncached_paths:
            try:
                sheet_names, layout = inspect_target_workbook_with_excel(excel, abs_path)
            except com_error:
                failures.append({'path': abs_path, 'error': '기존 저장 파일을 읽는 중 오류가 발생했습니다.'})
                continue
            set_cached_target_inspection(abs_path, sheet_names, layout)

    failed_keys = {normalized_path(item['path']) for item in failures}
    for abs_path in ordered_paths:
        if normalized_path(abs_path) in failed_keys:
            continue
        cached = get_cached_target_inspection(abs_path)
        if cached is not None:
            inspections.append(cached)

    common_sheet_names = intersect_sheet_names_in_order([item['sheet_names'] for item in inspections])
    detected_sheet_names = [item['layout']['sheet_name'] for item in inspections if item['layout']]
    return inspections, failures, common_sheet_names, detected_sheet_names


def inspect_target_workbooks_in_background(paths):
    inspections = []
    failures = []
    seen = set()
    ordered_paths = []

    for path in paths:
        abs_path = os.path.abspath(path)
        path_key = normalized_path(abs_path)
        if path_key in seen:
            continue
        seen.add(path_key)
        ordered_paths.append(abs_path)

    uncached_paths = [path for path in ordered_paths if get_cached_target_inspection(path) is None]

    excel = None
    if uncached_paths:
        ensure_excel_available()
        pythoncom.CoInitialize()
        try:
            excel = win32.DispatchEx('Excel.Application')
            configure_excel_app(excel)

            for abs_path in uncached_paths:
                try:
                    sheet_names, layout = inspect_target_workbook_with_excel(excel, abs_path)
                except com_error:
                    failures.append({'path': abs_path, 'error': '기존 저장 파일을 읽는 중 오류가 발생했습니다.'})
                    continue
                set_cached_target_inspection(abs_path, sheet_names, layout)
        finally:
            if excel is not None:
                excel.Quit()
            pythoncom.CoUninitialize()

    failed_keys = {normalized_path(item['path']) for item in failures}
    for abs_path in ordered_paths:
        if normalized_path(abs_path) in failed_keys:
            continue
        cached = get_cached_target_inspection(abs_path)
        if cached is not None:
            inspections.append(cached)

    common_sheet_names = intersect_sheet_names_in_order([item['sheet_names'] for item in inspections])
    detected_sheet_names = [item['layout']['sheet_name'] for item in inspections if item['layout']]
    return inspections, failures, common_sheet_names, detected_sheet_names


def inspect_total_workbook_in_background(path):
    abs_path = os.path.abspath(path)
    cached_layout = get_cached_total_layout(abs_path)
    if cached_layout is not None and has_cached_total_records(abs_path):
        return cached_layout

    ensure_excel_available()
    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32.DispatchEx('Excel.Application')
        configure_excel_app(excel)
        workbook = open_excel_workbook(excel, abs_path, read_only=True)
        layout = detect_total_layout_in_workbook(workbook)
        set_cached_total_layout(abs_path, layout)
        cache_total_records_from_layout(abs_path, workbook, layout)
        return clone_layout_metadata(layout)
    except com_error as exc:
        raise RuntimeError('업체 TOTAL 파일을 읽는 중 오류가 발생했습니다.') from exc
    finally:
        if workbook is not None:
            workbook.Close(False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()

def build_source_total_records(sheet, layout):
    _, _, last_row, _ = get_sheet_bounds(sheet)
    columns = layout['columns']
    records = []

    for row_index in range(layout['data_start_row'], last_row + 1):
        measurement = excel_cell_display_or_value(sheet, row_index, columns.get('measurement'))
        recipe = excel_cell_display_or_value(sheet, row_index, columns.get('recipe'))
        date_value = excel_cell_display_or_value(sheet, row_index, columns.get('date'))
        company = excel_cell_display_or_value(sheet, row_index, columns.get('company'))
        vendor = excel_cell_display_or_value(sheet, row_index, columns.get('vendor'))
        product = excel_cell_display_or_value(sheet, row_index, columns.get('product'))
        material = excel_cell_display_or_value(sheet, row_index, columns.get('material'))
        judgment = excel_cell_display_or_value(sheet, row_index, columns.get('judgment'))
        spot_value = excel_cell_display_or_value(sheet, row_index, layout.get('spot_column'))

        element_values = {}
        has_element_data = False
        for symbol_key, symbol_columns in layout['elements'].items():
            conc_value = excel_cell_display_or_value(sheet, row_index, symbol_columns['conc'])
            err_value = excel_cell_display_or_value(sheet, row_index, symbol_columns['err'])
            if conc_value not in (None, '') or err_value not in (None, ''):
                has_element_data = True
            element_values[symbol_key] = {'conc': conc_value, 'err': err_value}

        if not any([measurement, recipe, date_value, company, vendor, product, material, judgment, spot_value, has_element_data]):
            continue

        material_text = str(material).strip() if material not in (None, '') else ''
        product_value = product
        if not material_text and product_value not in (None, ''):
            material_text = str(product_value).strip()
        if product_value in (None, '') and material_text:
            product_value = material_text
        company_value = company if company not in (None, '') else vendor

        records.append(
            {
                'row_index': row_index,
                'measurement': measurement,
                'recipe': recipe,
                'date': date_value,
                'date_key': normalize_date_match_text(date_value),
                'company': company_value,
                'vendor': vendor,
                'product': product_value,
                'material': material_text,
                'judgment': judgment,
                'spot': spot_value,
                'elements': element_values,
                'material_key': normalize_match_text(material_text),
                'material_fallback_key': normalize_material_fallback_key(material_text),
                'used': False,
            }
        )

    return records


def cache_total_records_from_layout(path, workbook, layout):
    if not layout:
        return []

    cached_records = peek_cached_total_records(path)
    if cached_records is not None:
        return deepcopy(cached_records)

    source_sheet = layout.get('sheet')
    if source_sheet is None and workbook is not None and layout.get('sheet_name'):
        try:
            source_sheet = workbook.Worksheets(layout['sheet_name'])
        except Exception:
            source_sheet = None

    if source_sheet is None:
        return []

    layout_with_sheet = dict(layout)
    layout_with_sheet['sheet'] = source_sheet
    records = build_source_total_records(source_sheet, layout_with_sheet)
    set_cached_total_records(path, records)
    return records


def format_total_order_value(measurement, fallback_index):
    if measurement not in (None, ''):
        text = str(measurement).strip()
        if text:
            if re.fullmatch(r'\d+\.0+', text):
                return text.split('.', 1)[0]
            return text
    return str(fallback_index)


def build_total_sheet_order_lookup(records):
    order_lookup = {}

    for fallback_index, record in enumerate(records, start=1):
        order_value = format_total_order_value(record.get('measurement'), fallback_index)
        candidate_keys = (
            record.get('material_key', ''),
            record.get('material_fallback_key', ''),
            normalize_match_text(record.get('product')),
        )
        for key in candidate_keys:
            if key and key not in order_lookup:
                order_lookup[key] = order_value

    return order_lookup


def format_sheet_name_with_total_order(sheet_name, order_lookup):
    if not sheet_name or not order_lookup:
        return sheet_name

    order_value = order_lookup.get(normalize_match_text(sheet_name))
    if order_value is None:
        order_value = order_lookup.get(normalize_material_fallback_key(sheet_name))
    if order_value is None:
        return sheet_name

    return f'{order_value}. {sheet_name}'

def build_source_record_maps(records):
    material_map = defaultdict(list)
    material_date_map = defaultdict(list)
    material_fallback_map = defaultdict(list)
    material_date_fallback_map = defaultdict(list)
    product_map = defaultdict(list)
    product_date_map = defaultdict(list)

    for record in records:
        material_key = record['material_key']
        material_fallback_key = record.get('material_fallback_key', '')
        date_key = record.get('date_key', '')
        product_key = normalize_match_text(record.get('product'))

        if material_key:
            material_map[material_key].append(record)
            if date_key:
                material_date_map[(material_key, date_key)].append(record)

        if material_fallback_key:
            material_fallback_map[material_fallback_key].append(record)
            if date_key:
                material_date_fallback_map[(material_fallback_key, date_key)].append(record)

        if product_key:
            product_map[product_key].append(record)
            if date_key:
                product_date_map[(product_key, date_key)].append(record)

    return {
        'material': material_map,
        'material_date': material_date_map,
        'material_fallback': material_fallback_map,
        'material_date_fallback': material_date_fallback_map,
        'product': product_map,
        'product_date': product_date_map,
    }


def take_matching_source_record(record_map, match_key, allow_reuse=False):
    if not match_key:
        return None

    records = record_map.get(match_key, [])
    for record in records:
        if not record['used']:
            record['used'] = True
            return record

    if allow_reuse and records:
        return records[0]

    return None



def take_unique_matching_source_record(record_map, match_key, allow_reuse=False):
    if not match_key:
        return None

    records = record_map.get(match_key, [])
    available_records = [record for record in records if not record['used']]
    if len(available_records) == 1:
        available_records[0]['used'] = True
        return available_records[0]

    if allow_reuse and len(records) == 1:
        return records[0]

    return None



def find_matching_source_record_for_row(
    material_map,
    material_date_map,
    material_fallback_map,
    material_date_fallback_map,
    product_map,
    product_date_map,
    material_key,
    material_fallback_key,
    product_key,
    date_key,
    prefer_product=False,
):
    record = None

    if prefer_product:
        if product_key and date_key:
            record = take_matching_source_record(product_date_map, (product_key, date_key), allow_reuse=True)
        if record is None and product_key:
            record = take_unique_matching_source_record(product_map, product_key, allow_reuse=True)

    if record is None and material_key and date_key:
        record = take_matching_source_record(material_date_map, (material_key, date_key), allow_reuse=True)
    if record is None and material_key:
        record = take_matching_source_record(material_map, material_key, allow_reuse=True)

    if record is None and material_fallback_key:
        if date_key:
            record = take_matching_source_record(material_date_fallback_map, (material_fallback_key, date_key), allow_reuse=True)
        if record is None:
            record = take_unique_matching_source_record(material_fallback_map, material_fallback_key, allow_reuse=True)

    if record is None and not prefer_product and product_key:
        if date_key:
            record = take_matching_source_record(product_date_map, (product_key, date_key), allow_reuse=True)
        if record is None:
            record = take_unique_matching_source_record(product_map, product_key, allow_reuse=True)

    return record


def to_excel_cell_value(value):
    if value in (None, ''):
        return None
    return value



def get_record_write_values(record):
    company_value = record['vendor'] if record['vendor'] not in (None, '') else record['company']
    material_value = record['material'] if record['material'] not in (None, '') else record['product']
    return {
        'recipe': to_excel_cell_value(record['recipe']),
        'date': to_excel_cell_value(record['date']),
        'company': to_excel_cell_value(company_value),
        'vendor': to_excel_cell_value(record['vendor']),
        'material': to_excel_cell_value(material_value),
        'judgment': to_excel_cell_value(record['judgment']),
        'spot': to_excel_cell_value(record['spot']),
    }


def clear_target_update_cells(sheet, row_index, layout):
    for key in ('measurement', 'recipe', 'date', 'company', 'vendor', 'judgment'):
        column = layout['columns'].get(key)
        if column:
            excel_set_cell_value(sheet, row_index, column, None)

    spot_column = layout.get('spot_column')
    if spot_column:
        excel_set_cell_value(sheet, row_index, spot_column, None)

    for element_columns in layout['elements'].values():
        excel_set_cell_value(sheet, row_index, element_columns['conc'], None)
        excel_set_cell_value(sheet, row_index, element_columns['err'], None)



def write_target_row_from_record(sheet, row_index, layout, record):
    write_values = get_record_write_values(record)

    recipe_column = layout['columns'].get('recipe')
    if recipe_column:
        excel_set_cell_value(sheet, row_index, recipe_column, write_values['recipe'])

    date_column = layout['columns'].get('date')
    if date_column:
        excel_set_cell_value(sheet, row_index, date_column, write_values['date'])

    company_column = layout['columns'].get('company')
    if company_column:
        excel_set_cell_value(sheet, row_index, company_column, write_values['company'])

    vendor_column = layout['columns'].get('vendor')
    if vendor_column:
        excel_set_cell_value(sheet, row_index, vendor_column, write_values['vendor'])

    material_column = layout['columns'].get('material')
    if material_column:
        excel_set_cell_value(sheet, row_index, material_column, write_values['material'])

    judgment_column = layout['columns'].get('judgment')
    if judgment_column:
        excel_set_cell_value(sheet, row_index, judgment_column, write_values['judgment'])

    spot_column = layout.get('spot_column')
    if spot_column:
        excel_set_cell_value(sheet, row_index, spot_column, write_values['spot'])

    for symbol_key, element_columns in layout['elements'].items():
        source_element = record['elements'].get(symbol_key, {})
        excel_set_cell_value(sheet, row_index, element_columns['conc'], to_excel_cell_value(source_element.get('conc')))
        excel_set_cell_value(sheet, row_index, element_columns['err'], to_excel_cell_value(source_element.get('err')))


def row_has_identity_data(sheet, row_index, layout):
    for key in ('recipe', 'product', 'material'):
        column = layout['columns'].get(key)
        if column and excel_cell_text(sheet, row_index, column):
            return True
    return False



def renumber_target_measurements(sheet, layout, last_row=None):
    measurement_column = layout['columns'].get('measurement')
    if not measurement_column:
        return 0

    if last_row is None:
        last_row = find_last_target_data_row(sheet, layout)

    next_number = 1
    for row_index in range(layout['data_start_row'], last_row + 1):
        if row_has_identity_data(sheet, row_index, layout):
            excel_set_cell_value(sheet, row_index, measurement_column, next_number)
            next_number += 1
        else:
            excel_set_cell_value(sheet, row_index, measurement_column, None)

    return next_number - 1



def find_last_target_data_row(sheet, layout):
    _, _, last_row, _ = get_sheet_bounds(sheet)
    last_data_row = layout['data_start_row'] - 1
    blank_streak = 0
    scan_limit = min(last_row, layout['data_start_row'] + 5000)

    for row_index in range(layout['data_start_row'], scan_limit + 1):
        if row_has_identity_data(sheet, row_index, layout):
            last_data_row = row_index
            blank_streak = 0
        else:
            blank_streak += 1
            if blank_streak >= 40 and last_data_row >= layout['data_start_row']:
                break

    return last_data_row



def apply_total_records_to_target_sheet(target_sheet, target_layout, source_records, prefer_product=False):
    record_maps = build_source_record_maps(source_records)
    material_map = record_maps['material']
    material_date_map = record_maps['material_date']
    material_fallback_map = record_maps['material_fallback']
    material_date_fallback_map = record_maps['material_date_fallback']
    product_map = record_maps['product']
    product_date_map = record_maps['product_date']
    last_row = find_last_target_data_row(target_sheet, target_layout)
    stats = {
        'candidate_rows': 0,
        'matched_rows': 0,
        'unmatched_rows': 0,
        'unused_source_rows': 0,
        'numbered_rows': 0,
    }

    for row_index in range(target_layout['data_start_row'], last_row + 1):
        date_value = excel_cell_display_or_value(target_sheet, row_index, target_layout['columns'].get('date'))
        product_value = excel_cell_text(target_sheet, row_index, target_layout['columns'].get('product'))
        material_value = excel_cell_text(target_sheet, row_index, target_layout['columns'].get('material'))

        if not any([product_value, material_value]):
            continue

        stats['candidate_rows'] += 1
        material_key = normalize_match_text(material_value)
        material_fallback_key = normalize_material_fallback_key(material_value)
        product_key = normalize_match_text(product_value)
        date_key = normalize_date_match_text(date_value)

        record = find_matching_source_record_for_row(
            material_map,
            material_date_map,
            material_fallback_map,
            material_date_fallback_map,
            product_map,
            product_date_map,
            material_key,
            material_fallback_key,
            product_key,
            date_key,
            prefer_product=prefer_product,
        )

        if record is None:
            clear_target_update_cells(target_sheet, row_index, target_layout)
            stats['unmatched_rows'] += 1
            continue

        write_target_row_from_record(target_sheet, row_index, target_layout, record)
        stats['matched_rows'] += 1

    stats['numbered_rows'] = renumber_target_measurements(target_sheet, target_layout, last_row=last_row)
    stats['unused_source_rows'] = sum(1 for record in source_records if not record['used'])
    return stats



def resolve_cached_target_layout_in_workbook(workbook, cached_layout):
    if not cached_layout:
        return None

    sheet_name = cached_layout.get('sheet_name')
    if not sheet_name:
        return None

    try:
        target_sheet = workbook.Worksheets(sheet_name)
    except Exception:
        return None

    layout = clone_layout_metadata(cached_layout)
    layout['sheet'] = target_sheet
    return layout


def refresh_target_cache_after_insert(path, target_workbook, sheet_names_to_delete):
    current_sheet_names = [sheet.Name for sheet in target_workbook.Worksheets]
    cached = get_cached_target_inspection(path)
    layout = None

    if cached:
        cached_layout = cached.get('layout')
        cached_sheet_name = cached_layout.get('sheet_name') if cached_layout else ''
        if cached_sheet_name and cached_sheet_name in current_sheet_names and cached_sheet_name not in set(sheet_names_to_delete):
            layout = cached_layout

    if layout is None:
        layout = detect_target_layout_in_workbook(target_workbook)

    set_cached_target_inspection(path, current_sheet_names, layout)
    return layout

def update_hazardous_tables_from_total(total_path, target_paths):
    ensure_excel_available()

    abs_total_path = os.path.abspath(total_path)
    normalized_target_paths = []
    seen_targets = set()
    for path in target_paths:
        abs_path = os.path.abspath(path)
        path_key = normalized_path(abs_path)
        if path_key in seen_targets:
            continue
        seen_targets.add(path_key)
        normalized_target_paths.append(abs_path)

    total_key = normalized_path(abs_total_path)
    if any(normalized_path(path) == total_key for path in normalized_target_paths):
        raise RuntimeError('업체 TOTAL 파일과 기존 저장 파일은 서로 다른 파일이어야 합니다.')

    excel = None
    total_workbook = None
    results = []
    failures = []
    try:
        excel = ensure_excel_application()
        total_layout = get_cached_total_layout(abs_total_path)
        source_records_template = peek_cached_total_records(abs_total_path)

        if total_layout is None or source_records_template is None:
            total_workbook = open_excel_workbook(excel, abs_total_path, read_only=True)
            total_layout = detect_total_layout_in_workbook(total_workbook)
            if not total_layout:
                raise RuntimeError('업체 TOTAL 파일에서 분석 데이터 시트를 찾지 못했습니다.')
            set_cached_total_layout(abs_total_path, total_layout)

            source_records_template = build_source_total_records(total_layout['sheet'], total_layout)
            if not source_records_template:
                raise RuntimeError('업체 TOTAL 파일에서 업데이트할 데이터가 없습니다.')
            set_cached_total_records(abs_total_path, source_records_template)
        else:
            total_layout = clone_layout_metadata(total_layout)

        for target_path in normalized_target_paths:
            target_workbook = None
            save_changes = False
            try:
                target_workbook = open_excel_workbook(excel, target_path, read_only=False)
                target_layout = detect_target_layout_in_workbook(target_workbook)
                if not target_layout:
                    raise RuntimeError('기존 저장 파일에서 유해물질분석표 시트를 찾지 못했습니다.')

                source_records = deepcopy(source_records_template)
                stats = apply_total_records_to_target_sheet(
                    target_layout['sheet'],
                    target_layout,
                    source_records,
                    prefer_product=is_styler_xrf_path(target_path),
                )
                target_workbook.Save()
                set_cached_target_inspection(
                    target_path,
                    [sheet.Name for sheet in target_workbook.Worksheets],
                    target_layout,
                )
                save_changes = True
                stats['target_sheet_name'] = target_layout['sheet_name']
                stats['source_sheet_name'] = total_layout['sheet_name']
                results.append({'path': target_path, 'stats': stats})
            except RuntimeError as exc:
                failures.append({'path': target_path, 'error': str(exc)})
            except com_error:
                failures.append({'path': target_path, 'error': '유해물질표 업데이트 중 엑셀 자동화 오류가 발생했습니다.'})
            finally:
                if target_workbook is not None:
                    target_workbook.Close(save_changes)

        return results, failures
    except com_error as exc:
        raise RuntimeError('유해물질표 업데이트 중 엑셀 자동화 오류가 발생했습니다.') from exc
    finally:
        if total_workbook is not None:
            total_workbook.Close(False)


def update_hazardous_table_from_total(total_path, target_path):
    results, failures = update_hazardous_tables_from_total(total_path, [target_path])
    if failures:
        raise RuntimeError(failures[0]['error'])
    if not results:
        raise RuntimeError('유해물질표 업데이트에 실패했습니다.')
    return results[0]['stats']


def insert_sheets_into_open_target_workbook(source_workbooks, target_workbook, sheet_names_to_delete):
    existing_sheet_names = [sheet.Name for sheet in target_workbook.Worksheets]
    existing_sheet_name_set = set(existing_sheet_names)
    valid_delete_names = [name for name in sheet_names_to_delete if name in existing_sheet_name_set]

    placeholder_name = None
    if valid_delete_names and len(valid_delete_names) >= target_workbook.Worksheets.Count:
        placeholder_name = make_placeholder_sheet_name(existing_sheet_name_set)
        placeholder_sheet = target_workbook.Worksheets.Add(Before=target_workbook.Worksheets(1))
        placeholder_sheet.Name = placeholder_name

    for sheet_name in valid_delete_names:
        target_workbook.Worksheets(sheet_name).Delete()

    inserted_sheet_count = 0
    for source_workbook in source_workbooks:
        sheet_count = source_workbook.Worksheets.Count
        for sheet_index in range(1, sheet_count + 1):
            source_workbook.Worksheets(sheet_index).Copy(
                None,
                target_workbook.Worksheets(target_workbook.Worksheets.Count),
            )
            inserted_sheet_count += 1

    if placeholder_name and target_workbook.Worksheets.Count > 1:
        target_workbook.Worksheets(placeholder_name).Delete()

    return {
        'deleted_sheet_count': len(valid_delete_names),
        'inserted_sheet_count': inserted_sheet_count,
        'final_sheet_count': target_workbook.Worksheets.Count,
    }


def insert_sheets_into_target_workbooks(source_paths, target_paths, sheet_names_to_delete):
    ensure_excel_available()

    normalized_source_paths = []
    seen_sources = set()
    for path in source_paths:
        abs_path = os.path.abspath(path)
        path_key = normalized_path(abs_path)
        if path_key in seen_sources:
            continue
        seen_sources.add(path_key)
        normalized_source_paths.append(abs_path)

    normalized_target_paths = []
    seen_targets = set()
    for path in target_paths:
        abs_path = os.path.abspath(path)
        path_key = normalized_path(abs_path)
        if path_key in seen_targets:
            continue
        seen_targets.add(path_key)
        normalized_target_paths.append(abs_path)

    target_keys = {normalized_path(path) for path in normalized_target_paths}
    if any(normalized_path(path) in target_keys for path in normalized_source_paths):
        raise RuntimeError('기존 저장 파일은 삽입할 엑셀 목록에 포함할 수 없습니다.')

    excel = None
    source_workbooks = []
    results = []
    failures = []
    try:
        excel = ensure_excel_application()

        for source_path in normalized_source_paths:
            try:
                source_workbooks.append(open_excel_workbook(excel, source_path, read_only=True))
            except com_error as exc:
                raise RuntimeError('삽입할 엑셀 파일을 읽는 중 오류가 발생했습니다.') from exc

        for target_path in normalized_target_paths:
            target_workbook = None
            save_changes = False
            try:
                target_workbook = open_excel_workbook(excel, target_path, read_only=False)
                stats = insert_sheets_into_open_target_workbook(source_workbooks, target_workbook, sheet_names_to_delete)
                target_workbook.Save()
                set_cached_target_inspection(
                    target_path,
                    [sheet.Name for sheet in target_workbook.Worksheets],
                    detect_target_layout_in_workbook(target_workbook),
                )
                save_changes = True
                results.append({'path': target_path, 'stats': stats})
            except com_error:
                failures.append({'path': target_path, 'error': '시트 삽입 중 엑셀 자동화 오류가 발생했습니다.'})
            finally:
                if target_workbook is not None:
                    target_workbook.Close(save_changes)
    finally:
        for source_workbook in reversed(source_workbooks):
            try:
                source_workbook.Close(False)
            except Exception:
                pass

    return results, failures


def insert_sheets_into_target_workbook(source_paths, target_path, sheet_names_to_delete):
    results, failures = insert_sheets_into_target_workbooks(source_paths, [target_path], sheet_names_to_delete)
    if failures:
        raise RuntimeError(failures[0]['error'])
    if not results:
        raise RuntimeError('시트 삽입에 실패했습니다.')
    return results[0]['stats']



class SelectableList(tk.Frame):
    def __init__(self, master, width=320, height=420, bg='#cccccc'):
        super().__init__(master, bg=bg, bd=0, highlightthickness=0)
        self._bg = bg
        self._items = []
        self._row_frames = []
        self._view_width = width
        self._view_height = height
        self._check_font_size = 16
        self._text_font_size = 12
        self._text_wraplength = max(180, int(width) - 96)

        self.canvas = tk.Canvas(
            self,
            width=width,
            height=height,
            bg=bg,
            highlightthickness=0,
            bd=0,
        )
        self.scrollbar = tk.Scrollbar(self, orient='vertical', command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.inner = tk.Frame(self.canvas, bg=bg)
        self.window_id = self.canvas.create_window((0, 0), window=self.inner, anchor='nw')

        self.canvas.pack(side='left', fill='both', expand=True)
        self.scrollbar.pack(side='right', fill='y')

        self.inner.bind('<Configure>', self._on_inner_configure)
        self.canvas.bind('<Configure>', self._on_canvas_configure)
        self._bind_scroll_target(self)
        self._bind_scroll_target(self.canvas)
        self._bind_scroll_target(self.inner)
        self._bind_scroll_target(self.scrollbar)

    def _bind_scroll_target(self, widget):
        widget.bind('<MouseWheel>', self._on_mousewheel, add='+')
        widget.bind('<Button-4>', self._on_mousewheel_linux, add='+')
        widget.bind('<Button-5>', self._on_mousewheel_linux, add='+')

    def _on_inner_configure(self, _event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox('all'))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfigure(self.window_id, width=event.width)

    def _pointer_inside_list(self, event):
        widget = self.winfo_containing(event.x_root, event.y_root)
        while widget is not None:
            if widget is self:
                return True
            widget = widget.master
        return False

    def _can_scroll(self):
        first, last = self.canvas.yview()
        return not (first <= 0 and last >= 1)

    def _on_mousewheel(self, event):
        if not self._pointer_inside_list(event):
            return
        if not self._can_scroll():
            return 'break'
        step = -1 if event.delta > 0 else 1
        self.canvas.yview_scroll(step, 'units')
        return 'break'

    def _on_mousewheel_linux(self, event):
        if not self._pointer_inside_list(event):
            return
        if not self._can_scroll():
            return 'break'
        step = -1 if getattr(event, 'num', 0) == 4 else 1
        self.canvas.yview_scroll(step, 'units')
        return 'break'

    def clear(self):
        for frame in self._row_frames:
            frame.destroy()
        self._items.clear()
        self._row_frames.clear()

    def set_items(self, items):
        self.clear()
        for item in items:
            if isinstance(item, dict):
                label = item['label']
                value = item.get('value', label)
                selected = bool(item.get('selected', False))
            elif isinstance(item, (tuple, list)) and len(item) >= 2:
                label, value = item[0], item[1]
                selected = False
            else:
                label = str(item)
                value = item
                selected = False
            self._items.append({'label': label, 'value': value, 'selected': selected})
        self._rebuild_rows()

    def _rebuild_rows(self):
        for index, item in enumerate(self._items):
            row = tk.Frame(self.inner, bg=self._bg, padx=8, pady=4)
            row.pack(fill='x', anchor='w')

            check_label = tk.Label(
                row,
                text='☑' if item['selected'] else '☐',
                fg='#d11f1f' if item['selected'] else '#222222',
                bg=self._bg,
                font=('\ub9d1\uc740 \uace0\ub515', self._check_font_size, 'bold'),
                width=2,
                anchor='w',
            )
            text_label = tk.Label(
                row,
                text=item['label'],
                fg='#d11f1f' if item['selected'] else '#ffffff',
                bg=self._bg,
                font=('\ub9d1\uc740 \uace0\ub515', self._text_font_size, 'bold'),
                anchor='w',
                justify='left',
                wraplength=self._text_wraplength,
            )
            check_label.pack(side='left', anchor='w')
            text_label.pack(side='left', fill='x', expand=True, anchor='w')

            self._bind_scroll_target(row)
            self._bind_scroll_target(check_label)
            self._bind_scroll_target(text_label)
            row.bind('<Button-1>', lambda _event, i=index: self.toggle(i))
            check_label.bind('<Button-1>', lambda _event, i=index: self.toggle(i))
            text_label.bind('<Button-1>', lambda _event, i=index: self.toggle(i))
            self._row_frames.append(row)

    def toggle(self, index):
        self._items[index]['selected'] = not self._items[index]['selected']
        row = self._row_frames[index]
        check_label = row.winfo_children()[0]
        text_label = row.winfo_children()[1]
        selected = self._items[index]['selected']
        check_label.configure(text='☑' if selected else '☐', fg='#d11f1f' if selected else '#222222')
        text_label.configure(fg='#d11f1f' if selected else '#ffffff')

    def get_selected_values(self):
        return [item['value'] for item in self._items if item['selected']]

    def set_selected_values(self, values):
        target_values = set(values)
        for index, item in enumerate(self._items):
            item['selected'] = item['value'] in target_values
            if index < len(self._row_frames):
                row = self._row_frames[index]
                check_label = row.winfo_children()[0]
                text_label = row.winfo_children()[1]
                selected = item['selected']
                check_label.configure(text='☑' if selected else '☐', fg='#d11f1f' if selected else '#222222')
                text_label.configure(fg='#d11f1f' if selected else '#ffffff')

    def select_all(self):
        self.set_selected_values([item['value'] for item in self._items])

    def clear_selection(self):
        self.set_selected_values([])

    def has_items(self):
        return bool(self._items)

    def configure_view(self, width=None, height=None, check_font_size=None, text_font_size=None, wraplength=None):
        if width is not None:
            self._view_width = int(width)
            self.canvas.configure(width=self._view_width)
        if height is not None:
            self._view_height = int(height)
            self.canvas.configure(height=self._view_height)
        if check_font_size is not None:
            self._check_font_size = int(check_font_size)
        if text_font_size is not None:
            self._text_font_size = int(text_font_size)
        if wraplength is not None:
            self._text_wraplength = int(wraplength)
        self._refresh_row_styles()

    def _refresh_row_styles(self):
        for row in self._row_frames:
            children = row.winfo_children()
            if len(children) < 2:
                continue
            check_label = children[0]
            text_label = children[1]
            check_label.configure(font=('\ub9d1\uc740 \uace0\ub515', self._check_font_size, 'bold'))
            text_label.configure(font=('\ub9d1\uc740 \uace0\ub515', self._text_font_size, 'bold'), wraplength=self._text_wraplength)


class ScrollableInfoBox(tk.Frame):
    def __init__(self, master, width=760, height=120, bg='white', fg='#303030', initial_text=''):
        super().__init__(master, bg=bg, relief='solid', bd=1)
        self._font_size = 11
        self.text = tk.Text(
            self,
            bg=bg,
            fg=fg,
            font=('맑은 고딕', 11, 'bold'),
            relief='flat',
            bd=0,
            highlightthickness=0,
            wrap='word',
            padx=12,
            pady=12,
            cursor='arrow',
        )
        self.scrollbar = tk.Scrollbar(self, orient='vertical', command=self.text.yview)
        self.text.configure(yscrollcommand=self.scrollbar.set)

        self.configure(width=width, height=height)
        self.pack_propagate(False)

        self.text.pack(side='left', fill='both', expand=True)
        self.scrollbar.pack(side='right', fill='y')

        self._bind_scroll_target(self)
        self._bind_scroll_target(self.text)
        self._bind_scroll_target(self.scrollbar)
        self.set_text(initial_text)

    def _bind_scroll_target(self, widget):
        widget.bind('<MouseWheel>', self._on_mousewheel, add='+')
        widget.bind('<Button-4>', self._on_mousewheel_linux, add='+')
        widget.bind('<Button-5>', self._on_mousewheel_linux, add='+')

    def _can_scroll(self):
        first, last = self.text.yview()
        return not (first <= 0 and last >= 1)

    def _on_mousewheel(self, event):
        if not self._can_scroll():
            return 'break'
        step = -1 if event.delta > 0 else 1
        self.text.yview_scroll(step, 'units')
        return 'break'

    def _on_mousewheel_linux(self, event):
        if not self._can_scroll():
            return 'break'
        step = -1 if getattr(event, 'num', 0) == 4 else 1
        self.text.yview_scroll(step, 'units')
        return 'break'

    def set_text(self, text):
        self.text.configure(state='normal')
        self.text.delete('1.0', 'end')
        self.text.insert('1.0', text or '')
        self.text.configure(state='disabled')
        self.text.yview_moveto(0)

    def configure_view(self, width=None, height=None, font_size=None):
        config_kwargs = {}
        if width is not None:
            config_kwargs['width'] = int(width)
        if height is not None:
            config_kwargs['height'] = int(height)
        if config_kwargs:
            self.configure(**config_kwargs)
        if font_size is not None:
            self._font_size = int(font_size)
            self.text.configure(font=('\ub9d1\uc740 \uace0\ub515', self._font_size, 'bold'))

class XRFReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title('XRF Report Auto Input System')
        self.root.geometry(f'{DEFAULT_WINDOW_SIZE[0]}x{DEFAULT_WINDOW_SIZE[1]}')
        self.root.minsize(*MIN_WINDOW_SIZE)
        self.root.configure(bg='#b4b4b4')

        self.source_paths = []
        self.target_workbook_paths = []
        self.target_workbook_path = ''
        self.total_source_path = ''
        self.target_sheet_names = []
        self.detected_target_sheet_name = ''
        self.detected_total_sheet_name = ''
        self.total_sheet_order_lookup = {}
        self.target_load_request_id = 0
        self.total_load_request_id = 0
        self.target_inspection_pending = False
        self.total_inspection_pending = False

        self.target_info_message = '기존 저장 파일 미선택'
        self.total_info_var = tk.StringVar(value='업체 TOTAL 파일 미선택')
        self.status_var = tk.StringVar(value='기존 저장 파일과 업체 TOTAL 파일을 불러오세요.')
        self.result_var = tk.StringVar(value='READY')
        self.optimize_var = tk.BooleanVar(value=False)
        self.about_window = None
        self.about_logo_image = None
        self.version_history_window = None
        self.fullscreen_mode = False
        self.default_tk_scaling = float(self.root.tk.call('tk', 'scaling'))
        self.ui_scale_factor = 1.0
        self.window_resolution = DEFAULT_WINDOW_SIZE
        self.current_min_window_size = DEFAULT_MIN_WINDOW_SIZE
        self.ui_scale_var = tk.StringVar(value=self.get_scale_menu_value(1.0))
        self.resolution_var = tk.StringVar(value=self.get_resolution_menu_value(DEFAULT_WINDOW_SIZE))
        self.windowed_geometry = self.root.geometry()
        self._responsive_after_id = None
        self._middle_layout_mode = None
        self.update_check_in_progress = False
        self.update_download_in_progress = False
        self.load_display_settings()

        self._build_ui()
        self._build_menu()
        self.apply_display_settings(center=True)
        self.root.bind('<F11>', self.handle_toggle_fullscreen)
        self.root.bind('<Escape>', self.handle_exit_fullscreen)
        self.root.bind('<Configure>', self._on_root_configure, add='+')
        self.root.protocol('WM_DELETE_WINDOW', self.on_close)
        self.root.after(2500, lambda: self.check_for_updates(silent=True))

    def on_close(self):
        self.save_display_settings()
        shutdown_excel_application()
        self.root.destroy()

    def load_display_settings(self):
        settings_path = display_settings_path()
        try:
            with open(settings_path, 'r', encoding='utf-8') as file:
                data = json.load(file)
        except (OSError, ValueError, TypeError):
            return

        scale_value = data.get('ui_scale_factor')
        if isinstance(scale_value, (int, float)) and UI_SCALE_MIN <= float(scale_value) <= UI_SCALE_MAX:
            self.ui_scale_factor = float(scale_value)

        resolution_value = data.get('window_resolution')
        if isinstance(resolution_value, (list, tuple)) and len(resolution_value) == 2:
            try:
                width = int(resolution_value[0])
                height = int(resolution_value[1])
            except (TypeError, ValueError):
                width = height = 0
            if width > 0 and height > 0:
                self.window_resolution = (width, height)

        self.ui_scale_var.set(self.get_scale_menu_value())
        self.resolution_var.set(self.get_resolution_menu_value())


    def save_display_settings(self):
        settings_path = display_settings_path()
        try:
            os.makedirs(os.path.dirname(settings_path), exist_ok=True)
            with open(settings_path, 'w', encoding='utf-8') as file:
                json.dump(
                    {
                        'ui_scale_factor': self.ui_scale_factor,
                        'window_resolution': [int(self.window_resolution[0]), int(self.window_resolution[1])],
                    },
                    file,
                    ensure_ascii=False,
                    indent=2,
                )
        except OSError:
            pass

    def get_scale_menu_value(self, factor=None):
        target_factor = self.ui_scale_factor if factor is None else float(factor)
        return f'{int(round(target_factor * 100))}%'

    def get_resolution_menu_value(self, size=None):
        width, height = self.window_resolution if size is None else size
        return f'{int(width)}x{int(height)}'

    def get_screen_work_area(self):
        try:
            rect = ctypes.wintypes.RECT()
            if ctypes.windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(rect), 0):
                return rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top
        except Exception:
            pass
        return 0, 0, self.root.winfo_screenwidth(), self.root.winfo_screenheight()

    def get_available_window_area(self):
        left, top, width, height = self.get_screen_work_area()
        pad_x, pad_y = SCREEN_FIT_PADDING
        available_width = max(1, int(width) - pad_x)
        available_height = max(1, int(height) - pad_y)
        return left + pad_x // 2, top + pad_y // 2, available_width, available_height

    def _apply_tk_scaling(self, scale_factor):
        bounded_scale = max(UI_SCALE_MIN, min(UI_SCALE_MAX, float(scale_factor)))
        try:
            self.root.tk.call('tk', 'scaling', self.default_tk_scaling * bounded_scale)
        except tk.TclError:
            pass
        return bounded_scale

    def get_effective_ui_scale_factor(self):
        _, _, available_width, available_height = self.get_available_window_area()
        reference_width = max(DEFAULT_MIN_WINDOW_SIZE[0], int(self.window_resolution[0]))
        reference_height = max(DEFAULT_MIN_WINDOW_SIZE[1], int(self.window_resolution[1]))
        fit_width = available_width / max(reference_width, 1)
        fit_height = available_height / max(reference_height, 1)
        fit_factor = min(self.ui_scale_factor, fit_width, fit_height, 1.0)
        return max(UI_SCALE_MIN, min(UI_SCALE_MAX, fit_factor))

    def _parse_window_position(self, geometry_text):
        match = re.match(r'\d+x\d+([+-]\d+)([+-]\d+)', geometry_text or '')
        if not match:
            return None, None
        return int(match.group(1)), int(match.group(2))

    def _center_window_position(self, width, height):
        area_left, area_top, available_width, available_height = self.get_available_window_area()
        x = max(area_left + (available_width - width) // 2, area_left)
        y = max(area_top + (available_height - height) // 2, area_top)
        return x, y

    def apply_display_settings(self, center=False):
        effective_scale_factor = self._apply_tk_scaling(self.get_effective_ui_scale_factor())

        self.root.update_idletasks()
        area_left, area_top, available_width, available_height = self.get_available_window_area()
        actual_required_width = max(MIN_WINDOW_SIZE[0], self.root.winfo_reqwidth())
        actual_required_height = max(MIN_WINDOW_SIZE[1], self.root.winfo_reqheight())

        if actual_required_width > available_width or actual_required_height > available_height:
            extra_fit_ratio = min(
                available_width / max(actual_required_width, 1),
                available_height / max(actual_required_height, 1),
                1.0,
            )
            adjusted_scale_factor = max(
                UI_SCALE_MIN,
                min(UI_SCALE_MAX, effective_scale_factor * extra_fit_ratio),
            )
            if adjusted_scale_factor < effective_scale_factor - 0.01:
                effective_scale_factor = self._apply_tk_scaling(adjusted_scale_factor)
                self.root.update_idletasks()

        scaled_min_width = int(round(DEFAULT_MIN_WINDOW_SIZE[0] * effective_scale_factor))
        scaled_min_height = int(round(DEFAULT_MIN_WINDOW_SIZE[1] * effective_scale_factor))
        required_width = min(
            max(MIN_WINDOW_SIZE[0], scaled_min_width, self.root.winfo_reqwidth()),
            available_width,
        )
        required_height = min(
            max(MIN_WINDOW_SIZE[1], scaled_min_height, self.root.winfo_reqheight()),
            available_height,
        )
        self.current_min_window_size = (required_width, required_height)
        self.root.minsize(required_width, required_height)

        desired_width = min(max(int(self.window_resolution[0]), required_width), available_width)
        desired_height = min(max(int(self.window_resolution[1]), required_height), available_height)
        x, y = self._parse_window_position(self.windowed_geometry)
        if x is None or y is None or center:
            x, y = self._center_window_position(desired_width, desired_height)
        else:
            max_x = area_left + max(available_width - desired_width, 0)
            max_y = area_top + max(available_height - desired_height, 0)
            x = min(max(x, area_left), max_x)
            y = min(max(y, area_top), max_y)
        target_geometry = f'{desired_width}x{desired_height}+{x}+{y}'

        if self.fullscreen_mode:
            self.windowed_geometry = target_geometry
        else:
            self.root.geometry(target_geometry)
            self.root.update_idletasks()
            self.windowed_geometry = self.root.geometry()

        self.ui_scale_var.set(self.get_scale_menu_value())
        self.resolution_var.set(self.get_resolution_menu_value())
        self.refresh_responsive_layout()
    def apply_window_layout_constraints(self):
        self.apply_display_settings()

    def _on_root_configure(self, event=None):
        if event is not None and event.widget is not self.root:
            return
        if self._responsive_after_id is not None:
            try:
                self.root.after_cancel(self._responsive_after_id)
            except tk.TclError:
                pass
        self._responsive_after_id = self.root.after(80, self.refresh_responsive_layout)

    def _set_middle_layout_mode(self, mode):
        if self._middle_layout_mode == mode:
            return
        self.delete_panel.pack_forget()
        self.result_panel.pack_forget()
        if mode == 'stacked':
            self.delete_panel.pack(side='top', fill='both', expand=True, pady=(0, 14))
            self.result_panel.pack(side='top', fill='x', expand=False)
        else:
            self.delete_panel.pack(side='left', fill='both', expand=True, padx=(0, 14))
            self.result_panel.pack(side='left', fill='both', expand=True)
        self._middle_layout_mode = mode

    def refresh_responsive_layout(self):
        self._responsive_after_id = None
        if not self.root.winfo_exists():
            return

        self.root.update_idletasks()
        window_width = max(self.root.winfo_width(), MIN_WINDOW_SIZE[0])
        window_height = max(self.root.winfo_height(), MIN_WINDOW_SIZE[1])
        compact = window_width < 1440 or window_height < 860
        tight = window_width < 1280 or window_height < 760

        self.title_label.configure(
            font=('\ub9d1\uc740 \uace0\ub515', 18 if compact else 22, 'bold'),
            padx=18 if compact else 24,
            pady=8 if compact else 10,
        )

        list_text_size = 10 if tight else 12
        list_check_size = 14 if tight else 16
        source_list_width = 280 if tight else (320 if compact else 360)
        source_list_height = min(560, max(320, window_height - (320 if compact else 290)))
        self.source_list.configure_view(
            width=source_list_width,
            height=source_list_height,
            check_font_size=list_check_size,
            text_font_size=list_text_size,
            wraplength=max(160, source_list_width - 92),
        )

        layout_mode = 'stacked' if (window_width < 1380 or window_height < 760) else 'side_by_side'
        self._set_middle_layout_mode(layout_mode)

        right_width = max(420, window_width - source_list_width - 110)
        info_width = max(380, right_width - 32)
        self.target_info_box.configure_view(
            width=info_width,
            height=96 if compact else 122,
            font_size=10 if compact else 11,
        )
        self.total_info_label.configure(
            font=('\ub9d1\uc740 \uace0\ub515', 10 if compact else 11, 'bold'),
            wraplength=max(320, info_width - 48),
        )

        delete_list_width = max(340, int(right_width * (0.92 if layout_mode == 'stacked' else 0.54)))
        delete_list_height = 220 if tight else (260 if compact else 320)
        self.target_sheet_list.configure_view(
            width=delete_list_width,
            height=delete_list_height,
            check_font_size=list_check_size,
            text_font_size=list_text_size,
            wraplength=max(180, delete_list_width - 92),
        )
        self.delete_header_label.configure(
            font=('\ub9d1\uc740 \uace0\ub515', 12 if compact else 14, 'bold'),
            wraplength=max(320, delete_list_width - 8),
        )

        self.result_label.configure(
            font=('\ub9d1\uc740 \uace0\ub515', 52 if tight else (60 if compact else 72), 'bold'),
            pady=22 if compact else 40,
        )
        status_wrap = max(260, info_width - 48) if layout_mode == 'stacked' else max(240, int(right_width * 0.28))
        self.status_label.configure(
            font=('\ub9d1\uc740 \uace0\ub515', 11 if compact else 13, 'bold'),
            wraplength=status_wrap,
            padx=14 if compact else 18,
            pady=12 if compact else 16,
        )

    def set_ui_scale(self, factor):
        self.ui_scale_factor = float(factor)
        self.apply_display_settings(center=not self.fullscreen_mode)
        self.save_display_settings()
        self.status_var.set(f'화면 배율을 {self.get_scale_menu_value()}로 변경했습니다.')

    def set_window_resolution(self, width, height):
        self.window_resolution = (int(width), int(height))
        self.apply_display_settings(center=not self.fullscreen_mode)
        self.save_display_settings()
        self.status_var.set(f'해상도를 {width} x {height}로 변경했습니다.')

    def reset_display_settings(self):
        self.ui_scale_factor = 1.0
        self.window_resolution = DEFAULT_WINDOW_SIZE
        self.apply_display_settings(center=True)
        self.save_display_settings()
        self.status_var.set('화면 설정을 기본값으로 초기화했습니다.')
    def handle_toggle_fullscreen(self, _event=None):
        self.set_fullscreen_mode(not self.fullscreen_mode)
        return 'break'

    def handle_exit_fullscreen(self, _event=None):
        if self.fullscreen_mode:
            self.set_fullscreen_mode(False)
            return 'break'
        return None

    def enter_fullscreen_mode(self):
        self.set_fullscreen_mode(True)

    def exit_fullscreen_mode(self):
        self.set_fullscreen_mode(False)

    def set_fullscreen_mode(self, enabled):
        enabled = bool(enabled)
        current_fullscreen = bool(self.root.attributes('-fullscreen'))
        if enabled == self.fullscreen_mode and enabled == current_fullscreen:
            return

        if enabled:
            self.windowed_geometry = self.root.geometry()
            self.root.attributes('-fullscreen', True)
            self.fullscreen_mode = True
            self.status_var.set('전체 모드로 전환되었습니다. Esc 키로 창 모드로 돌아갈 수 있습니다.')
        else:
            self.root.attributes('-fullscreen', False)
            self.root.state('normal')
            self.root.geometry(self.windowed_geometry)
            self.root.update_idletasks()
            self.fullscreen_mode = False
            self.status_var.set('창 모드로 전환되었습니다. F11 키로 전체 모드로 전환할 수 있습니다.')


    def _build_menu(self):

        menu_bar = tk.Menu(self.root)

        view_menu = tk.Menu(menu_bar, tearoff=0)
        view_menu.add_command(label='전체 모드', command=self.enter_fullscreen_mode)
        view_menu.add_command(label='창 모드', command=self.exit_fullscreen_mode)
        view_menu.add_separator()

        scale_menu = tk.Menu(view_menu, tearoff=0)
        for label, factor in UI_SCALE_PRESETS:
            scale_menu.add_radiobutton(
                label=label,
                value=self.get_scale_menu_value(factor),
                variable=self.ui_scale_var,
                command=lambda selected_factor=factor: self.set_ui_scale(selected_factor),
            )

        resolution_menu = tk.Menu(view_menu, tearoff=0)
        for label, size in RESOLUTION_PRESETS:
            resolution_menu.add_radiobutton(
                label=label,
                value=self.get_resolution_menu_value(size),
                variable=self.resolution_var,
                command=lambda selected_size=size: self.set_window_resolution(*selected_size),
            )

        view_menu.add_cascade(label='화면 배율', menu=scale_menu)
        view_menu.add_cascade(label='해상도 설정', menu=resolution_menu)
        view_menu.add_separator()
        view_menu.add_command(label='화면 설정 초기화', command=self.reset_display_settings)
        menu_bar.add_cascade(label='화면 모드', menu=view_menu)

        help_menu = tk.Menu(menu_bar, tearoff=0)
        help_menu.add_command(label='사용가이드', command=self.show_help_usage_guide)
        help_menu.add_separator()
        help_menu.add_command(label='문제 해결', command=self.show_help_troubleshooting)
        help_menu.add_command(label='\uc5c5\ub370\uc774\ud2b8 \ud655\uc778', command=lambda: self.check_for_updates(silent=False))
        help_menu.add_command(label='프로그램 폴더 열기', command=self.open_program_folder)
        help_menu.add_separator()
        help_menu.add_command(label='Keyboard Shortcuts', command=self.show_help_shortcuts)
        help_menu.add_command(label='버전 정보', command=self.show_help_about)
        menu_bar.add_cascade(label='Help', menu=help_menu)
        self.root.configure(menu=menu_bar)
        self.menu_bar = menu_bar

    def show_help_usage_guide(self):

        messagebox.showinfo('사용가이드', '\n'.join([
            '[기본 사용 순서]',
            '1. 기존 저장 파일을 불러옵니다.',
            '2. 업체 TOTAL 파일을 불러옵니다.',
            '3. 재질파일불러오기 후 필요한 재질을 체크합니다.',
            '4. 시트 삽입 실행 또는 유해물질표 업데이트를 실행합니다.',
            '',
            '[시트 삽입 안내]',
            '- 왼쪽 재질 목록에서 삽입할 파일을 체크합니다.',
            '- 기존 저장 파일을 먼저 불러옵니다.',
            '- 삭제할 시트를 체크합니다.',
            '- 시트 삽입 실행을 누르면 선택한 재질 시트가 들어갑니다.',
            '',
            '[유해물질표 업데이트 안내]',
            '- 기존 저장 파일과 업체 TOTAL 파일을 불러옵니다.',
            '- 유해물질표 업데이트를 누르면 자동으로 매칭합니다.',
            '- 일반 파일은 재질 우선, Styler 파일은 품명 우선으로 매칭합니다.',
            '- 품명은 기존 값 유지, 업체명은 AZENTEK NO. 기준으로 들어갑니다.',
        ]))

    def show_help_troubleshooting(self):
        messagebox.showinfo('문제 해결', '\n'.join([
            '- 작업할 엑셀 파일은 열려 있지 않게 해주세요.',
            '- 기존 저장 파일과 업체 TOTAL 파일은 서로 다른 파일이어야 합니다.',
            '- 파일 경로가 바뀌면 다시 불러오면 됩니다.',
            '- TOTAL 데이터 시트가 안 잡히면 헤더명과 시트 구성을 확인해 주세요.',
        ]))

    def show_help_shortcuts(self):
        messagebox.showinfo('Keyboard Shortcuts', '\n'.join([
            'F11 : 전체 모드 전환',
            'Esc : 전체 모드에서 창 모드로 복귀',
            '목록은 마우스 휠로 스크롤할 수 있습니다.',
        ]))


    def check_for_updates(self, silent=False):
        if self.update_check_in_progress:
            if not silent:
                messagebox.showinfo('\uc5c5\ub370\uc774\ud2b8 \ud655\uc778', '\uc774\ubbf8 \uc5c5\ub370\uc774\ud2b8\ub97c \ud655\uc778\ud558\ub294 \uc911\uc785\ub2c8\ub2e4.')
            return

        self.update_check_in_progress = True
        if not silent:
            self.status_var.set('GitHub\uc5d0\uc11c \ucd5c\uc2e0 \ubc84\uc804\uc744 \ud655\uc778\ud558\ub294 \uc911\uc785\ub2c8\ub2e4.')

        thread = threading.Thread(target=self._check_for_updates_worker, args=(silent,), daemon=True)
        thread.start()

    def _check_for_updates_worker(self, silent):
        try:
            release_info = fetch_latest_release_info()
        except Exception as exc:
            self._schedule_ui(lambda: self._handle_update_check_failed(exc, silent))
            return
        self._schedule_ui(lambda: self._handle_update_check_result(release_info, silent))

    def _schedule_ui(self, callback):
        try:
            self.root.after(0, callback)
        except tk.TclError:
            pass

    def _handle_update_check_failed(self, error, silent):
        self.update_check_in_progress = False
        if silent:
            return
        messagebox.showerror(
            '\uc5c5\ub370\uc774\ud2b8 \ud655\uc778 \uc2e4\ud328',
            '\n'.join([
                'GitHub \uc5c5\ub370\uc774\ud2b8 \uc815\ubcf4\ub97c \ubd88\ub7ec\uc624\uc9c0 \ubabb\ud588\uc2b5\ub2c8\ub2e4.',
                '',
                str(error),
            ]),
        )
        self.status_var.set('\uc5c5\ub370\uc774\ud2b8 \ud655\uc778\uc5d0 \uc2e4\ud328\ud588\uc2b5\ub2c8\ub2e4.')

    def _handle_update_check_result(self, release_info, silent):
        self.update_check_in_progress = False
        latest_version = release_info.get('version') or ''
        if not latest_version:
            if not silent:
                messagebox.showwarning('\uc5c5\ub370\uc774\ud2b8 \ud655\uc778', 'GitHub \ub9b4\ub9ac\uc2a4\uc5d0\uc11c \ubc84\uc804 \uc815\ubcf4\ub97c \ucc3e\uc9c0 \ubabb\ud588\uc2b5\ub2c8\ub2e4.')
            return

        if not is_newer_version(latest_version, APP_VERSION):
            if not silent:
                messagebox.showinfo(
                    '\uc5c5\ub370\uc774\ud2b8 \ud655\uc778',
                    f'\ud604\uc7ac \ucd5c\uc2e0 \ubc84\uc804\uc785\ub2c8\ub2e4.\n\n\ud604\uc7ac \ubc84\uc804: {APP_VERSION}\nGitHub \ucd5c\uc2e0 \ubc84\uc804: {latest_version}',
                )
                self.status_var.set('\ud604\uc7ac \ucd5c\uc2e0 \ubc84\uc804\uc785\ub2c8\ub2e4.')
            return

        if not release_info.get('asset_url'):
            messagebox.showwarning(
                '\uc5c5\ub370\uc774\ud2b8 \ud30c\uc77c \uc5c6\uc74c',
                '\n'.join([
                    f'\uc0c8 \ubc84\uc804 {latest_version}\uc744 \ucc3e\uc558\uc9c0\ub9cc \uc124\uce58\ud30c\uc77c asset\uc774 \uc5c6\uc2b5\ub2c8\ub2e4.',
                    '',
                    f'GitHub Releases\uc5d0 {UPDATE_ASSET_NAME_PREFIX}{latest_version}.exe \ud615\uc2dd\uc758 \uc124\uce58\ud30c\uc77c\uc744 \uc62c\ub824\uc8fc\uc138\uc694.',
                ]),
            )
            self.status_var.set('\uc0c8 \ubc84\uc804\uc740 \uc788\uc73c\ub098 \uc124\uce58\ud30c\uc77c\uc744 \ucc3e\uc9c0 \ubabb\ud588\uc2b5\ub2c8\ub2e4.')
            return

        if messagebox.askyesno(
            '\uc0c8 \uc5c5\ub370\uc774\ud2b8 \ubc1c\uacac',
            '\n'.join([
                f'\ud604\uc7ac \ubc84\uc804: {APP_VERSION}',
                f'\ucd5c\uc2e0 \ubc84\uc804: {latest_version}',
                '',
                '\uc124\uce58\ud30c\uc77c\uc744 \ub2e4\uc6b4\ub85c\ub4dc\ud558\uace0 \uc5c5\ub370\uc774\ud2b8\ub97c \uc2dc\uc791\ud560\uae4c\uc694?',
            ]),
        ):
            self.download_and_run_update(release_info)
        else:
            self.status_var.set('\uc5c5\ub370\uc774\ud2b8\uac00 \ucde8\uc18c\ub418\uc5c8\uc2b5\ub2c8\ub2e4.')

    def download_and_run_update(self, release_info):
        if self.update_download_in_progress:
            messagebox.showinfo('\uc5c5\ub370\uc774\ud2b8 \ub2e4\uc6b4\ub85c\ub4dc', '\uc774\ubbf8 \uc5c5\ub370\uc774\ud2b8 \ud30c\uc77c\uc744 \ub2e4\uc6b4\ub85c\ub4dc\ud558\ub294 \uc911\uc785\ub2c8\ub2e4.')
            return

        self.update_download_in_progress = True
        self.status_var.set('\uc5c5\ub370\uc774\ud2b8 \uc124\uce58\ud30c\uc77c\uc744 \ub2e4\uc6b4\ub85c\ub4dc\ud558\ub294 \uc911\uc785\ub2c8\ub2e4.')
        thread = threading.Thread(target=self._download_update_worker, args=(release_info,), daemon=True)
        thread.start()

    def _download_update_worker(self, release_info):
        try:
            installer_path = download_update_installer(
                release_info.get('asset_url'),
                release_info.get('asset_name'),
            )
        except Exception as exc:
            self._schedule_ui(lambda: self._handle_update_download_failed(exc))
            return
        self._schedule_ui(lambda: self._handle_update_download_complete(installer_path, release_info.get('version') or ''))

    def _handle_update_download_failed(self, error):
        self.update_download_in_progress = False
        messagebox.showerror(
            '\uc5c5\ub370\uc774\ud2b8 \ub2e4\uc6b4\ub85c\ub4dc \uc2e4\ud328',
            '\n'.join([
                '\uc5c5\ub370\uc774\ud2b8 \uc124\uce58\ud30c\uc77c \ub2e4\uc6b4\ub85c\ub4dc\uc5d0 \uc2e4\ud328\ud588\uc2b5\ub2c8\ub2e4.',
                '',
                str(error),
            ]),
        )
        self.status_var.set('\uc5c5\ub370\uc774\ud2b8 \ub2e4\uc6b4\ub85c\ub4dc\uc5d0 \uc2e4\ud328\ud588\uc2b5\ub2c8\ub2e4.')

    def _handle_update_download_complete(self, installer_path, latest_version):
        self.update_download_in_progress = False
        self.status_var.set('\uc5c5\ub370\uc774\ud2b8 \uc124\uce58\ud30c\uc77c \ub2e4\uc6b4\ub85c\ub4dc\uac00 \uc644\ub8cc\ub418\uc5c8\uc2b5\ub2c8\ub2e4.')
        if not messagebox.askyesno(
            '\uc5c5\ub370\uc774\ud2b8 \uc2e4\ud589',
            '\n'.join([
                f'\ubc84\uc804 {latest_version} \uc124\uce58\ud30c\uc77c \ub2e4\uc6b4\ub85c\ub4dc\uac00 \uc644\ub8cc\ub418\uc5c8\uc2b5\ub2c8\ub2e4.',
                '',
                '\uc124\uce58\ud30c\uc77c\uc744 \uc2e4\ud589\ud558\uace0 \ud604\uc7ac \ud504\ub85c\uadf8\ub7a8\uc744 \uc885\ub8cc\ud560\uae4c\uc694?',
            ]),
        ):
            return

        try:
            os.startfile(installer_path)
        except OSError as exc:
            messagebox.showerror('\uc5c5\ub370\uc774\ud2b8 \uc2e4\ud589 \uc2e4\ud328', f'\uc124\uce58\ud30c\uc77c\uc744 \uc2e4\ud589\ud558\uc9c0 \ubabb\ud588\uc2b5\ub2c8\ub2e4.\n{exc}')
            return
        self.on_close()

    def show_previous_version_history(self):
        previous_rows = list(VERSION_HISTORY[1:])
        if self.version_history_window is not None and self.version_history_window.winfo_exists():
            self.version_history_window.lift()
            self.version_history_window.focus_force()
            return

        parent_window = self.about_window if self.about_window is not None and self.about_window.winfo_exists() else self.root
        window = tk.Toplevel(parent_window)
        self.version_history_window = window
        window.title('이전 버전 이력')
        window.configure(bg='#eef1f4')
        window.resizable(False, False)
        window.transient(parent_window)

        def close_version_history_window():
            if window.winfo_exists():
                window.destroy()
            self.version_history_window = None

        window.protocol('WM_DELETE_WINDOW', close_version_history_window)

        outer = tk.Frame(window, bg='#eef1f4', padx=14, pady=14)
        outer.pack(fill='both', expand=True)

        card = tk.Frame(
            outer,
            bg='white',
            bd=0,
            highlightbackground='#d4dae2',
            highlightthickness=1,
            padx=16,
            pady=14,
        )
        card.pack(fill='both', expand=True)

        tk.Label(
            card,
            text='이전 버전 업데이트 이력',
            bg='white',
            fg='#202020',
            font=('Malgun Gothic', 15, 'bold'),
        ).pack(anchor='w')

        tk.Label(
            card,
            text='현재 버전을 제외한 이전 버전의 주요 변경 내용을 확인할 수 있습니다.',
            bg='white',
            fg='#5f6b7a',
            font=('Malgun Gothic', 9),
        ).pack(anchor='w', pady=(4, 10))

        if not previous_rows:
            tk.Label(
                card,
                text='이전 버전 이력이 없습니다.',
                bg='#f8fafc',
                fg='#344150',
                font=('Malgun Gothic', 10),
                padx=18,
                pady=18,
                anchor='w',
            ).pack(fill='x')
        else:
            table_shell = tk.Frame(card, bg='#d4dae2', padx=1, pady=1)
            table_shell.pack(fill='both', expand=True)

            table_frame = tk.Frame(table_shell, bg='white')
            table_frame.pack(fill='both', expand=True)

            column_specs = (
                ('버전', 86, 'center', 0),
                ('주요 기능', 220, 'w', 210),
                ('날짜', 110, 'center', 0),
            )

            header_frame = tk.Frame(table_frame, bg='#eef3f8')
            header_frame.pack(fill='x')

            for column_index, (title, min_width, anchor, _wraplength) in enumerate(column_specs):
                header_frame.grid_columnconfigure(column_index, minsize=min_width, weight=1 if column_index == 1 else 0)
                tk.Label(
                    header_frame,
                    text=title,
                    bg='#eef3f8',
                    fg='#202020',
                    font=('Malgun Gothic', 10, 'bold'),
                    padx=6,
                    pady=8,
                    anchor=anchor,
                    bd=0,
                ).grid(row=0, column=column_index, sticky='nsew')

            tk.Frame(table_frame, bg='#dfe5ec', height=1).pack(fill='x')

            body_frame = tk.Frame(table_frame, bg='white')
            body_frame.pack(fill='both', expand=True)

            canvas = tk.Canvas(
                body_frame,
                bg='white',
                width=440,
                height=180,
                highlightthickness=0,
                bd=0,
            )
            scrollbar = tk.Scrollbar(body_frame, orient='vertical', command=canvas.yview)
            rows_frame = tk.Frame(canvas, bg='white')
            window_id = canvas.create_window((0, 0), window=rows_frame, anchor='nw')
            canvas.configure(yscrollcommand=scrollbar.set)
            canvas.pack(side='left', fill='both', expand=True)
            scrollbar.pack(side='right', fill='y')

            def refresh_scroll_region(_event=None):
                canvas.configure(scrollregion=canvas.bbox('all'))

            def sync_rows_width(event):
                canvas.itemconfigure(window_id, width=event.width)

            def on_mousewheel(event):
                if event.delta:
                    canvas.yview_scroll(int(-event.delta / 120), 'units')
                return 'break'

            rows_frame.bind('<Configure>', refresh_scroll_region)
            canvas.bind('<Configure>', sync_rows_width)
            canvas.bind('<MouseWheel>', on_mousewheel)

            for row_index, row_values in enumerate(previous_rows):
                row_bg = '#ffffff' if row_index % 2 == 0 else '#f8fafc'
                for column_index, (_title, min_width, anchor, wraplength) in enumerate(column_specs):
                    rows_frame.grid_columnconfigure(column_index, minsize=min_width, weight=1 if column_index == 1 else 0)
                    tk.Label(
                        rows_frame,
                        text=row_values[column_index],
                        bg=row_bg,
                        fg='#202020',
                        font=('Malgun Gothic', 10),
                        padx=8,
                        pady=8,
                        anchor=anchor,
                        justify='left',
                        wraplength=wraplength if wraplength else 0,
                        bd=0,
                    ).grid(row=row_index * 2, column=column_index, sticky='nsew')

                if row_index < len(previous_rows) - 1:
                    tk.Frame(rows_frame, bg='#edf1f5', height=1).grid(
                        row=row_index * 2 + 1,
                        column=0,
                        columnspan=len(column_specs),
                        sticky='ew',
                    )

        button_frame = tk.Frame(card, bg='white', pady=(12, 0))
        button_frame.pack(fill='x')

        tk.Button(
            button_frame,
            text='닫기',
            width=10,
            command=close_version_history_window,
            bg='#f3f4f6',
            fg='#202020',
            font=('Malgun Gothic', 10, 'bold'),
            relief='solid',
            bd=1,
            activebackground='#e7ebf0',
            activeforeground='#202020',
        ).pack(side='right')

        window.update_idletasks()
        popup_width = max(window.winfo_width(), 520)
        popup_height = max(window.winfo_height(), 360)
        parent_x = parent_window.winfo_rootx()
        parent_y = parent_window.winfo_rooty()
        parent_width = parent_window.winfo_width()
        parent_height = parent_window.winfo_height()
        x = parent_x + max((parent_width - popup_width) // 2, 20)
        y = parent_y + max((parent_height - popup_height) // 2, 20)
        window.geometry(f'{popup_width}x{popup_height}+{x}+{y}')
        window.focus_force()

    def show_help_about(self):
        if self.about_window is not None and self.about_window.winfo_exists():
            self.about_window.lift()
            self.about_window.focus_force()
            return

        window = tk.Toplevel(self.root)
        self.about_window = window
        window.title('버전 정보')
        window.configure(bg='#eef1f4')
        window.resizable(False, False)
        window.transient(self.root)

        def close_about_window():
            if self.version_history_window is not None and self.version_history_window.winfo_exists():
                self.version_history_window.destroy()
            if window.winfo_exists():
                window.destroy()
            self.about_window = None
            self.about_logo_image = None
            self.version_history_window = None

        window.protocol('WM_DELETE_WINDOW', close_about_window)

        outer = tk.Frame(window, bg='#eef1f4', padx=14, pady=14)
        outer.pack(fill='both', expand=True)

        card = tk.Frame(
            outer,
            bg='white',
            bd=0,
            highlightbackground='#d4dae2',
            highlightthickness=1,
        )
        card.pack(fill='both', expand=True)

        top_section = tk.Frame(card, bg='white', padx=16, pady=14)
        top_section.pack(fill='x')

        tk.Label(
            top_section,
            text='버전 정보',
            bg='white',
            fg='#202020',
            font=('Malgun Gothic', 18, 'bold'),
        ).pack(anchor='w')

        tk.Label(
            top_section,
            text='현재 버전, 프로그램 정보, 업데이트 이력을 한 번에 확인할 수 있습니다.',
            bg='white',
            fg='#5f6b7a',
            font=('Malgun Gothic', 10),
        ).pack(anchor='w', pady=(4, 0))

        summary_shell = tk.Frame(card, bg='#d4dae2', padx=1, pady=1)
        summary_shell.pack(fill='x', padx=16, pady=(12, 0))

        summary_frame = tk.Frame(summary_shell, bg='white', padx=16, pady=14)
        summary_frame.pack(fill='x')

        version_header = tk.Frame(summary_frame, bg='white')
        version_header.pack(fill='x')

        version_text_frame = tk.Frame(version_header, bg='white')
        version_text_frame.pack(side='left', fill='x', expand=True)

        tk.Label(
            version_text_frame,
            text='현재 버전',
            bg='white',
            fg='#202020',
            font=('Malgun Gothic', 11, 'bold'),
        ).pack(anchor='w')

        tk.Label(
            version_text_frame,
            text='설치된 프로그램의 최신 버전 정보입니다.',
            bg='white',
            fg='#5f6b7a',
            font=('Malgun Gothic', 9),
        ).pack(anchor='w', pady=(2, 0))

        tk.Label(
            version_text_frame,
            text='우측 버전을 누르면 이전 버전 이력을 볼 수 있습니다.',
            bg='white',
            fg='#0a72c4',
            font=('Malgun Gothic', 9, 'bold'),
        ).pack(anchor='w', pady=(8, 0))

        tk.Button(
            version_header,
            text=f'Ver {APP_VERSION}',
            command=self.show_previous_version_history,
            cursor='hand2',
            bg='#0a72c4',
            fg='white',
            font=('Malgun Gothic', 12, 'bold'),
            padx=14,
            pady=8,
            relief='flat',
            bd=0,
            highlightthickness=0,
            activebackground='#095ea3',
            activeforeground='white',
        ).pack(side='right', anchor='ne')

        info_grid = tk.Frame(summary_frame, bg='white')
        info_grid.pack(fill='x', pady=(16, 0))
        info_grid.grid_columnconfigure(1, weight=1)

        program_rows = (
            ('프로그램명', 'XRF Auto Input System'),
            ('제작자', '우도권'),
            ('사용자', '품질팀'),
            ('사용 목적', 'XRF자료 자동업데이트'),
        )

        for row_index, (label_text, value_text) in enumerate(program_rows):
            bottom_pad = 5 if row_index < len(program_rows) - 1 else 0
            tk.Label(
                info_grid,
                text=f'{label_text} :',
                bg='white',
                fg='#202020',
                font=('Malgun Gothic', 10, 'bold'),
                anchor='w',
            ).grid(row=row_index, column=0, sticky='w', pady=(0, bottom_pad))
            tk.Label(
                info_grid,
                text=value_text,
                bg='white',
                fg='#202020',
                font=('Malgun Gothic', 10),
                anchor='w',
            ).grid(row=row_index, column=1, sticky='w', padx=(8, 0), pady=(0, bottom_pad))

        tk.Label(
            summary_frame,
            text='이 프로그램은 풍원공업(주) 전용 프로그램이며 당사 직원 외 무단배포, 무단사용 등을 금합니다.',
            bg='white',
            fg='#344150',
            font=('Malgun Gothic', 10),
            justify='left',
            wraplength=470,
        ).pack(anchor='w', pady=(16, 0))

        self.about_logo_image = load_about_logo_image()
        if self.about_logo_image is not None:
            tk.Label(
                summary_frame,
                image=self.about_logo_image,
                bg='white',
                bd=0,
                highlightthickness=0,
            ).pack(anchor='center', pady=(14, 0))

        tk.Frame(card, bg='#e6e9ef', height=1).pack(fill='x', padx=16, pady=(14, 0))

        section_frame = tk.Frame(card, bg='white', padx=16, pady=(12, 8))
        section_frame.pack(fill='x')

        tk.Label(
            section_frame,
            text='버전 업데이트 이력',
            bg='white',
            fg='#202020',
            font=('Malgun Gothic', 13, 'bold'),
        ).pack(anchor='w')

        tk.Label(
            section_frame,
            text='최근 버전부터 순서대로 주요 변경 내용을 볼 수 있습니다.',
            bg='white',
            fg='#5f6b7a',
            font=('Malgun Gothic', 9),
        ).pack(anchor='w', pady=(3, 0))

        table_shell = tk.Frame(card, bg='#d4dae2', padx=1, pady=1)
        table_shell.pack(fill='both', expand=True, padx=16, pady=(0, 10))

        table_frame = tk.Frame(table_shell, bg='white')
        table_frame.pack(fill='both', expand=True)

        column_specs = (
            ('버전', 86, 'center', 0),
            ('주요 기능', 220, 'w', 210),
            ('날짜', 110, 'center', 0),
        )

        header_frame = tk.Frame(table_frame, bg='#eef3f8')
        header_frame.pack(fill='x')

        for column_index, (title, min_width, anchor, _wraplength) in enumerate(column_specs):
            header_frame.grid_columnconfigure(column_index, minsize=min_width, weight=1 if column_index == 1 else 0)
            tk.Label(
                header_frame,
                text=title,
                bg='#eef3f8',
                fg='#202020',
                font=('Malgun Gothic', 10, 'bold'),
                padx=6,
                pady=8,
                anchor=anchor,
                bd=0,
            ).grid(row=0, column=column_index, sticky='nsew')

        tk.Frame(table_frame, bg='#dfe5ec', height=1).pack(fill='x')

        body_frame = tk.Frame(table_frame, bg='white')
        body_frame.pack(fill='both', expand=True)

        canvas = tk.Canvas(
            body_frame,
            bg='white',
            width=440,
            height=170,
            highlightthickness=0,
            bd=0,
        )
        scrollbar = tk.Scrollbar(body_frame, orient='vertical', command=canvas.yview)
        rows_frame = tk.Frame(canvas, bg='white')
        window_id = canvas.create_window((0, 0), window=rows_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        def refresh_scroll_region(_event=None):
            canvas.configure(scrollregion=canvas.bbox('all'))

        def sync_rows_width(event):
            canvas.itemconfigure(window_id, width=event.width)

        def on_mousewheel(event):
            if event.delta:
                canvas.yview_scroll(int(-event.delta / 120), 'units')
            return 'break'

        rows_frame.bind('<Configure>', refresh_scroll_region)
        canvas.bind('<Configure>', sync_rows_width)
        canvas.bind('<MouseWheel>', on_mousewheel)

        display_rows = list(VERSION_HISTORY)
        while len(display_rows) < VERSION_INFO_MIN_ROWS:
            display_rows.append(('', '', ''))

        for row_index, row_values in enumerate(display_rows):
            row_bg = '#eef6ff' if row_index == 0 and row_values[0] else ('#ffffff' if row_index % 2 == 0 else '#f8fafc')
            for column_index, (_title, min_width, anchor, wraplength) in enumerate(column_specs):
                rows_frame.grid_columnconfigure(column_index, minsize=min_width, weight=1 if column_index == 1 else 0)
                text = row_values[column_index]
                tk.Label(
                    rows_frame,
                    text=text,
                    bg=row_bg,
                    fg='#202020',
                    font=('Malgun Gothic', 10, 'bold' if row_index == 0 and row_values[0] else 'normal'),
                    padx=8,
                    pady=8,
                    anchor=anchor,
                    justify='left',
                    wraplength=wraplength if wraplength else 0,
                    bd=0,
                ).grid(row=row_index * 2, column=column_index, sticky='nsew')

            if row_index < len(display_rows) - 1:
                separator = '#dfe5ec' if row_values[0] else '#edf1f5'
                tk.Frame(rows_frame, bg=separator, height=1).grid(
                    row=row_index * 2 + 1,
                    column=0,
                    columnspan=len(column_specs),
                    sticky='ew',
                )

        button_frame = tk.Frame(card, bg='white', padx=16, pady=(0, 14))
        button_frame.pack(fill='x')

        tk.Button(
            button_frame,
            text='닫기',
            width=10,
            command=close_about_window,
            bg='#f3f4f6',
            fg='#202020',
            font=('Malgun Gothic', 10, 'bold'),
            relief='solid',
            bd=1,
            activebackground='#e7ebf0',
            activeforeground='#202020',
        ).pack(side='right')

        window.update_idletasks()
        popup_width = max(window.winfo_width(), 560)
        popup_height = max(window.winfo_height(), 660)
        root_x = self.root.winfo_rootx()
        root_y = self.root.winfo_rooty()
        root_width = self.root.winfo_width()
        root_height = self.root.winfo_height()
        x = root_x + max((root_width - popup_width) // 2, 20)
        y = root_y + max((root_height - popup_height) // 2, 20)
        window.geometry(f'{popup_width}x{popup_height}+{x}+{y}')
        window.grab_set()
        window.focus_force()

    def open_program_folder(self):
        try:
            os.startfile(application_dir())
        except OSError as exc:
            messagebox.showerror('열기 실패', f'프로그램 폴더를 열지 못했습니다.\n{exc}')

    def _build_ui(self):
        top_bar = tk.Frame(self.root, bg='#b4b4b4')
        top_bar.pack(fill='x', padx=18, pady=(10, 0))

        version_label = tk.Label(
            top_bar,
            text=f'Ver {APP_VERSION}',
            bg='#b4b4b4',
            fg='white',
            font=('Malgun Gothic', 11, 'bold'),
        )
        version_label.pack(side='left', anchor='w')

        self.top_logo_image = load_top_bar_logo_image()
        if self.top_logo_image is not None:
            tk.Label(
                top_bar,
                image=self.top_logo_image,
                bg='#b4b4b4',
                bd=0,
                highlightthickness=0,
            ).pack(side='left', anchor='w', padx=(14, 0))

        self.title_label = tk.Label(
            self.root,
            text='XRF Report Auto Input System',
            bg='white',
            fg='#303030',
            font=('\ub9d1\uc740 \uace0\ub515', 22, 'bold'),
            relief='solid',
            bd=1,
            padx=24,
            pady=10,
        )
        self.title_label.pack(pady=(14, 12))

        body = tk.Frame(self.root, bg='#b4b4b4')
        body.pack(fill='both', expand=True, padx=18, pady=(0, 18))

        left_panel = tk.Frame(body, bg='#b4b4b4')
        left_panel.pack(side='left', fill='y', padx=(0, 18))

        tk.Label(
            left_panel,
            text='[재질 목록]',
            bg='#b4b4b4',
            fg='white',
            font=('맑은 고딕', 18, 'bold'),
        ).pack(anchor='w', pady=(0, 8))

        self.source_list = SelectableList(left_panel, width=360, height=560, bg='#c3c3c3')
        self.source_list.pack(fill='both', expand=True)

        source_select_row = tk.Frame(left_panel, bg='#b4b4b4')
        source_select_row.pack(fill='x', pady=(10, 0))

        tk.Button(
            source_select_row,
            text='전체선택',
            command=self.source_list.select_all,
            bg='white',
            fg='#202020',
            font=('맑은 고딕', 11, 'bold'),
            relief='solid',
            bd=1,
            padx=12,
            pady=8,
        ).pack(side='left', fill='x', expand=True)

        tk.Button(
            source_select_row,
            text='전체해제',
            command=self.source_list.clear_selection,
            bg='white',
            fg='#202020',
            font=('맑은 고딕', 11, 'bold'),
            relief='solid',
            bd=1,
            padx=12,
            pady=8,
        ).pack(side='left', fill='x', expand=True, padx=(10, 0))

        left_buttons = tk.Frame(left_panel, bg='#b4b4b4')
        left_buttons.pack(fill='x', pady=(14, 0))

        tk.Button(
            left_buttons,
            text='재질파일불러오기',
            command=self.load_source_files,
            bg='white',
            fg='#202020',
            font=('맑은 고딕', 13, 'bold'),
            relief='solid',
            bd=1,
            padx=18,
            pady=12,
        ).pack(side='left', fill='x', expand=True)

        tk.Button(
            left_buttons,
            text='목록 지우기',
            command=self.clear_source_files,
            bg='white',
            fg='#202020',
            font=('맑은 고딕', 13, 'bold'),
            relief='solid',
            bd=1,
            padx=18,
            pady=12,
        ).pack(side='left', fill='x', expand=True, padx=(10, 0))

        right_panel = tk.Frame(body, bg='#b4b4b4')
        right_panel.pack(side='left', fill='both', expand=True)

        control_panel = tk.Frame(right_panel, bg='#b4b4b4')
        control_panel.pack(fill='x')

        tk.Button(
            control_panel,
            text='기존 저장 파일 불러오기',
            command=self.load_target_workbook,
            bg='white',
            fg='#202020',
            font=('맑은 고딕', 13, 'bold'),
            relief='solid',
            bd=1,
            padx=16,
            pady=10,
        ).grid(row=0, column=0, sticky='ew', padx=(0, 10), pady=(0, 10))

        tk.Button(
            control_panel,
            text='업체 TOTAL 불러오기',
            command=self.load_total_workbook,
            bg='white',
            fg='#202020',
            font=('맑은 고딕', 13, 'bold'),
            relief='solid',
            bd=1,
            padx=16,
            pady=10,
        ).grid(row=0, column=1, sticky='ew', pady=(0, 10))

        control_panel.grid_columnconfigure(0, weight=1)
        control_panel.grid_columnconfigure(1, weight=1)

        self.target_info_box = ScrollableInfoBox(
            control_panel,
            width=760,
            height=122,
            initial_text=self.target_info_message,
        )
        self.target_info_box.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(0, 10))

        self.total_info_label = tk.Label(
            control_panel,
            textvariable=self.total_info_var,
            bg='white',
            fg='#303030',
            font=('\ub9d1\uc740 \uace0\ub515', 11, 'bold'),
            relief='solid',
            bd=1,
            anchor='w',
            justify='left',
            wraplength=760,
            padx=12,
            pady=12,
        )
        self.total_info_label.grid(row=2, column=0, columnspan=2, sticky='ew')

        self.middle_panel = tk.Frame(right_panel, bg='#b4b4b4')
        self.middle_panel.pack(fill='both', expand=True, pady=(16, 16))

        self.delete_panel = tk.Frame(self.middle_panel, bg='#b4b4b4')
        self.delete_panel.pack(side='left', fill='both', expand=True, padx=(0, 14))

        self.delete_header_label = tk.Label(
            self.delete_panel,
            text='[\uc0ad\uc81c\ud560 \uc2dc\ud2b8] \uae30\uc874 \uc800\uc7a5 \ud30c\uc77c\uc5d0\uc11c \uc0ad\uc81c \ud6c4 \uc0c8 \uc2dc\ud2b8\ub97c \uc0bd\uc785\ud569\ub2c8\ub2e4.',
            bg='#b4b4b4',
            fg='white',
            font=('\ub9d1\uc740 \uace0\ub515', 14, 'bold'),
            justify='left',
        )
        self.delete_header_label.pack(anchor='w', pady=(0, 8))

        self.target_sheet_list = SelectableList(self.delete_panel, width=420, height=320, bg='#c3c3c3')
        self.target_sheet_list.pack(fill='both', expand=True)

        target_select_row = tk.Frame(self.delete_panel, bg='#b4b4b4')
        target_select_row.pack(fill='x', pady=(10, 0))

        tk.Button(
            target_select_row,
            text='전체선택',
            command=self.target_sheet_list.select_all,
            bg='white',
            fg='#202020',
            font=('맑은 고딕', 11, 'bold'),
            relief='solid',
            bd=1,
            padx=12,
            pady=8,
        ).pack(side='left', fill='x', expand=True)

        tk.Button(
            target_select_row,
            text='전체해제',
            command=self.target_sheet_list.clear_selection,
            bg='white',
            fg='#202020',
            font=('맑은 고딕', 11, 'bold'),
            relief='solid',
            bd=1,
            padx=12,
            pady=8,
        ).pack(side='left', fill='x', expand=True, padx=(10, 0))

        option_row = tk.Frame(self.delete_panel, bg='#b4b4b4')
        option_row.pack(fill='x', pady=(12, 0))

        tk.Checkbutton(
            option_row,
            text='파일 최적화 사용 (xlsx/xlsm)',
            variable=self.optimize_var,
            bg='#b4b4b4',
            fg='white',
            activebackground='#b4b4b4',
            activeforeground='white',
            selectcolor='#b4b4b4',
            font=('맑은 고딕', 12, 'bold'),
        ).pack(side='left', anchor='w')

        action_row = tk.Frame(self.delete_panel, bg='#b4b4b4')
        action_row.pack(fill='x', pady=(12, 0))

        tk.Button(
            action_row,
            text='시트 삽입 실행',
            command=self.run_insert,
            bg='white',
            fg='#202020',
            font=('맑은 고딕', 14, 'bold'),
            relief='solid',
            bd=1,
            padx=18,
            pady=14,
        ).pack(side='left', fill='x', expand=True)

        tk.Button(
            action_row,
            text='유해물질표 업데이트',
            command=self.run_update_hazardous_table,
            bg='white',
            fg='#202020',
            font=('맑은 고딕', 14, 'bold'),
            relief='solid',
            bd=1,
            padx=18,
            pady=14,
        ).pack(side='left', fill='x', expand=True, padx=(10, 0))

        self.result_panel = tk.Frame(self.middle_panel, bg='white', relief='solid', bd=1)
        self.result_panel.pack(side='left', fill='both', expand=True)

        self.result_label = tk.Label(
            self.result_panel,
            textvariable=self.result_var,
            bg='white',
            fg='#0a72c4',
            font=('\ub9d1\uc740 \uace0\ub515', 72, 'bold'),
            pady=40,
        )
        self.result_label.pack(fill='x')

        self.status_label = tk.Label(
            self.result_panel,
            textvariable=self.status_var,
            bg='white',
            fg='#404040',
            font=('\ub9d1\uc740 \uace0\ub515', 13, 'bold'),
            justify='left',
            wraplength=360,
            padx=18,
            pady=16,
        )
        self.status_label.pack(fill='both', expand=True)

    def set_status(self, message, result_text='READY', result_color='#0a72c4'):
        self.status_var.set(message)
        self.result_var.set(result_text)
        self.root.update_idletasks()
        for widget in self.root.winfo_children():
            self._apply_result_color(widget, result_color)

    def _apply_result_color(self, widget, color):
        if isinstance(widget, tk.Label) and widget.cget('textvariable') == str(self.result_var):
            widget.configure(fg=color)
            return
        for child in widget.winfo_children():
            self._apply_result_color(child, color)

    def update_total_sheet_order_lookup(self):
        if not self.total_source_path:
            self.total_sheet_order_lookup = {}
            return

        records = peek_cached_total_records(self.total_source_path) or []
        self.total_sheet_order_lookup = build_total_sheet_order_lookup(records)

    def refresh_target_sheet_items(self, selected_names=None):
        selected_set = set(selected_names or [])
        self.target_sheet_list.set_items(
            [
                {
                    'label': format_sheet_name_with_total_order(name, self.total_sheet_order_lookup),
                    'value': name,
                    'selected': name in selected_set,
                }
                for name in self.target_sheet_names
            ]
        )

    def update_total_sheet_order_lookup(self):
        if not self.total_source_path:
            self.total_sheet_order_lookup = {}
            return

        records = peek_cached_total_records(self.total_source_path) or []
        self.total_sheet_order_lookup = build_total_sheet_order_lookup(records)

    def refresh_target_sheet_items(self, selected_names=None):
        selected_set = set(selected_names or [])
        self.target_sheet_list.set_items(
            [
                {
                    'label': format_sheet_name_with_total_order(name, self.total_sheet_order_lookup),
                    'value': name,
                    'selected': name in selected_set,
                }
                for name in self.target_sheet_names
            ]
        )

    def set_target_workbooks(self, inspections, failures, common_sheet_names, detected_sheet_names, selected_names=None):
        self.target_workbook_paths = [item['path'] for item in inspections]
        self.target_workbook_path = self.target_workbook_paths[0] if self.target_workbook_paths else ''
        self.target_sheet_names = list(common_sheet_names)
        self.detected_target_sheet_name = detected_sheet_names[0] if detected_sheet_names else ''

        self.refresh_target_sheet_items(selected_names=selected_names)

        info_lines = []
        if not self.target_workbook_paths:
            info_lines.append('기존 저장 파일 미선택')
        else:
            info_lines.append(f'기존 저장 파일: {len(self.target_workbook_paths)}개 선택')
            for target_path in self.target_workbook_paths:
                info_lines.append(f'- {os.path.basename(target_path)}')
            if self.target_sheet_names:
                info_lines.append(f'공통 시트: {len(self.target_sheet_names)}개')
            else:
                info_lines.append('공통 시트를 찾지 못했습니다.')
            unique_detected = list(dict.fromkeys(detected_sheet_names))
            if unique_detected:
                if len(unique_detected) == 1:
                    info_lines.append(f'유해물질표 시트: {unique_detected[0]}')
                else:
                    info_lines.append(f'유해물질표 시트: {unique_detected[0]} 외 {len(unique_detected) - 1}개')
            else:
                info_lines.append('유해물질표 시트를 자동으로 찾지 못했습니다.')
        if failures:
            info_lines.append(f'읽기 실패: {len(failures)}개')
        self.target_info_box.set_text('\n'.join(info_lines))

    def refresh_source_list(self, selected_values=None, default_selected=False):
        selected_set = set(selected_values or [])
        labels = build_file_labels(self.source_paths)
        items = [
            {
                'label': label,
                'value': path,
                'selected': (path in selected_set) if selected_values is not None else default_selected,
            }
            for label, path in zip(labels, self.source_paths)
        ]
        self.source_list.set_items(items)

    def clear_source_files(self):
        self.source_paths = []
        self.refresh_source_list()
        self.set_status('엑셀 목록을 비웠습니다.', 'READY', '#0a72c4')

    def load_source_files(self):
        selected_paths = filedialog.askopenfilenames(
            title='삽입할 엑셀 파일 선택',
            filetypes=EXCEL_FILE_TYPES,
        )
        if not selected_paths:
            return

        target_keys = {normalized_path(path) for path in self.target_workbook_paths}
        total_key = normalized_path(self.total_source_path) if self.total_source_path else None
        filtered_paths = []
        seen = set()
        skipped = 0

        for path in selected_paths:
            path_key = normalized_path(path)
            if path_key in seen:
                continue
            if path_key in target_keys:
                skipped += 1
                continue
            if total_key and path_key == total_key:
                skipped += 1
                continue
            seen.add(path_key)
            filtered_paths.append(os.path.abspath(path))

        if not filtered_paths:
            messagebox.showwarning('확인 필요', '기존 저장 파일 또는 업체 TOTAL 파일은 엑셀 목록에 넣을 수 없습니다.')
            return

        self.source_paths = filtered_paths
        self.refresh_source_list()
        message = f'엑셀 파일 {len(self.source_paths)}개를 불러왔습니다.'
        if skipped:
            message += f' 제외된 파일 {skipped}개가 있습니다.'
        self.set_status(message, 'READY', '#0a72c4')
    def set_target_workbooks_loading_state(self, paths):
        self.target_workbook_paths = list(paths)
        self.target_workbook_path = self.target_workbook_paths[0] if self.target_workbook_paths else ''
        self.target_sheet_names = []
        self.detected_target_sheet_name = ''
        self.target_sheet_list.set_items([])

        info_lines = []
        if not self.target_workbook_paths:
            info_lines.append('기존 저장 파일 미선택')
        else:
            info_lines.append(f'기존 저장 파일: {len(self.target_workbook_paths)}개 선택')
            for target_path in self.target_workbook_paths:
                info_lines.append(f'- {os.path.basename(target_path)}')
            info_lines.append('공통 시트: 확인 중...')
            info_lines.append('유해물질표 시트: 확인 중...')
        self.target_info_box.set_text('\n'.join(info_lines))

    def ensure_target_metadata_ready(self):
        if not self.target_inspection_pending:
            return True
        message = '기존 저장 파일 정보를 확인 중입니다. 잠시만 기다려 주세요.'
        messagebox.showinfo('불러오는 중', message)
        self.set_status(message, 'CHECK', '#c0392b')
        return False

    def ensure_total_metadata_ready(self):
        if not self.total_inspection_pending:
            return True
        message = '업체 TOTAL 파일 정보를 확인 중입니다. 잠시만 기다려 주세요.'
        messagebox.showinfo('불러오는 중', message)
        self.set_status(message, 'CHECK', '#c0392b')
        return False

    def _deliver_target_inspection_result(self, request_id, inspections, failures, common_sheet_names, detected_sheet_names):
        if request_id != self.target_load_request_id or not self.root.winfo_exists():
            return

        self.target_inspection_pending = False
        if not inspections:
            error_message = failures[0]['error'] if failures else '기존 저장 파일을 읽지 못했습니다.'
            self.set_target_workbooks([], failures, [], [])
            messagebox.showerror('불러오기 실패', error_message)
            self.set_status(error_message, 'CHECK', '#c0392b')
            return

        self.set_target_workbooks(inspections, failures, common_sheet_names, detected_sheet_names)
        message = f'기존 저장 파일 {len(inspections)}개를 불러왔습니다.'
        if common_sheet_names:
            message += f' 공통 시트 {len(common_sheet_names)}개를 확인했습니다.'
        if failures:
            message += f' 읽기 실패 {len(failures)}개가 있습니다.'
        self.set_status(message, 'READY', '#0a72c4')

    def _inspect_target_workbooks_async(self, request_id, selected_paths):
        try:
            result = inspect_target_workbooks_in_background(selected_paths)
        except RuntimeError as exc:
            try:
                self.root.after(
                    0,
                    lambda: self._deliver_target_inspection_result(
                        request_id,
                        [],
                        [{'path': '', 'error': str(exc)}],
                        [],
                        [],
                    ),
                )
            except tk.TclError:
                pass
            return

        try:
            self.root.after(0, lambda: self._deliver_target_inspection_result(request_id, *result))
        except tk.TclError:
            pass

    def load_target_workbook(self):
        selected_paths = filedialog.askopenfilenames(
            title='기존 저장 파일 선택',
            filetypes=EXCEL_FILE_TYPES,
        )
        if not selected_paths:
            return

        normalized_selected_paths = []
        seen = set()
        for path in selected_paths:
            abs_path = os.path.abspath(path)
            path_key = normalized_path(abs_path)
            if path_key in seen:
                continue
            seen.add(path_key)
            normalized_selected_paths.append(abs_path)

        if not normalized_selected_paths:
            return

        selected_values = self.source_list.get_selected_values() if self.source_list.has_items() else []
        target_keys = {normalized_path(path) for path in normalized_selected_paths}
        if self.source_paths:
            self.source_paths = [path for path in self.source_paths if normalized_path(path) not in target_keys]
            self.refresh_source_list(selected_values=selected_values)

        if self.total_source_path and normalized_path(self.total_source_path) in target_keys:
            self.total_source_path = ''
            self.detected_total_sheet_name = ''
            self.total_sheet_order_lookup = {}
            self.total_inspection_pending = False
            self.total_load_request_id += 1
            self.total_info_var.set('업체 TOTAL 파일 미선택')

        self.target_load_request_id += 1
        request_id = self.target_load_request_id
        self.target_inspection_pending = True
        self.set_target_workbooks_loading_state(normalized_selected_paths)
        self.set_status(
            f'기존 저장 파일 {len(normalized_selected_paths)}개를 바로 불러왔습니다. 시트 정보를 확인 중입니다.',
            'READY',
            '#0a72c4',
        )
        threading.Thread(
            target=self._inspect_target_workbooks_async,
            args=(request_id, normalized_selected_paths),
            daemon=True,
        ).start()

    def _deliver_total_inspection_result(self, request_id, selected_path, layout, error_message=None):
        if request_id != self.total_load_request_id or not self.root.winfo_exists():
            return

        self.total_inspection_pending = False
        if error_message:
            self.total_source_path = ''
            self.detected_total_sheet_name = ''
            self.total_sheet_order_lookup = {}
            self.total_info_var.set('업체 TOTAL 파일 미선택')
            if self.target_sheet_names:
                selected_names = self.target_sheet_list.get_selected_values() if self.target_sheet_list.has_items() else []
                self.refresh_target_sheet_items(selected_names=selected_names)
            messagebox.showerror('불러오기 실패', error_message)
            self.set_status(error_message, 'CHECK', '#c0392b')
            return

        self.total_source_path = os.path.abspath(selected_path)
        self.detected_total_sheet_name = layout['sheet_name'] if layout else ''
        self.update_total_sheet_order_lookup()
        info_lines = [f'업체 TOTAL 파일: {self.total_source_path}']
        if self.detected_total_sheet_name:
            info_lines.append(f'데이터 시트: {self.detected_total_sheet_name}')
        else:
            info_lines.append('TOTAL 데이터 시트를 자동으로 찾지 못했습니다.')
        self.total_info_var.set('\n'.join(info_lines))
        if self.target_sheet_names:
            selected_names = self.target_sheet_list.get_selected_values() if self.target_sheet_list.has_items() else []
            self.refresh_target_sheet_items(selected_names=selected_names)
        self.set_status('업체 TOTAL 파일을 불러왔습니다.', 'READY', '#0a72c4')

    def _inspect_total_workbook_async(self, request_id, selected_path):
        try:
            layout = inspect_total_workbook_in_background(selected_path)
        except RuntimeError as exc:
            try:
                self.root.after(0, lambda: self._deliver_total_inspection_result(request_id, selected_path, None, str(exc)))
            except tk.TclError:
                pass
            return

        try:
            self.root.after(0, lambda: self._deliver_total_inspection_result(request_id, selected_path, layout))
        except tk.TclError:
            pass

    def load_total_workbook(self):
        selected_path = filedialog.askopenfilename(
            title='업체 TOTAL 파일 선택',
            filetypes=EXCEL_FILE_TYPES,
        )
        if not selected_path:
            return

        target_keys = {normalized_path(path) for path in self.target_workbook_paths}
        if normalized_path(selected_path) in target_keys:
            messagebox.showwarning('확인 필요', '업체 TOTAL 파일은 기존 저장 파일과 다른 파일이어야 합니다.')
            return

        abs_path = os.path.abspath(selected_path)
        self.total_load_request_id += 1
        request_id = self.total_load_request_id
        self.total_inspection_pending = True
        self.total_source_path = abs_path
        self.detected_total_sheet_name = ''
        self.total_sheet_order_lookup = {}
        self.total_info_var.set(f'업체 TOTAL 파일: {abs_path}\n데이터 시트: 확인 중...')
        if self.target_sheet_names:
            selected_names = self.target_sheet_list.get_selected_values() if self.target_sheet_list.has_items() else []
            self.refresh_target_sheet_items(selected_names=selected_names)
        self.set_status('업체 TOTAL 파일을 바로 불러왔습니다. 데이터 시트를 확인 중입니다.', 'READY', '#0a72c4')
        threading.Thread(
            target=self._inspect_total_workbook_async,
            args=(request_id, abs_path),
            daemon=True,
        ).start()

    def refresh_target_sheet_list(self):
        if not self.target_workbook_paths or self.target_inspection_pending:
            return

        selected_names = self.target_sheet_list.get_selected_values() if self.target_sheet_list.has_items() else []
        inspections, failures, common_sheet_names, detected_sheet_names = inspect_target_workbooks(self.target_workbook_paths)
        self.set_target_workbooks(
            inspections,
            failures,
            common_sheet_names,
            detected_sheet_names,
            selected_names=selected_names,
        )


    def format_missing_path_preview(self, paths):
        preview = ', '.join(os.path.basename(path) for path in paths[:3])
        if len(paths) > 3:
            preview += f' 외 {len(paths) - 3}개'
        return preview

    def ensure_target_workbooks_exist(self):
        missing_paths = [path for path in self.target_workbook_paths if not os.path.isfile(path)]
        if not missing_paths:
            return True

        preview = self.format_missing_path_preview(missing_paths)
        reload_now = messagebox.askyesno(
            '경로 변경 확인',
            f'선택한 기존 저장 파일을 찾지 못했습니다.\n{preview}\n다시 불러오시겠습니까?',
        )
        if not reload_now:
            self.set_status('기존 저장 파일 경로를 다시 선택해 주세요.', 'CHECK', '#c0392b')
            return False

        self.load_target_workbook()
        if not self.target_workbook_paths:
            self.set_status('기존 저장 파일을 다시 불러오지 못했습니다.', 'CHECK', '#c0392b')
            return False

        missing_after = [path for path in self.target_workbook_paths if not os.path.isfile(path)]
        if missing_after:
            self.set_status('기존 저장 파일 경로를 다시 확인해 주세요.', 'CHECK', '#c0392b')
            return False

        return True

    def ensure_total_source_exists(self):
        if os.path.isfile(self.total_source_path):
            return True

        reload_now = messagebox.askyesno(
            '경로 변경 확인',
            '선택한 업체 TOTAL 파일을 찾지 못했습니다.\n다시 불러오시겠습니까?',
        )
        if not reload_now:
            self.set_status('업체 TOTAL 파일 경로를 다시 선택해 주세요.', 'CHECK', '#c0392b')
            return False

        self.load_total_workbook()
        if not self.total_source_path or not os.path.isfile(self.total_source_path):
            self.set_status('업체 TOTAL 파일을 다시 불러오지 못했습니다.', 'CHECK', '#c0392b')
            return False

        return True

    def apply_optimization_if_enabled(self, target_path=None):
        if not self.optimize_var.get():
            return None
        path = target_path or self.target_workbook_path
        if not path:
            return None
        return optimize_excel_file(path)
    def format_optimization_message(self, optimization_result):
        if not optimization_result:
            return ''
        if not optimization_result.get('applied'):
            reason = optimization_result.get('reason')
            if reason == 'unsupported_extension':
                return '파일 최적화는 xlsx/xlsm 형식에서만 적용됩니다.'
            return '파일 최적화를 생략했습니다.'
        before_size = format_file_size(optimization_result['before_size'])
        after_size = format_file_size(optimization_result['after_size'])
        return f' 파일 최적화: {before_size} -> {after_size}'

    def run_insert(self):
        if not self.target_workbook_paths:
            messagebox.showwarning('확인 필요', '먼저 기존 저장 파일을 불러오세요.')
            return
        if not self.ensure_target_metadata_ready():
            return

        selected_paths = self.source_list.get_selected_values()
        sheet_names_to_delete = self.target_sheet_list.get_selected_values()
        if not selected_paths and not sheet_names_to_delete:
            messagebox.showwarning('확인 필요', '삽입할 엑셀 파일이나 삭제할 시트를 선택하세요.')
            return
        if (
            not selected_paths
            and sheet_names_to_delete
            and self.target_sheet_names
            and len(sheet_names_to_delete) >= len(self.target_sheet_names)
        ):
            messagebox.showwarning('확인 필요', '모든 시트를 삭제하려면 삽입할 엑셀 파일도 같이 선택하세요.')
            return

        success_items = []
        failure_items = []
        optimization_results = []
        total_deleted = 0
        total_inserted = 0

        try:
            insert_results, insert_failures = insert_sheets_into_target_workbooks(
                selected_paths,
                self.target_workbook_paths,
                sheet_names_to_delete,
            )
        except RuntimeError as exc:
            messagebox.showerror('실행 실패', str(exc))
            self.set_status(str(exc), 'NG', '#c0392b')
            return

        for item in insert_results:
            target_path = item['path']
            stats = item['stats']
            success_items.append((target_path, stats))
            total_deleted += stats['deleted_sheet_count']
            total_inserted += stats['inserted_sheet_count']
            optimization_result = self.apply_optimization_if_enabled(target_path)
            if optimization_result:
                optimization_results.append(optimization_result)

        for item in insert_failures:
            failure_items.append((item['path'], item['error']))

        if success_items:
            self.refresh_target_sheet_list()
        else:
            error_message = failure_items[0][1] if failure_items else '시트 삽입에 실패했습니다.'
            messagebox.showerror('실행 실패', error_message)
            self.set_status(error_message, 'NG', '#c0392b')
            return

        message_lines = [
            f'대상 파일: {len(self.target_workbook_paths)}개',
            f'성공: {len(success_items)}개',
            f'실패: {len(failure_items)}개',
            f'총 삭제된 시트: {total_deleted}개',
            f'총 삽입된 시트: {total_inserted}개',
        ]
        if len(success_items) == 1:
            message_lines.append(f"현재 시트 수: {success_items[0][1]['final_sheet_count']}개")
        if optimization_results:
            applied_count = sum(1 for result in optimization_results if result.get('applied'))
            message_lines.append(f'파일 최적화 적용: {applied_count}개')
        if failure_items:
            preview = ', '.join(os.path.basename(path) for path, _ in failure_items[:3])
            if len(failure_items) > 3:
                preview += f' 외 {len(failure_items) - 3}개'
            message_lines.append(f'실패 파일: {preview}')

        messagebox.showinfo('시트 삽입 완료', '\n'.join(message_lines))
        if failure_items:
            self.set_status(f'시트 삽입이 완료되었습니다. 성공 {len(success_items)}개 / 실패 {len(failure_items)}개', 'NG', '#c0392b')
        else:
            self.set_status(f'시트 삽입이 완료되었습니다. 성공 {len(success_items)}개', 'OK', '#0a72c4')
    def run_update_hazardous_table(self):
        if not self.target_workbook_paths:
            messagebox.showwarning('확인 필요', '먼저 기존 저장 파일을 불러오세요.')
            return
        if not self.total_source_path:
            messagebox.showwarning('확인 필요', '먼저 업체 TOTAL 파일을 불러오세요.')
            return
        if not self.ensure_target_metadata_ready():
            return
        if not self.ensure_total_metadata_ready():
            return
        if not self.ensure_target_workbooks_exist():
            return
        if not self.ensure_total_source_exists():
            return

        success_items = []
        failure_items = []
        optimization_results = []
        total_matched = 0
        total_unmatched = 0

        try:
            update_results, update_failures = update_hazardous_tables_from_total(
                self.total_source_path,
                self.target_workbook_paths,
            )
        except RuntimeError as exc:
            messagebox.showerror('업데이트 실패', str(exc))
            self.set_status(str(exc), 'NG', '#c0392b')
            return

        for item in update_results:
            target_path = item['path']
            stats = item['stats']
            success_items.append((target_path, stats))
            total_matched += stats['matched_rows']
            total_unmatched += stats['unmatched_rows']
            optimization_result = self.apply_optimization_if_enabled(target_path)
            if optimization_result:
                optimization_results.append(optimization_result)

        for item in update_failures:
            failure_items.append((item['path'], item['error']))

        if not success_items:
            error_message = failure_items[0][1] if failure_items else '유해물질표 업데이트에 실패했습니다.'
            messagebox.showerror('업데이트 실패', error_message)
            self.set_status(error_message, 'NG', '#c0392b')
            return

        message_lines = [
            f'대상 파일: {len(self.target_workbook_paths)}개',
            f'성공: {len(success_items)}개',
            f'실패: {len(failure_items)}개',
            f'총 매칭 성공 행: {total_matched}',
            f'총 미매칭 행: {total_unmatched}',
        ]
        if len(success_items) == 1:
            stats = success_items[0][1]
            message_lines.insert(0, f"TOTAL 시트: {stats['source_sheet_name']}")
            message_lines.insert(0, f"대상 시트: {stats['target_sheet_name']}")
            message_lines.append(f"사용되지 않은 TOTAL 행: {stats['unused_source_rows']}")
        if optimization_results:
            applied_count = sum(1 for result in optimization_results if result.get('applied'))
            message_lines.append(f'파일 최적화 적용: {applied_count}개')
        if failure_items:
            preview = ', '.join(os.path.basename(path) for path, _ in failure_items[:3])
            if len(failure_items) > 3:
                preview += f' 외 {len(failure_items) - 3}개'
            message_lines.append(f'실패 파일: {preview}')

        messagebox.showinfo('유해물질표 업데이트 완료', '\n'.join(message_lines))
        if failure_items:
            self.set_status(f'유해물질분석표 업데이트가 완료되었습니다. 성공 {len(success_items)}개 / 실패 {len(failure_items)}개', 'NG', '#c0392b')
        else:
            self.set_status(f'유해물질분석표 업데이트가 완료되었습니다. 성공 {len(success_items)}개', 'OK', '#0a72c4')


def main():
    root = tk.Tk()
    app = XRFReportApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()


































