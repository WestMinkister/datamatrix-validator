# 설정 파일 경로
CONFIG_FILE = "datamatrix_config.json"

def load_config():
    """설정 파일에서 구성 불러오기"""
    default_config = {
        "b_range_check": False,
        "b_min_value": 80,
        "b_max_value": 250,
        "i_n_check": True,
        "i_to_n_mapping": {str(i): 10 for i in range(10, 60)}
    }
    
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as f:
                return json.load(f)
        else:
            # 설정 파일이 없으면 기본 설정 저장
            with open(CONFIG_FILE, 'w') as f:
                json.dump(default_config, f, indent=2)
            return default_config
    except Exception as e:
        st.error(f"설정 파일 로드 중 오류: {str(e)}")
        return default_config

def save_config(config):
    """설정을 파일에 저장"""
    try:
        with open(CONFIG_FILE, 'w') as f:
            json.dump(config, f, indent=2)
        return True
    except Exception as e:
        st.error(f"설정 파일 저장 중 오류: {str(e)}")
        return False

def save_current_config():
    """현재 세션에서 설정 값을 파일로 저장"""
    config = {
        "b_range_check": st.session_state.b_range_check,
        "b_min_value": st.session_state.b_min_value,
        "b_max_value": st.session_state.b_max_value,
        "i_n_check": st.session_state.i_n_check,
        "i_to_n_mapping": st.session_state.i_to_n_mapping
    }
    return save_config(config)

import streamlit as st
import subprocess
import os
import io
import re
import tempfile
import sys
import platform
import numpy as np
from PIL import Image
import time
import shutil
import base64
from io import BytesIO
import json

# 페이지 설정을 가장 먼저 호출해야 함
st.set_page_config(
    page_title="DataMatrix 바코드 검증 도구",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 초기 설정 스크립트 실행 (페이지 설정 이후)
try:
    if os.path.exists("init_script.sh"):
        subprocess.run(["bash", "init_script.sh"], check=True)
        st.success("시스템 라이브러리 설치 완료")
except Exception as e:
    st.warning(f"시스템 라이브러리 설치 중 오류 발생: {str(e)}")

# CSS 스타일 적용
st.markdown("""
<style>
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    h1, h2, h3 {
        margin-bottom: 0.5rem;
    }
    .stProgress > div > div > div > div {
        background-color: #4CAF50;
    }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #F0FFF0;
        border: 1px solid #CCFFCC;
    }
    .warning-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #FFFAF0;
        border: 1px solid #FAEBD7;
    }
    .error-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #FFF0F0;
        border: 1px solid #FFCCCC;
    }
</style>
""", unsafe_allow_html=True)

# 각 라이브러리 개별 로드 시도
# OpenCV 로드 시도
try:
    import cv2
    HAVE_CV2 = True
except ImportError:
    HAVE_CV2 = False
    st.warning("OpenCV (cv2) 라이브러리를 불러올 수 없습니다. 이미지 처리 기능이 제한됩니다.")

# pylibdmtx 로드 시도
try:
    from pylibdmtx.pylibdmtx import decode
    HAVE_PYLIBDMTX = True
except ImportError:
    HAVE_PYLIBDMTX = False
    import platform
    if platform.system() == "Windows":
        st.warning("pylibdmtx 라이브러리를 불러올 수 없습니다.")
        st.info("Windows에서 pylibdmtx 설치하기: pip install pylibdmtx 후 libdmtx.dll 파일을 Python 실행 폴더에 복사하세요.")
    else:
        st.warning("pylibdmtx 라이브러리를 불러올 수 없습니다. 바코드 검출 기능을 사용할 수 없습니다.")
        st.info("pylibdmtx 설치를 위해서는 libdmtx 시스템 라이브러리가 필요합니다.")
    
    # 폴백 함수 정의
    def decode(image, **kwargs):
        return []

# pdf2image 로드 시도
try:
    import pdf2image
    HAVE_PDF2IMAGE = True
except ImportError:
    HAVE_PDF2IMAGE = False
    st.warning("pdf2image 라이브러리를 불러올 수 없습니다. PDF 이미지 추출 기능이 제한됩니다.")

# pypdfium2 로드 시도
try:
    import pypdfium2 as pdfium
    HAVE_PDFIUM = True
except ImportError:
    HAVE_PDFIUM = False
    st.warning("pypdfium2 라이브러리를 불러올 수 없습니다. PDF 처리 기능이 제한됩니다.")

# Office 관련 라이브러리 로드 시도
try:
    from openpyxl import load_workbook
    HAVE_OPENPYXL = True
except ImportError:
    HAVE_OPENPYXL = False
    st.warning("openpyxl 라이브러리를 불러올 수 없습니다. Excel 파일 처리 기능을 사용할 수 없습니다.")

try:
    from pptx import Presentation
    HAVE_PPTX = True
except ImportError:
    HAVE_PPTX = False
    st.warning("python-pptx 라이브러리를 불러올 수 없습니다. PowerPoint 파일 처리 기능을 사용할 수 없습니다.")

try:
    from PyPDF2 import PdfReader
    HAVE_PYPDF2 = True
except ImportError:
    HAVE_PYPDF2 = False
    st.warning("PyPDF2 라이브러리를 불러올 수 없습니다. PDF 텍스트 추출 기능을 사용할 수 없습니다.")

# 필요한 라이브러리 설치 확인 메시지
st.info("라이브러리 로드가 완료되었습니다. 일부 라이브러리가 로드되지 않은 경우 해당 기능이 제한될 수 있습니다.")

# 필요한 시스템 패키지 확인 (서버에 미리 설치되어 있어야 함)
def check_system_dependencies():
    """시스템에 필요한 라이브러리가 설치되어 있는지 확인"""
    # 운영체제 확인
    import platform
    
    current_os = platform.system()
    
    # Windows에서는 다른 검사 방법 사용
    if current_os == "Windows":
        # Windows용 검사 코드
        try:
            # pylibdmtx가 로드되었는지만 확인
            if not HAVE_PYLIBDMTX:
                st.warning("pylibdmtx 라이브러리를 불러올 수 없습니다. Windows용 설치 방법을 확인하세요.")
                st.info("Windows에 libdmtx를 설치하려면 https://github.com/dmtx/libdmtx/releases 에서 다운로드하세요.")
            
            # LibreOffice 확인 (Windows 방식)
            import os
            libreoffice_paths = [
                "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
                "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe"
            ]
            libreoffice_found = any(os.path.exists(path) for path in libreoffice_paths)
            
            if not libreoffice_found:
                st.warning("LibreOffice가 설치되어 있지 않습니다. Office 파일 변환이 작동하지 않을 수 있습니다.")
                st.info("LibreOffice를 https://www.libreoffice.org/download/download/ 에서 다운로드하세요.")
                
        except Exception as e:
            st.warning(f"시스템 확인 중 오류 발생: {str(e)}")
            st.info("일부 기능이 제한될 수 있지만, 앱은 계속 작동합니다.")
        
        return
    
    # Linux/macOS 검사 코드 (기존 코드)
    try:
        # libdmtx 확인
        result = subprocess.run(["pkg-config", "--exists", "libdmtx"],
                               stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if result.returncode != 0:
            st.warning("libdmtx가 설치되어 있지 않습니다. 바코드 검출이 작동하지 않을 수 있습니다.")
            if current_os == "Darwin":  # macOS
                st.info("macOS에서는 'brew install libdmtx'로 설치할 수 있습니다.")
            else:  # Linux
                st.info("Ubuntu에서는 'sudo apt-get install libdmtx0a libdmtx-dev'로 설치할 수 있습니다.")
    
    except Exception as e:
        st.warning(f"시스템 의존성 확인 중 오류 발생: {str(e)}")
        st.info("이 앱이 정상적으로 작동하려면 libdmtx, libreoffice, poppler-utils가 필요합니다.")

# =========================================================
# 유효성 검증 함수
# =========================================================

def validate_44x44_matrix(data, b_range_check=False, b_min_value=0, b_max_value=9999, i_n_check=False, i_to_n_mapping=None):
    """44x44 매트릭스 데이터 검증 함수"""
    result = {"valid": False, "errors": [], "warnings": [], "data": {}, "pattern_match": False}
    
    # 잘못된 구분자(,) 사용 확인
    if re.search(r'C[A-Za-z0-9]{3},', data) or re.search(r'I\d{2},', data) or re.search(r'W(?:LO|SE),', data):
        result["errors"].append("잘못된 구분자(,)를 사용했습니다. 구분자는 '.'(마침표)여야 합니다.")
        # 어떤 식별자에서 잘못된 구분자를 사용했는지 확인
        if re.search(r'C[A-Za-z0-9]{3},', data):
            result["errors"].append("C 식별자 뒤에 잘못된 구분자(,)를 사용했습니다.")
        if re.search(r'I\d{2},', data):
            result["errors"].append("I 식별자 뒤에 잘못된 구분자(,)를 사용했습니다.")
        if re.search(r'W(?:LO|SE),', data):
            result["errors"].append("W 식별자 뒤에 잘못된 구분자(,)를 사용했습니다.")
        if re.search(r'T\d{2},', data):
            result["errors"].append("T 식별자 뒤에 잘못된 구분자(,)를 사용했습니다.")
        if re.search(r'N\d{3},', data):
            result["errors"].append("N 식별자 뒤에 잘못된 구분자(,)를 사용했습니다.")
        if re.search(r'D\d{8},', data):
            result["errors"].append("D 식별자 뒤에 잘못된 구분자(,)를 사용했습니다.")
        if re.search(r'S\d{3},', data):
            result["errors"].append("S 식별자 뒤에 잘못된 구분자(,)를 사용했습니다.")
        if re.search(r'B[0-9]{120},', data):
            result["errors"].append("B 식별자 뒤에 잘못된 구분자(,)를 사용했습니다.")
        return result

    # 바코드 데이터에서 패턴 확인
    pattern = r'C([A-Za-z0-9]{3})\.I(\d{2})\.W(LO|SE)\.T(\d{2})\.N(\d{3})\.D(\d{8})\.S(\d{3})\.B([0-9]{120})\.'
    match = re.search(pattern, data)
    
    if not match:
        result["errors"].append("44x44 매트릭스 형식이 맞지 않습니다.")
        
        # 개별 패턴 확인으로 어떤 부분이 문제인지 디버깅
        C_match = re.search(r'C([A-Za-z0-9]{3})\.', data)
        if not C_match:
            result["errors"].append("C 식별자를 찾을 수 없거나 형식이 올바르지 않습니다")
        
        I_match = re.search(r'I(\d{2})\.', data)
        if not I_match:
            result["errors"].append("I 식별자를 찾을 수 없거나 형식이 올바르지 않습니다")
        
        W_match = re.search(r'W(LO|SE)\.', data)
        if not W_match:
            result["errors"].append("W 식별자를 찾을 수 없거나 형식이 올바르지 않습니다 (LO 또는 SE 값이어야 함)")
        
        T_match = re.search(r'T(\d{2})\.', data)
        if not T_match:
            result["errors"].append("T 식별자를 찾을 수 없거나 형식이 올바르지 않습니다")
        
        N_match = re.search(r'N(\d{3})\.', data)
        if not N_match:
            result["errors"].append("N 식별자를 찾을 수 없거나 형식이 올바르지 않습니다")
        
        D_match = re.search(r'D(\d{8})\.', data)
        if not D_match:
            result["errors"].append("D 식별자를 찾을 수 없거나 형식이 올바르지 않습니다 (YYYYMMDD 형식)")
        
        S_match = re.search(r'S(\d{3})\.', data)
        if not S_match:
            result["errors"].append("S 식별자를 찾을 수 없거나 형식이 올바르지 않습니다")
        
        B_match = re.search(r'B([0-9]{120})\.', data)
        if not B_match:
            result["errors"].append("B 식별자를 찾을 수 없거나 형식이 올바르지 않습니다 (120자리 숫자)")
        
        return result
    
    C_val, I_val, W_val, T_val, N_val, D_val, S_val, B_val = match.groups()
    
    # 데이터 저장
    result["data"] = {
        "C": C_val,
        "I": I_val,
        "W": W_val,
        "T": T_val,
        "N": N_val,
        "D": D_val,
        "S": S_val,
        "B": B_val
    }
    
    result["pattern_match"] = True
    
    # 추가 검증
    # D: 날짜 형식 검증 (YYYYMMDD)
    try:
        year = int(D_val[:4])
        month = int(D_val[4:6])
        day = int(D_val[6:8])
        
        if not (1900 <= year <= 2100):
            result["errors"].append(f"D 식별자: 연도 범위가 올바르지 않습니다 ({year})")
        if not (1 <= month <= 12):
            result["errors"].append(f"D 식별자: 월 범위가 올바르지 않습니다 ({month})")
        if not (1 <= day <= 31):
            result["errors"].append(f"D 식별자: 일 범위가 올바르지 않습니다 ({day})")
    except ValueError:
        result["errors"].append("D 식별자: 날짜 형식이 올바르지 않습니다")
    
    # B: 숫자 세트 검증
    B_sets = []
    non_zero_sets_count = 0
    
    for i in range(0, len(B_val), 4):
        if i+4 <= len(B_val):
            B_set = B_val[i:i+4]
            B_sets.append(B_set)
            if B_set != '0000':
                non_zero_sets_count += 1
    
    # N: B의 세트 수와 일치하는지 확인
    if int(N_val) != non_zero_sets_count:
        result["errors"].append(f"N 식별자: 값 {N_val}이 B 식별자의 비어있지 않은 세트 수 {non_zero_sets_count}와 일치하지 않습니다")
    
    # I 값에 따른 N 최대값 검증
    if i_n_check and i_to_n_mapping and I_val:
        try:
            i_val_int = int(I_val)
            current_n_value = int(N_val)
            
            # I 값에 해당하는 최대 N 값 찾기
            i_val_str = str(i_val_int)
            if i_val_str in i_to_n_mapping:
                max_n = i_to_n_mapping[i_val_str]
                if current_n_value > max_n:
                    result["errors"].append(f"N 식별자: I{I_val}에 대한 N 값이 최대 허용치({max_n})를 초과했습니다 (현재 값: {current_n_value})")
        except (ValueError, TypeError):
            # I 값이나 N 값이 정수로 변환할 수 없는 경우
            pass
    
    # B 값 범위 검사 (활성화된 경우)
    if b_range_check:
        out_of_range_sets = []
        for B_set in B_sets:
            if B_set != '0000':
                b_val = int(B_set)
                if b_val < b_min_value or b_val > b_max_value:
                    out_of_range_sets.append(f"{B_set} ({b_val})")
        
        if out_of_range_sets:
            error_msg = f"B 식별자: 다음 값들