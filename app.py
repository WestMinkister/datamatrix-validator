import streamlit as st
import re
import os
import io
import cv2
import numpy as np
import tempfile
import shutil
import pandas as pd
from PIL import Image
from pylibdmtx.pylibdmtx import decode
import base64
import pdf2image
import pypdfium2 as pdfium
from openpyxl import load_workbook
from pptx import Presentation

# 페이지 제목 설정
st.set_page_config(page_title="데이터매트릭스 검증기", layout="wide")
st.title("데이터매트릭스 검증기")
st.markdown("PDF, PPTX, XLSX 파일의 각 페이지에서 44x44 및 18x18 DataMatrix 바코드를 검색하고 검증합니다.")

# =========================================================
# 데이터 매트릭스 검증 함수
# =========================================================

def validate_44x44_matrix(data):
    """44x44 매트릭스 데이터 검증 함수"""
    result = {"valid": False, "errors": [], "data": {}, "pattern_match": False}
    
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
    
    # B: 숫자 세트가 오름차순인지 확인
    prev_set = None
    for B_set in B_sets:
        if B_set != '0000':
            if prev_set and int(B_set) <= int(prev_set):
                result["errors"].append(f"B 식별자: 숫자 세트가 오름차순이 아닙니다 ({prev_set} -> {B_set})")
            prev_set = B_set
    
    result["valid"] = len(result["errors"]) == 0
    
    return result

def validate_18x18_matrix(data):
    """18x18 매트릭스 데이터 검증 함수"""
    result = {"valid": False, "errors": [], "data": {}, "pattern_match": False}
    
    # 잘못된 구분자(,) 사용 확인
    if re.search(r'M[A-Za-z0-9]{4},', data) or re.search(r'I\d{2},', data) or re.search(r'C[A-Za-z0-9]{3},', data):
        result["errors"].append("잘못된 구분자(,)를 사용했습니다. 구분자는 '.'(마침표)여야 합니다.")
        # 어떤 식별자에서 잘못된 구분자를 사용했는지 확인
        if re.search(r'M[A-Za-z0-9]{4},', data):
            result["errors"].append("M 식별자 뒤에 잘못된 구분자(,)를 사용했습니다.")
        if re.search(r'I\d{2},', data):
            result["errors"].append("I 식별자 뒤에 잘못된 구분자(,)를 사용했습니다.")
        if re.search(r'C[A-Za-z0-9]{3},', data):
            result["errors"].append("C 식별자 뒤에 잘못된 구분자(,)를 사용했습니다.")
        if re.search(r'P\d{3},', data):
            result["errors"].append("P 식별자 뒤에 잘못된 구분자(,)를 사용했습니다.")
        return result

    # 바코드 데이터에서 패턴 확인
    pattern = r'M([A-Za-z0-9]{4})\.I(\d{2})\.C([A-Za-z0-9]{3})\.P(\d{3})\.'
    match = re.search(pattern, data)
    
    if not match:
        result["errors"].append("18x18 매트릭스 형식이 맞지 않습니다.")
        
        # 개별 패턴 확인으로 어떤 부분이 문제인지 디버깅
        M_match = re.search(r'M([A-Za-z0-9]{4})\.', data)
        if not M_match:
            result["errors"].append("M 식별자를 찾을 수 없거나 형식이 올바르지 않습니다 (4자리 문자+숫자 조합)")
        
        I_match = re.search(r'I(\d{2})\.', data)
        if not I_match:
            result["errors"].append("I 식별자를 찾을 수 없거나 형식이 올바르지 않습니다 (2자리 숫자)")
        
        C_match = re.search(r'C([A-Za-z0-9]{3})\.', data)
        if not C_match:
            result["errors"].append("C 식별자를 찾을 수 없거나 형식이 올바르지 않습니다 (3자리 문자+숫자 조합)")
        
        P_match = re.search(r'P(\d{3})\.', data)
        if not P_match:
            result["errors"].append("P 식별자를 찾을 수 없거나 형식이 올바르지 않습니다 (3자리 숫자)")
        
        return result
    
    M_val, I_val, C_val, P_val = match.groups()
    
    # 데이터 저장
    result["data"] = {
        "M": M_val,
        "I": I_val,
        "C": C_val,
        "P": P_val
    }
    
    result["pattern_match"] = True
    
    # 여기에 필요한 추가 검증 로직 추가
    
    result["valid"] = len(result["errors"]) == 0
    
    return result

def cross_validate_matrices(matrix_44x44, matrix_18x18):
    """두 매트릭스 간의 교차 검증"""
    errors = []
    
    # 둘 중 하나라도 패턴 매치가 실패한 경우
    if not matrix_44x44["pattern_match"] or not matrix_18x18["pattern_match"]:
        return ["교차 검증을 수행할 수 없습니다. 두 매트릭스의 기본 형식이 올바르지 않습니다."]
    
    # 1. [18x18]의 I 값과 [44x44]의 I 값이 동일한지 확인
    if matrix_18x18["data"]["I"] != matrix_44x44["data"]["I"]:
        errors.append(f"교차 검증 실패: [18x18]의 I({matrix_18x18['data']['I']})와 [44x44]의 I({matrix_44x44['data']['I']})가 일치하지 않습니다.")
    
    # 2. [18x18]의 C 값과 [44x44]의 C 값이 일치하는지 확인
    if matrix_18x18["data"]["C"] != matrix_44x44["data"]["C"]:
        errors.append(f"교차 검증 실패: [18x18]의 C({matrix_18x18['data']['C']})와 [44x44]의 C({matrix_44x44['data']['C']})가 일치하지 않습니다.")
    
    return errors if errors else ["교차 검증이 성공적으로 완료되었습니다."]

# =========================================================
# 이미지 처리 및 바코드 검출 함수
# =========================================================

def split_image_for_detection(image):
    """이미지를 여러 영역으로 분할하여 바코드 인식률 향상"""
    width, height = image.size
    sections = []
    
    # 원본 이미지 추가
    sections.append(image)
    
    # 이미지를 상하좌우로 분할 (4분할)
    half_width = width // 2
    half_height = height // 2
    
    # 좌상단
    sections.append(image.crop((0, 0, half_width, half_height)))
    # 우상단
    sections.append(image.crop((half_width, 0, width, half_height)))
    # 좌하단
    sections.append(image.crop((0, half_height, half_width, height)))
    # 우하단
    sections.append(image.crop((half_width, half_height, width, height)))
    
    # 이미지를 가로로 3등분
    third_height = height // 3
    sections.append(image.crop((0, 0, width, third_height)))
    sections.append(image.crop((0, third_height, width, 2*third_height)))
    sections.append(image.crop((0, 2*third_height, width, height)))
    
    # 이미지를 세로로 3등분
    third_width = width // 3
    sections.append(image.crop((0, 0, third_width, height)))
    sections.append(image.crop((third_width, 0, 2*third_width, height)))
    sections.append(image.crop((2*third_width, 0, width, height)))
    
    return sections

def enhance_image_for_detection(image):
    """이미지 전처리를 통해 DataMatrix 인식률 향상 (개선 버전)"""
    # OpenCV로 이미지 처리
    img_array = np.array(image)
    
    results = [image]  # 원본 이미지 포함
    
    # 그레이스케일로 변환
    if len(img_array.shape) == 3:  # 컬러 이미지인 경우
        gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
    else:  # 이미 그레이스케일인 경우
        gray = img_array
    
    # 기본 처리: 노이즈 제거
    denoised = cv2.GaussianBlur(gray, (5, 5), 0)
    
    # 이미지 크기 조정 (확대)
    height, width = gray.shape
    scale_factors = [1.5, 2.0]
    for scale in scale_factors:
        resized = cv2.resize(gray, (int(width * scale), int(height * scale)), 
                            interpolation=cv2.INTER_CUBIC)
        results.append(Image.fromarray(resized))
    
    # 여러 이진화 방법 적용
    # 1. 적응형 이진화 (Adaptive Thresholding)
    binary_adaptive = cv2.adaptiveThreshold(denoised, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                                           cv2.THRESH_BINARY, 11, 2)
    results.append(Image.fromarray(binary_adaptive))
    
    # 2. Otsu 이진화
    _, binary_otsu = cv2.threshold(denoised, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    results.append(Image.fromarray(binary_otsu))
    
    # 3. 반전된 이진화 (바코드가 역상인 경우)
    _, binary_inv = cv2.threshold(denoised, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    results.append(Image.fromarray(binary_inv))
    
    # 대비 향상 (CLAHE)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    enhanced = clahe.apply(gray)
    results.append(Image.fromarray(enhanced))
    
    # CLAHE 적용 후 이진화
    _, clahe_binary = cv2.threshold(enhanced, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    results.append(Image.fromarray(clahe_binary))
    
    # 모폴로지 연산
    kernels = [(3, 3), (5, 5)]
    for k_size in kernels:
        kernel = np.ones(k_size, np.uint8)
        
        # 열림 연산 (침식 후 팽창) - 작은 노이즈 제거
        morph_open = cv2.morphologyEx(binary_adaptive, cv2.MORPH_OPEN, kernel)
        results.append(Image.fromarray(morph_open))
        
        # 닫힘 연산 (팽창 후 침식) - 작은 구멍 채우기
        morph_close = cv2.morphologyEx(binary_adaptive, cv2.MORPH_CLOSE, kernel)
        results.append(Image.fromarray(morph_close))
    
    # 엣지 검출
    edges = cv2.Canny(denoised, 50, 150)
    results.append(Image.fromarray(edges))
    
    # 선명화 필터
    sharpen_kernel = np.array([[-1, -1, -1], [-1, 9, -1], [-1, -1, -1]])
    sharpened = cv2.filter2D(gray, -1, sharpen_kernel)
    results.append(Image.fromarray(sharpened))
    
    return results

def detect_datamatrix(image):
    """이미지에서 DataMatrix 바코드 검출 (개선 버전)"""
    # 원본 이미지 전처리
    processed_images = enhance_image_for_detection(image)
    
    all_results = []
    
    # 원본 이미지의 다양한 처리 버전에서 바코드 검출 시도
    for img in processed_images:
        try:
            results = decode(img, timeout=5000, max_count=10)
            if results:
                all_results.extend(results)
        except Exception as e:
            st.warning(f"디코딩 중 오류 발생: {str(e)}")
    
    # 이미지가 복잡하거나 바코드가 작을 경우를 위해 이미지 분할 접근
    if len(all_results) < 2:  # 아직 두 개의 바코드를 찾지 못했다면
        # 이미지 분할
        sections = split_image_for_detection(image)
        
        # 각 섹션에 전처리 적용 및 바코드 검출
        for section in sections:
            # 섹션 전처리
            section_processed = enhance_image_for_detection(section)
            
            # 처리된 각 섹션에서 바코드 검출
            for img in section_processed:
                try:
                    results = decode(img, timeout=5000, max_count=10)
                    if results:
                        all_results.extend(results)
                except Exception as e:
                    continue  # 에러는 무시하고 계속 진행
    
    # 중복 제거 (바코드 값 기준)
    unique_data = set()
    decoded_data = []
    
    for result in all_results:
        try:
            data = result.data.decode('utf-8', errors='replace')
            if data not in unique_data:
                unique_data.add(data)
                decoded_data.append(data)
        except Exception as e:
            st.warning(f"결과 디코딩 중 오류 발생: {str(e)}")
    
    # 디버깅 정보 출력
    if decoded_data:
        st.info(f"총 {len(all_results)}개의 바코드 후보를 찾았고, 중복 제거 후 {len(decoded_data)}개 바코드 식별")
    
    return decoded_data

# =========================================================
# 파일 처리 함수
# =========================================================

@st.cache_data
def extract_images_from_file(uploaded_file):
    """업로드된 파일에서 이미지 추출"""
    file_content = uploaded_file.read()
    file_extension = uploaded_file.name.split('.')[-1].lower()
    
    # 이미지 파일인 경우 직접 처리
    if file_extension in ['jpg', 'jpeg', 'png', 'bmp', 'tiff']:
        image = Image.open(io.BytesIO(file_content))
        return {1: [image]}  # 슬라이드 1에 이미지 할당
    
    # PDF 또는 Office 파일인 경우 별도 처리
    if file_extension == 'pdf':
        try:
            with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_file:
                temp_file.write(file_content)
                temp_path = temp_file.name
            
            # PDF를 이미지로 변환 (외부 라이브러리 사용)
            import fitz  # PyMuPDF
            doc = fitz.open(temp_path)
            slide_images = {}
            
            for page_index in range(len(doc)):
                page = doc.load_page(page_index)
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # 더 높은 해상도
                img_data = pix.tobytes("png")
                image = Image.open(io.BytesIO(img_data))
                slide_num = page_index + 1
                if slide_num not in slide_images:
                    slide_images[slide_num] = []
                slide_images[slide_num].append(image)
            
            # 임시 파일 삭제
            os.unlink(temp_path)
            return slide_images
        except Exception as e:
            st.error(f"PDF 변환 중 오류: {str(e)}")
            return {}
    
    # Office 파일은 웹 애플리케이션에서 직접 처리하기 어려움을 안내
    if file_extension in ['xlsx', 'xls', 'pptx', 'ppt', 'docx', 'doc']:
        st.warning(f"{file_extension.upper()} 파일은 현재 웹 버전에서 완전히 지원되지 않습니다. 이미지 파일이나 PDF로 변환하여 업로드하세요.")
        return {}
    
    # 지원되지 않는 파일 형식
    st.error(f"지원되지 않는 파일 형식: {file_extension}")
    return {}

# =========================================================
# 결과 출력 함수
# =========================================================

def render_barcode_result(idx, data, result, matrix_type="44x44"):
    """바코드 검증 결과를 더 자세히 출력"""
    # 바코드 데이터 출력 - 앞부분과 뒷부분을 함께 표시
    if len(data) > 60:
        display_data = f"{data[:30]}...{data[-20:]}"
    else:
        display_data = data
    
    st.markdown(f"### 바코드 #{idx+1}")
    st.code(display_data, language="text")
    st.markdown(f"**{matrix_type} 매트릭스 형식으로 판단됩니다.**")
    
    if result["valid"]:
        st.success("✅ 매트릭스 형식이 올바릅니다 (전체 규격 검증 완료)")
        
        # 데이터 테이블로 표시
        data_dict = {}
        if matrix_type == "44x44":
            # 44x44 매트릭스 결과 출력
            data_dict = {
                "필드": ["제품 코드(C)", "품목 코드(I)", "창고 유형(W)", "트레이 번호(T)", 
                      "데이터 수(N)", "날짜(D)", "일련번호(S)", "바이너리 데이터(B)"],
                "값": [
                    result['data'].get('C', '없음'),
                    result['data'].get('I', '없음'),
                    result['data'].get('W', '없음'),
                    result['data'].get('T', '없음'),
                    result['data'].get('N', '없음'),
                    result['data'].get('D', '없음'),
                    result['data'].get('S', '없음'),
                    f"{result['data'].get('B', '')[:15]}...{result['data'].get('B', '')[-15:]} ({len(result['data'].get('B', ''))}자리)"
                    if len(result['data'].get('B', '')) > 30 else result['data'].get('B', '없음')
                ]
            }
        else:
            # 18x18 매트릭스 결과 출력
            data_dict = {
                "필드": ["모델 코드(M)", "품목 코드(I)", "제품 코드(C)", "제품 번호(P)"],
                "값": [
                    result['data'].get('M', '없음'),
                    result['data'].get('I', '없음'),
                    result['data'].get('C', '없음'),
                    result['data'].get('P', '없음')
                ]
            }
        
        # 데이터프레임으로 변환하여 표시
        df = pd.DataFrame(data_dict)
        st.table(df)
    else:
        if result["pattern_match"]:
            st.error("❌ 매트릭스의 기본 패턴은 일치하나 추가 검증에 실패했습니다:")
        else:
            st.error("❌ 매트릭스 형식이 맞지 않습니다:")
        
        for msg in result["errors"]:
            st.warning(f"- {msg}")

def render_summary_results(page_results):
    """페이지별 검증 결과 요약 테이블 출력"""
    st.markdown("## 📊 페이지별 검증 결과 요약")
    
    # 결과 요약 데이터 생성
    summary_data = []
    overall_valid = True
    issues_pages = []
    
    for page_num, result in sorted(page_results.items()):
        matrix_44x44 = "✅ 발견" if result["44x44_found"] else "❌ 없음"
        matrix_18x18 = "✅ 발견" if result["18x18_found"] else "❌ 없음"
        
        if result["44x44_found"] and result["44x44_valid"] and result["18x18_found"] and result["18x18_valid"]:
            validation = "✅ 통과"
        elif (not result["44x44_found"]) or (not result["18x18_found"]):
            validation = "❌ 실패 (미발견)"
            overall_valid = False
            issues_pages.append(page_num)
        elif (not result["44x44_valid"]) or (not result["18x18_valid"]):
            validation = "❌ 실패 (규격불일치)"
            overall_valid = False
            issues_pages.append(page_num)
        else:
            validation = "⚠️ 일부만 통과"
            overall_valid = False
            issues_pages.append(page_num)
        
        if result["44x44_found"] and result["18x18_found"] and result["cross_valid"]:
            cross_validation = "✅ 통과"
        elif not (result["44x44_found"] and result["18x18_found"]):
            cross_validation = "❓ 검증불가"
            if page_num not in issues_pages:
                issues_pages.append(page_num)
                overall_valid = False
        else:
            cross_validation = "❌ 실패"
            if page_num not in issues_pages:
                issues_pages.append(page_num)
                overall_valid = False
        
        summary_data.append({
            "페이지": page_num,
            "44x44 검출": matrix_44x44,
            "18x18 검출": matrix_18x18,
            "규격 검증": validation,
            "교차 검증": cross_validation
        })
    
    # 요약 테이블 출력
    if summary_data:
        df = pd.DataFrame(summary_data)
        st.table(df)
        
        # 최종 결과 출력
        st.markdown("### 🔍 최종 검증 결과:")
        if overall_valid:
            st.success("✅ 성공: 모든 페이지가 검증을 통과했습니다.")
        else:
            st.error(f"❌ 실패: {', '.join(map(str, sorted(issues_pages)))} 페이지에서 문제가 발견되었습니다.")
    else:
        st.warning("데이터가 없어 결과 요약을 생성할 수 없습니다.")

def render_format_help():
    """데이터 매트릭스 형식 정보 출력"""
    with st.expander("데이터매트릭스 형식 참고 정보"):
        st.markdown("### 올바른 데이터 매트릭스 형식")
        
        st.markdown("#### [44x44] 형식")
        st.code("CXXX.IYY.WZZ.TYY.NYYY.DYYYYMMDD.SYYY.BNNNNNNNN...", language="text")
        st.markdown("**예시:** `CAB1.I21.WLO.T10.N010.D20250317.S001.B000100020003000400050006000700080009001000000000000000000000000000000000000000000000000000000000000000000000000000000000.`")
        
        st.markdown("#### [18x18] 형식")
        st.code("MXXXX.IYY.CZZZ.PYYY.", language="text")
        st.markdown("**예시:** `MD213.I30.CSW1.P001.`")
        
        st.markdown("#### 요구사항")
        requirements = [
            "1. 각 페이지에는 두 종류(44x44, 18x18)의 데이터 매트릭스가 있어야 합니다.",
            "2. 44x44 매트릭스에서 C, I, T 식별자 값은 3자리의 (문자+숫자) 조합이어야 합니다.",
            "3. 44x44 매트릭스에서 W 식별자 값은 'LO' 또는 'SE'이어야 합니다.",
            "4. 44x44 매트릭스의 N 값과 B 식별자의 세트 수가 일치해야 합니다.",
            "5. 44x44 매트릭스의 B 식별자 숫자 세트는 오름차순이어야 합니다.",
            "6. 18x18 매트릭스의 C 값과 44x44 매트릭스의 C 값이 일치해야 합니다.",
            "7. 18x18 매트릭스의 I 값과 44x44 매트릭스의 I 값이 일치해야 합니다."
        ]
        for req in requirements:
            st.markdown(req)

# =========================================================
# 메인 애플리케이션 로직
# =========================================================

def main():
    # 사이드바 설정
    st.sidebar.header("설정")
    
    # 테스트 모드 설정
    test_mode = st.sidebar.checkbox("테스트 모드", value=False, help="샘플 데이터로 테스트할 때 선택하세요")
    
    if test_mode:
        st.sidebar.subheader("샘플 데이터")
        test_44x44 = st.sidebar.text_area(
            "44x44 매트릭스 샘플", 
            value="CSW1.I22.WLO.T10.N010.D20250317.S001.B0001000200030004000500060007000800090010000000000000000000000000000000000000000000000000000000000000000000000000000000.",
            height=100
        )
        test_18x18 = st.sidebar.text_area(
            "18x18 매트릭스 샘플",
            value="MR154.I22.CSW1.P001.",
            height=50
        )
        
        if st.sidebar.button("테스트 실행"):
            st.markdown("## 테스트 모드 결과")
            
            # 44x44 매트릭스 테스트
            st.markdown("### 44x44 매트릭스 테스트")
            result_44x44 = validate_44x44_matrix(test_44x44)
            render_barcode_result(0, test_44x44, result_44x44, "44x44")
            
            # 18x18 매트릭스 테스트
            st.markdown("### 18x18 매트릭스 테스트")
            result_18x18 = validate_18x18_matrix(test_18x18)
            render_barcode_result(0, test_18x18, result_18x18, "18x18")
            
            # 교차 검증 테스트
            st.markdown("### 교차 검증 테스트")
            cross_results = cross_validate_matrices(result_44x44, result_18x18)
            
            if "교차 검증이 성공적으로 완료되었습니다." in cross_results:
                st.success("✅ 교차 검증 성공")
            else:
                st.error("❌ 교차 검증 실패")
                for msg in cross_results:
                    st.warning(f"- {msg}")
    
    else:
        # 파일 업로드 섹션
        st.markdown("## 파일 업로드")
        st.markdown("이미지(.jpg, .png), PDF(.pdf) 파일을 업로드하세요.")
        uploaded_file = st.file_uploader("파일 선택", type=["jpg", "jpeg", "png", "pdf"])
        
        if uploaded_file is not None:
            # 파일 정보 표시
            st.success(f"파일명: {uploaded_file.name} ({uploaded_file.size} 바이트)")
            
            # 프로그레스 바 표시
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            status_text.text("이미지 추출 중...")
            slide_images = extract_images_from_file(uploaded_file)
            progress_bar.progress(25)
            
            if not slide_images:
                st.error("파일에서 이미지를 추출할 수 없습니다.")
                return
            
            status_text.text("바코드 검출 중...")
            
            # 페이지별 결과를 저장할 딕셔너리
            page_results = {}
            
            # 결과 섹션 시작
            st.markdown("## 분석 결과")
            
            # 각 슬라이드/페이지에서 모든 이미지 처리
            for slide_idx, (slide_num, images) in enumerate(sorted(slide_images.items())):
                # 프로그레스 업데이트
                progress_percent = 25 + (50 * slide_idx / len(slide_images))
                progress_bar.progress(int(progress_percent))
                status_text.text(f"페이지 {slide_num} 분석 중...")
                
                # 페이지 결과 탭 생성
                with st.expander(f"페이지/슬라이드 {slide_num} 분석 결과", expanded=(slide_idx == 0)):
                    st.markdown(f"### 페이지/슬라이드 {slide_num}")
                    st.markdown(f"슬라이드에서 추출된 이미지: {len(images)}개")
                    
                    # 페이지 결과 초기화
                    page_results[slide_num] = {
                        "44x44_found": False,
                        "18x18_found": False,
                        "44x44_valid": False,
                        "18x18_valid": False,
                        "cross_valid": False
                    }
                    
                    # 이 슬라이드에서 발견된 모든 바코드 저장
                    all_barcodes = []
                    
                    # 각 이미지에서 바코드 검출 및 통합
                    for img_idx, image in enumerate(images):
                        st.image(image, caption=f"이미지 #{img_idx+1} ({image.width} x {image.height})", width=400)
                        
                        # 이미지에서 데이터매트릭스 검출
                        import time
                        start_time = time.time()
                        decoded_data = detect_datamatrix(image)
                        end_time = time.time()
                        
                        if decoded_data:
                            st.success(f"이미지 #{img_idx+1}에서 {len(decoded_data)}개 바코드 발견 (검색 시간: {end_time - start_time:.2f}초)")
                            all_barcodes.extend(decoded_data)
                        else:
                            st.warning(f"이미지 #{img_idx+1}에서 바코드를 찾을 수 없습니다 (검색 시간: {end_time - start_time:.2f}초)")
                    
                    # 중복 제거
                    all_barcodes = list(set(all_barcodes))
                    
                    if not all_barcodes:
                        st.error(f"페이지/슬라이드 {slide_num}에서 DataMatrix 바코드를 찾을 수 없습니다.")
                        continue
                    
                    st.markdown(f"### 페이지/슬라이드 {slide_num}에서 총 {len(all_barcodes)}개의 DataMatrix 바코드를 발견했습니다.")
                    
                    # 바코드 데이터 저장 변수
                    data_44x44 = None
                    data_18x18 = None
                    result_44x44 = {"valid": False, "pattern_match": False}
                    result_18x18 = {"valid": False, "pattern_match": False}
                    
                    # 각 바코드 데이터 처리
                    for idx, data in enumerate(all_barcodes):
                        # 44x44 매트릭스 패턴 검사
                        if re.search(r'C[A-Za-z0-9]{3}[.,]I\d{2}[.,]W(?:LO|SE)[.,]', data):
                            # 이미 44x44 데이터가 있는 경우 기존 것이 유효한지 확인하고 결정
                            if data_44x44 is None or not result_44x44["valid"]:
                                result_44x44 = validate_44x44_matrix(data)
                                data_44x44 = data
                                # 개선된 출력 함수 사용
                                render_barcode_result(idx, data, result_44x44, "44x44")
                                
                                # 결과 업데이트
                                page_results[slide_num]["44x44_found"] = True
                                page_results[slide_num]["44x44_valid"] = result_44x44["valid"]
                        
                        # 18x18 매트릭스 패턴 검사
                        if re.search(r'M[A-Za-z0-9]{4}\.I\d{2}\.C[A-Za-z0-9]{3}\.', data):
                            # 이미 18x18 데이터가 있는 경우 기존 것이 유효한지 확인하고 결정
                            if data_18x18 is None or not result_18x18["valid"]:
                                result_18x18 = validate_18x18_matrix(data)
                                data_18x18 = data
                                # 개선된 출력 함수 사용
                                render_barcode_result(idx, data, result_18x18, "18x18")
                                
                                # 결과 업데이트
                                page_results[slide_num]["18x18_found"] = True
                                page_results[slide_num]["18x18_valid"] = result_18x18["valid"]
                    
                    # 페이지에 두 종류의 매트릭스가 모두 있는지 확인
                    if not data_44x44:
                        st.warning("⚠️ 경고: 이 페이지에서 44x44 매트릭스를 찾을 수 없습니다!")
                    
                    if not data_18x18:
                        st.warning("⚠️ 경고: 이 페이지에서 18x18 매트릭스를 찾을 수 없습니다!")
                    
                    # 교차 검증 수행
                    st.markdown("### 교차 검증 결과")
                    if data_44x44 and data_18x18:
                        if result_44x44["pattern_match"] and result_18x18["pattern_match"]:
                            cross_results = cross_validate_matrices(result_44x44, result_18x18)
                            
                            if "교차 검증이 성공적으로 완료되었습니다." in cross_results:
                                st.success("✅ 교차 검증 성공")
                                page_results[slide_num]["cross_valid"] = True
                            else:
                                st.error("❌ 교차 검증 실패")
                                for msg in cross_results:
                                    st.warning(f"- {msg}")
                        else:
                            st.error("❌ 교차 검증을 수행할 수 없습니다. 두 매트릭스 모두 기본 형식이 일치해야 합니다.")
                    else:
                        st.error("❌ 교차 검증을 수행할 수 없습니다. 페이지에 44x44와 18x18 매트릭스가 모두 필요합니다.")
            
            # 프로그레스 업데이트
            progress_bar.progress(75)
            status_text.text("결과 요약 생성 중...")
            
            # 모든 페이지 분석 후 결과 요약 출력
            render_summary_results(page_results)
            
            # 프로그레스 완료
            progress_bar.progress(100)
            status_text.text("분석 완료!")
            
            # 형식 도움말 출력
            render_format_help()
            
            # 결과 내보내기
            with st.expander("결과 내보내기"):
                # 요약 데이터를 CSV로 변환
                summary_data = []
                for page_num, result in sorted(page_results.items()):
                    summary_data.append({
                        "페이지": page_num,
                        "44x44_발견": result["44x44_found"],
                        "18x18_발견": result["18x18_found"],
                        "44x44_검증성공": result["44x44_valid"],
                        "18x18_검증성공": result["18x18_valid"],
                        "교차검증성공": result["cross_valid"]
                    })
                
                summary_df = pd.DataFrame(summary_data)
                csv = summary_df.to_csv(index=False)
                
                b64 = base64.b64encode(csv.encode()).decode()
                href = f'<a href="data:file/csv;base64,{b64}" download="datamatrix_검증결과.csv">CSV 파일로 결과 다운로드</a>'
                st.markdown(href, unsafe_allow_html=True)

if __name__ == "__main__":
    main()