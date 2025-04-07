"""
데이터매트릭스 검증기 추가 기능 모듈
- 페이지간 검증 기능을 위한 함수들이 포함되어 있습니다.
"""

def validate_pages_p_values(page_results):
    """
    같은 I 값을 가진 페이지들에서 18x18 매트릭스의 P 값이 모두 다른지 검증
    """
    # I 값별로 페이지 그룹화
    i_value_groups = {}
    for page_num, result in page_results.items():
        if not (result["18x18_found"] and result["18x18_valid"]):
            continue
            
        # 18x18 매트릭스의 I 값과 P 값 추출
        i_value = result.get("18x18_data", {}).get("I", None)
        p_value = result.get("18x18_data", {}).get("P", None)
        
        if i_value and p_value:
            if i_value not in i_value_groups:
                i_value_groups[i_value] = []
            i_value_groups[i_value].append((page_num, p_value))
    
    # 각 I 값 그룹 내에서 P 값 중복 검사
    duplicates = {}
    for i_value, pages in i_value_groups.items():
        if len(pages) <= 1:  # 같은 I 값을 가진 페이지가 하나뿐이면 검사 필요 없음
            continue
            
        p_values = {}
        for page_num, p_value in pages:
            if p_value in p_values:
                # 중복 발견
                original_page = p_values[p_value]
                if page_num not in duplicates:
                    duplicates[page_num] = {}
                duplicates[page_num]["p_value_duplicate"] = True
                duplicates[page_num]["p_duplicate_with"] = original_page
            else:
                p_values[p_value] = page_num
    
    # 결과 업데이트
    for page_num, duplicate_info in duplicates.items():
        if page_num in page_results:
            page_results[page_num]["p_value_duplicate"] = True
            page_results[page_num]["p_duplicate_with"] = duplicate_info.get("p_duplicate_with")
            page_results[page_num]["p_duplicate_message"] = f"18x18 매트릭스의 P 값이 페이지 {duplicate_info.get('p_duplicate_with')}와(과) 중복됩니다."
    
    return page_results

def validate_pages_s_values(page_results):
    """
    같은 I 값을 가진 페이지들에서 44x44 매트릭스의 S 값이 모두 다른지 검증하고,
    S 값이 B 세트의 오름차순과 일치하는지 검증
    """
    # I 값별로 페이지 그룹화
    i_value_groups = {}
    for page_num, result in page_results.items():
        if not (result["44x44_found"] and result["44x44_valid"]):
            continue
            
        # 44x44 매트릭스의 I 값, S 값, B 값 추출
        i_value = result.get("44x44_data", {}).get("I", None)
        s_value = result.get("44x44_data", {}).get("S", None)
        b_value = result.get("44x44_data", {}).get("B", None)
        
        if i_value and s_value and b_value:
            if i_value not in i_value_groups:
                i_value_groups[i_value] = []
            i_value_groups[i_value].append((page_num, s_value, b_value))
    
    # 각 I 값 그룹 내에서 S 값 검증
    invalid_s_values = {}
    for i_value, pages in i_value_groups.items():
        if len(pages) <= 1:  # 같은 I 값을 가진 페이지가 하나뿐이면 S 값 중복 검사 필요 없음
            continue
            
        # S 값 중복 검사
        s_values = {}
        for page_num, s_value, _ in pages:
            if s_value in s_values:
                # 중복 발견
                original_page = s_values[s_value]
                if page_num not in invalid_s_values:
                    invalid_s_values[page_num] = {}
                invalid_s_values[page_num]["s_value_duplicate"] = True
                invalid_s_values[page_num]["s_duplicate_with"] = original_page
                invalid_s_values[page_num]["s_invalid_message"] = f"44x44 매트릭스의 S 값이 페이지 {original_page}와(과) 중복됩니다."
            else:
                s_values[s_value] = page_num
        
        # B 세트 오름차순에 따른 S 값 검증
        # 먼저 B 세트의 첫 번째 유효한 값(0000이 아닌)을 기준으로 정렬
        pages_with_b = []
        for page_num, s_value, b_value in pages:
            # B 세트 파싱 (4자리씩)
            b_sets = []
            for i in range(0, len(b_value), 4):
                if i+4 <= len(b_value):
                    b_set = b_value[i:i+4]
                    if b_set != '0000':
                        b_sets.append(int(b_set))
            
            if b_sets:  # 유효한 B 세트가 있는 경우
                first_valid_b = min(b_sets)  # 가장 작은 B 값
                pages_with_b.append((page_num, int(s_value), first_valid_b))
        
        # B 세트 기준으로 정렬
        pages_with_b.sort(key=lambda x: x[2])  # B 값 기준 정렬
        
        # S 값이 순차적(1씩 증가)인지 확인
        expected_s = 1
        for index, (page_num, s_value, _) in enumerate(pages_with_b):
            # S 값은 3자리 숫자여야 하고, 1부터 시작해서 순차적으로 증가해야 함
            expected_s_str = f"{expected_s:03d}"  # 예: 001, 002, 003, ...
            
            if s_value != int(expected_s_str):
                if page_num not in invalid_s_values:
                    invalid_s_values[page_num] = {}
                invalid_s_values[page_num]["s_value_out_of_order"] = True
                invalid_s_values[page_num]["s_invalid_message"] = f"44x44 매트릭스의 S 값이 {s_value}이지만, B 세트 오름차순 기준 {expected_s_str}이어야 합니다."
                invalid_s_values[page_num]["s_expected_value"] = expected_s_str
            
            expected_s += 1
    
    # 결과 업데이트
    for page_num, invalid_info in invalid_s_values.items():
        if page_num in page_results:
            page_results[page_num]["s_value_invalid"] = True
            # 메시지 저장
            if "s_invalid_message" in invalid_info:
                page_results[page_num]["s_invalid_message"] = invalid_info["s_invalid_message"]
            # 중복 정보 저장
            if "s_duplicate_with" in invalid_info:
                page_results[page_num]["s_duplicate_with"] = invalid_info["s_duplicate_with"]
            # 예상 값 저장
            if "s_expected_value" in invalid_info:
                page_results[page_num]["s_expected_value"] = invalid_info["s_expected_value"]
    
    return page_results

def process_page_validation(page_results, slide_images, page_tabs, session_state):
    """
    페이지간 검증 처리를 수행하는 통합 함수
    
    Parameters:
    -----------
    page_results : dict
        각 페이지의 검증 결과를 담은 딕셔너리
    slide_images : dict
        각 슬라이드/페이지의 이미지 정보를 담은 딕셔너리
    page_tabs : list
        각 페이지 탭 객체의 리스트
    session_state : object
        StreamLit 세션 상태 객체
        
    Returns:
    --------
    dict : 업데이트된 page_results 딕셔너리
    """
    import streamlit as st
    
    if not page_results:
        return page_results
        
    # 1. 18x18의 P 값 중복 검사 (18x18 검증 모드일 때만)
    if session_state.validation_mode in ["both", "18x18"]:
        page_results = validate_pages_p_values(page_results)
    
    # 2. 44x44의 S 값 검증 (중복 및 B 순서 확인) (44x44 검증 모드일 때만)
    if session_state.validation_mode in ["both", "44x44"]:
        page_results = validate_pages_s_values(page_results)
    
    # 페이지간 유효성 검사 결과 표시
    for page_num, result in page_results.items():
        # P 값 중복 관련 오류 표시 (18x18 검증 모드일 때만)
        if session_state.validation_mode in ["both", "18x18"] and result.get("p_value_duplicate", False):
            with page_tabs[list(sorted(slide_images.keys())).index(page_num)]:
                st.error(f"\u274c 페이지간 검증 오류: {result.get('p_duplicate_message')}")
        
        # S 값 관련 오류/경고 표시 (44x44 검증 모드일 때만)
        if session_state.validation_mode in ["both", "44x44"]:
            # S 값 중복은 오류로 표시
            if result.get("s_value_invalid", False) and result.get("s_duplicate_with", None):
                with page_tabs[list(sorted(slide_images.keys())).index(page_num)]:
                    st.error(f"\u274c 페이지간 검증 오류: {result.get('s_invalid_message')}")
            # S 값 순서나 B 연속성 문제는 경고로 표시
            elif result.get("s_value_warning", False) or (result.get("s_value_invalid", False) and not result.get("s_duplicate_with", None)):
                with page_tabs[list(sorted(slide_images.keys())).index(page_num)]:
                    st.warning(f"\u26a0\ufe0f 페이지간 검증 확인 필요: {result.get('s_invalid_message')}")
    
    return page_results
