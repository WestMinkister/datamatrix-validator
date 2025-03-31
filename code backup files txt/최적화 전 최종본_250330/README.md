# DataMatrix 바코드 검증 도구

DataMatrix 바코드를 검색하고 검증하는 웹 애플리케이션입니다. PDF, PowerPoint(PPTX), Excel(XLSX) 파일에서 바코드를 자동으로 찾아 형식을 검증합니다.

## 주요 기능

- PDF, PowerPoint, Excel 파일에서 DataMatrix 바코드 자동 검출
- 44x44 매트릭스와 18x18 매트릭스 형식 검증
- 다양한 이미지 처리 기법을 통한 바코드 인식률 향상
- 두 바코드 간의 교차 검증
- 페이지/슬라이드별 상세 결과 및 요약 보고서
- 결과 보고서 다운로드

## 시스템 요구사항

이 애플리케이션은 다음 시스템 패키지를 필요로 합니다:

- libdmtx (바코드 디코딩)
- poppler-utils (PDF 처리)
- libreoffice (Office 파일 변환)

### Ubuntu/Debian에서 설치

```bash
sudo apt-get update
sudo apt-get install -y libdmtx0b libdmtx-dev
sudo apt-get install -y libreoffice
sudo apt-get install -y poppler-utils
```

## 로컬에서 실행하기

### 1. 저장소 클론

```bash
git clone https://github.com/yourusername/datamatrix-validator.git
cd datamatrix-validator
```

### 2. 필요한 패키지 설치

```bash
pip install -r requirements.txt
```

### 3. Streamlit 앱 실행

```bash
streamlit run app.py
```

### 4. 웹 브라우저에서 접속

브라우저에서 http://localhost:8501 으로 접속하면 애플리케이션을 사용할 수 있습니다.

## Docker로 실행하기

### 1. Docker 이미지 빌드

```bash
docker build -t datamatrix-validator .
```

### 2. Docker 컨테이너 실행

```bash
docker run -p 8501:8501 datamatrix-validator
```

### 3. 웹 브라우저에서 접속

브라우저에서 http://localhost:8501 으로 접속하면 애플리케이션을 사용할 수 있습니다.

## 바코드 형식 안내

### 44x44 매트릭스 형식

```
CXXX.IYY.WZZ.TYY.NYYY.DYYYYMMDD.SYYY.BNNNNNNNN...
```
예시: `CAB1.I21.WLO.T10.N010.D20250317.S001.B000100020003000400050006000700080009001000000000000000000000000000000000000000000000000000000000000000000000000000000000.`

### 18x18 매트릭스 형식

```
MXXXX.IYY.CZZZ.PYYY.
```
예시: `MD213.I30.CSW1.P001.`

## 기여하기

1. 이 저장소를 포크합니다.
2. 새 기능 브랜치를 만듭니다 (`git checkout -b feature/amazing-feature`)
3. 변경사항을 커밋합니다 (`git commit -m 'Add some amazing feature'`)
4. 브랜치를 푸시합니다 (`git push origin feature/amazing-feature`)
5. Pull Request를 생성합니다.

## 라이선스

이 프로젝트는 MIT 라이선스 하에 있습니다. 자세한 내용은 [LICENSE](LICENSE) 파일을 참조하세요.