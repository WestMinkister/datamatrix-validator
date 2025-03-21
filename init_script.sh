#!/bin/bash

# 운영체제 확인
if [[ "$OSTYPE" == "darwin"* ]]; then
    # macOS
    echo "macOS 환경 감지, Homebrew로 패키지 설치 시도..."
    # 홈브루로 필요한 라이브러리 설치
    brew install libdmtx
    brew install libreoffice
    brew install poppler
else
    # Linux/기타
    echo "Linux/기타 환경, apt로 패키지 설치 시도..."
    apt-get update
    apt-get install -y libdmtx0b libdmtx-dev libreoffice poppler-utils
fi

# pylibdmtx 재설치
pip install pylibdmtx --no-cache-dir