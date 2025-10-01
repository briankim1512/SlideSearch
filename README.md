# SlideSearch

SlideSearch는 파워포인트(`.pptx`) 파일에서 슬라이드를 검색하고, 원하는 슬라이드만 선택하여 새로운 프레젠테이션으로 합치는 로컬 윈도우 프로그램입니다.

## 주요 기능

* **파일 불러오기 (Ingest)**: `.pptx` 파일을 불러와 각 슬라이드의 텍스트와 노트를 데이터베이스에 저장하고, 썸네일 이미지를 생성합니다.
* **검색**: 슬라이드 텍스트와 노트, 파일 제목, 수정 날짜를 기준으로 원하는 슬라이드를 검색할 수 있습니다.
* **슬라이드 합치기 (Stitch)**: 여러 프레젠테이션에서 선택한 슬라이드를 모아 새로운 파워포인트 파일을 생성합니다.

## 시스템 요구 사항

* 운영체제: Windows
* Microsoft PowerPoint 설치 필수 (COM 자동화를 사용함)
* Python 3.9 이상 권장

## 설치 방법

1. 저장소를 클론하거나 소스 코드를 다운로드합니다.

   ```bash
   git clone <repo-url>
   cd <repo-directory>
   ```
2. 가상환경 생성 및 활성화 (venv 사용).

   ```bash
   python -m venv venv
   venv\Scripts\activate   # Windows
   ```
3. 필요한 패키지 설치.

   ```bash
   pip install -r requirements.txt
   ```
4. 프로그램 실행.

   ```bash
   python main.py
   ```

## 배포 (Windows 실행 파일 만들기)

PyInstaller를 사용하여 독립 실행 파일을 만들 수 있습니다.

```bash
pyinstaller --windowed --onedir --add-data "assets;assets" --icon icon.ico main.py
```

* `--wibndowed`: 콘솔 없이 실행.
* `--add-data`: HTML, JS, CSS 등 `assets` 폴더를 포함합니다.

빌드 후 `dist/` 폴더 안에 실행 파일이 생성됩니다.

## 사용 방법

1. 프로그램 실행 후, 메인 화면에서 `.pptx` 파일을 불러옵니다.

   * **주의**: 파일을 불러온 뒤에는 원래 위치에서 이동하거나 이름을 변경하지 마세요. 슬라이드 합치기 기능이 정상 동작하지 않을 수 있습니다.
2. 검색창에서 원하는 키워드나 조건(제목, 수정 날짜 범위)을 입력하여 슬라이드를 찾습니다.
3. 필요한 슬라이드를 선택하고 **“선택한 장표로 새 PPT 만들기”** 버튼을 누르면 새로운 파워포인트 파일이 열립니다.
4. 새로 열린 파일은 직접 저장하면 됩니다.

## 데이터 저장 위치

* 데이터베이스: `assets/data/slides.db`
* 로그 파일: `assets/data/logs.log`
* 슬라이드 썸네일: `assets/data/<pptx_hash>/슬라이드N.PNG`

---
