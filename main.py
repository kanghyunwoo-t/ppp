import os
from vision_extractor import DocumentVisionExtractor
from html_to_word import HtmlToDocxConverter

# 설정값
GOOGLE_API_KEY = "" # 보안을 위해 비워둡니다. 배포 시 안전!
IMAGE_FILES = [
    "sample_page_1.jpg", 
    "sample_page_2.jpg"
]  # 처리할 이미지 파일 경로 리스트
OUTPUT_DOCX_FILE = "converted_document.docx"

def main():
    if not GOOGLE_API_KEY:
        print("경고: Google API 키를 입력해 주세요!")
        return

    print("=== [1단계] 이미지에서 HTML 추출 시작 ===")
    extractor = DocumentVisionExtractor(api_key=GOOGLE_API_KEY)
    
    # 여러 이미지를 순차 처리하여 HTML 문자열 하나로 묶음
    combined_html = extractor.process_multiple_images(IMAGE_FILES, delay_seconds=4)
    
    # 추출된 HTML 로깅 (선택 사항)
    with open("temp_extracted.html", "w", encoding="utf-8") as f:
        f.write(combined_html)
    print("-> HTML 구조 추출 완료 (임시 파일 temp_extracted.html 저장)\n")

    print("=== [2단계 & 3단계] HTML 정제 및 Word(.docx)로 변환 시작 ===")
    converter = HtmlToDocxConverter()
    converter.parse_and_convert(combined_html, OUTPUT_DOCX_FILE)
    
    print("\n🎉 모든 변환 작업이 완료되었습니다!")

if __name__ == "__main__":
    main()
