import time
import google.generativeai as genai
from PIL import Image

class DocumentVisionExtractor:
    def __init__(self, api_key: str):
        # Gemini API 초기화 (HTML 구조 생성에 탁월한 성능)
        genai.configure(api_key=api_key)
        
        # 1. API 키로 사용 가능한 모든 모델 검색
        available_models = [
            m.name.replace("models/", "") 
            for m in genai.list_models() 
            if 'generateContent' in m.supported_generation_methods
        ]
        
        # 2. 성능이 좋은 순서대로 우선순위 지정
        preferences = [
            "gemini-1.5-pro",
            "gemini-1.5-pro-latest",
            "gemini-1.5-flash",
            "gemini-1.5-flash-latest",
            "gemini-pro-vision",
            "gemini-1.0-pro-vision-latest"
        ]
        
        self.model_name = None
        for pref in preferences:
            if pref in available_models:
                self.model_name = pref
                break
        
        # 3. 우선순위 목록에 없으면 비전(Vision) 지원 모델 중 아무거나 자동 선택
        if not self.model_name:
            for m in available_models:
                if 'vision' in m or '1.5' in m:
                    self.model_name = m
                    break
                    
        if not self.model_name:
            raise Exception(f"이미지 분석을 지원하는 모델을 찾을 수 없습니다.\n현재 API 키로 사용 가능한 모델: {available_models}")
            
        self.model = genai.GenerativeModel(self.model_name)
        self.total_input_tokens = 0
        self.total_output_tokens = 0
        
    def extract_html_from_image(self, image_path: str) -> str:
        """단일 이미지에서 HTML 구조를 추출합니다."""
        print(f"[{image_path}] 분석 중...")
        img = Image.open(image_path)
        
        prompt = """
        당신은 문서 구조화 전문가입니다. 첨부된 문서 이미지를 분석하여 완벽한 시맨틱 HTML로 변환하세요.
        - 제목(<h1>~<h3>), 본문(<p>), 리스트(<ul>, <li>)를 정확히 사용하세요.
        - 표(Table)가 있다면 <table>, <tr>, <th>, <td> 태그를 사용하고, 병합된 셀이 있다면 반드시 colspan과 rowspan 속성을 정확히 계산하여 포함하세요.
        - 수학 수식이나 특수 기호는 있는 그대로 텍스트로 보존하세요.
        - CSS 스타일이나 <html>, <body> 태그는 제외하고 내부 HTML 코드만 순수하게 출력하세요. Markdown 래퍼(```html 등) 없이 출력하세요.
        """
        
        response = self.model.generate_content([prompt, img])
        try:
            if response.usage_metadata:
                self.total_input_tokens += response.usage_metadata.prompt_token_count
                self.total_output_tokens += response.usage_metadata.candidates_token_count
        except Exception:
            pass
        text = response.text.strip()
        
        # AI가 마크다운(```html 등)으로 감싸서 출력한 경우 불필요한 텍스트 강제 제거
        if text.startswith("```html"):
            text = text[7:]
        elif text.startswith("```"):
            text = text[3:]
        if text.endswith("```"):
            text = text[:-3]
            
        return text.strip()

    def process_multiple_images(self, image_paths: list, delay_seconds: float = 3.0, progress_callback=None) -> str:
        """여러 장의 이미지를 순차 처리하여 API 제한을 방지하고 하나의 HTML로 통합합니다."""
        combined_html = ""
        total = len(image_paths)
        
        for idx, path in enumerate(image_paths):
            if progress_callback:
                progress_callback(idx + 1, total)
            try:
                html_chunk = self.extract_html_from_image(path)
                
                # 첫 페이지가 아니면 페이지 나누기(Page Break)용 태그 추가
                if idx > 0:
                    combined_html += "\n<hr class=\"page-break\">\n"
                    
                combined_html += f"\n<!-- Page {idx + 1} -->\n" + html_chunk
                
                # 마지막 이미지가 아니면 API 호출 제한 방지를 위해 대기
                if idx < total - 1:
                    print(f"API Rate Limit 방지를 위해 {delay_seconds}초 대기합니다...")
                    time.sleep(delay_seconds)
            except Exception as e:
                raise Exception(f"AI가 이미지를 분석하는 중 오류가 발생했습니다 ({path}). 상세 에러: {e}")
                
        return combined_html
