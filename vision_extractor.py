import time
import google.generativeai as genai
from PIL import Image
import concurrent.futures
import threading

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
        
        # 2. 유료 사용자를 위한 최고 성능(Pro) 모델을 다시 최우선으로 지정
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
        self._lock = threading.Lock() # 다중 스레드에서 토큰을 안전하게 더하기 위한 자물쇠
        
    def extract_html_from_image(self, image_path: str) -> str:
        """단일 이미지에서 HTML 구조를 추출합니다."""
        print(f"[{image_path}] 분석 중...")
        img = Image.open(image_path)
        
        prompt = """
        당신은 문서 구조화 및 표(Table) 복원 전문가입니다. 첨부된 문서 이미지를 분석하여 완벽한 시맨틱 HTML로 변환하세요.
        특히, 표를 분석할 때 다음 규칙을 뼈대로 삼아 엄격하게 지키세요:
        1. [그리드 계산] 표의 전체 행(Row)과 열(Column)의 개수를 시각적으로 먼저 완벽히 파악하세요.
        2. [병합 검증] 병합된 셀(colspan, rowspan)이 있다면, 각 줄(<tr>)의 총 칸 수가 전체 열의 개수와 논리적으로 완벽히 일치하도록 꼼꼼하게 검증하며 태그를 작성하세요.
        3. 투명한 테두리나 배경색으로만 구분된 암묵적인 표도 놓치지 마세요.
        4. 복잡하게 얽힌 표라도 '중첩 표(표 안의 표)'를 만들지 말고, 단일 표 내에서의 colspan/rowspan만으로 깔끔하게 구현하세요.
        
        일반 텍스트 구조화 규칙:
        - 제목(<h1>~<h4>), 본문(<p>), 리스트(<ul>, <li>)를 상황에 맞게 정확히 사용하세요.
        - 수식 및 특수 기호는 누락 없이 있는 그대로 보존하세요.
        - <html>, <body> 등의 컨테이너나 CSS를 절대 포함하지 말고 순수 HTML 태그만 출력하세요. Markdown 표기(```html)도 절대 쓰지 마세요.
        """
        
        response = self.model.generate_content([prompt, img])
        try:
            if response.usage_metadata:
                with self._lock: # 여러 스레드가 동시에 토큰을 더할 때 충돌하지 않도록 보호
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
        results = [None] * total
        completed = 0
        
        def _process_single(idx, path):
            """단일 이미지를 처리하는 독립된 작업 함수"""
            for attempt in range(3): # 최대 3번 끈질기게 재시도 (안정성 강화)
                try:
                    chunk = self.extract_html_from_image(path)
                    return idx, chunk, None
                except Exception as e:
                    err_msg = str(e)
                    # 속도가 너무 빨라 429(Too Many Requests)가 뜨면 잠시 대기 후 재시도
                    if "429" in err_msg or "Quota" in err_msg or "exhausted" in err_msg.lower():
                        time.sleep(2 * (attempt + 1)) 
                        continue
                    return idx, None, err_msg
            return idx, None, "API 호출 제한(429) 에러가 계속 발생했습니다."
                
        # [핵심] 유료 API의 압도적인 트래픽 허용량을 활용하여 일꾼 10명이 동시 처리 (리미트 해제)
        with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
            futures = [executor.submit(_process_single, i, p) for i, p in enumerate(image_paths)]
            
            # 완료되는 작업부터 화면(프로그래스 바)에 즉시 반영
            for future in concurrent.futures.as_completed(futures):
                idx, chunk, err = future.result()
                completed += 1
                
                if progress_callback:
                    progress_callback(completed, total)
                    
                if err:
                    raise Exception(f"AI 분석 중 오류 발생 ({image_paths[idx]}): {err}")
                    
                # 완료된 순서가 뒤죽박죽이어도 배열에 정확한 자기 번호(idx) 자리에 저장
                results[idx] = chunk
                
        # 모든 병렬 작업이 끝난 후 1페이지부터 순서대로 HTML 조합
        for idx, chunk in enumerate(results):
            if idx > 0:
                combined_html += "\n<hr class=\"page-break\">\n"
            combined_html += f"\n<!-- Page {idx + 1} -->\n" + chunk
                
        return combined_html
