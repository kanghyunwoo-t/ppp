import os
import tempfile
import streamlit as st
import streamlit.components.v1 as components
from PIL import Image
from vision_extractor import DocumentVisionExtractor
from html_to_word import HtmlToDocxConverter

# 페이지 기본 설정
st.set_page_config(page_title="문서 변환기", layout="centered")

st.title("📄 문서 구조 보존형 디지털 변환기 (✨업데이트됨)")
st.write("이미지 파일을 아래에 **드래그 앤 드롭**하거나 클릭해서 업로드하세요.")

# 넓고 큼직한 파일 업로드 칸을 만들기 위한 커스텀 CSS 적용
st.markdown("""
<style>
    /* Streamlit 버전에 따른 호환성 추가 */
    [data-testid="stFileUploadDropzone"], [data-testid="stFileUploaderDropzone"] {
        min-height: 400px !important;
        padding: 120px 20px !important; /* 위아래 여백을 대폭 늘려서 압도적으로 거대하게 만듦 */
        border: 4px dashed #a0aab8 !important;
        background-color: #f4f6f9 !important;
        border-radius: 20px !important;
    }
    [data-testid="stFileUploadDropzone"] svg {
        width: 80px !important;
        height: 80px !important;
    }
    [data-testid="stFileUploadDropzone"]:hover {
        border-color: #0068c9 !important;
        background-color: #e8f1f8 !important;
    }
</style>
""", unsafe_allow_html=True)

# 파일 드래그 앤 드롭 업로드 칸 (여러 장 지원)
uploaded_files = st.file_uploader("이미지 파일 업로드", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)

# 업데이트 기능 안내 (사용자 혼란 방지용)
st.info("💡 **안내:** 이미지를 올리고 **[🚀 변환 시작하기]**를 누르시면, 작업 완료 후 하단에 **'미리보기'**와 **'API 요금 계산기'**가 나타납니다!")

# 사용자로부터 API 키 직접 입력받기 (보안 강화)
user_api_key = st.text_input("🔑 본인의 Gemini API Key를 입력하세요", type="password", help="Google AI Studio에서 발급받은 API 키를 입력하세요. 입력하신 키는 서버에 저장되지 않고 변환 즉시 폐기됩니다.")

if st.button("🚀 변환 시작하기"):
    if not user_api_key:
        st.warning("API 키를 입력해 주세요!")
    elif not uploaded_files:
        st.warning("변환할 이미지 파일을 업로드해 주세요!")
    else:
        with st.spinner("AI가 이미지를 분석하고 워드로 변환 중입니다... (시간이 조금 걸릴 수 있습니다)"):
            # 업로드된 파일을 처리하기 위해 임시 폴더에 저장
            temp_dir = tempfile.mkdtemp()
            image_paths = []
            
            for file in uploaded_files:
                # 갤럭시 모션 포토 등 특수 포맷(MPO) 에러를 방지하기 위해 순수 일반 JPEG로 강제 변환
                path = os.path.join(temp_dir, f"{os.path.splitext(file.name)[0]}_safe.jpg")
                img = Image.open(file)
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                img.save(path, format='JPEG')
                image_paths.append(path)
            
            try:
                # 1단계: HTML 추출
                extractor = DocumentVisionExtractor(api_key=user_api_key)
                st.info(f"🤖 **자동 적용된 AI 모델:** `{extractor.model_name}`")
                
                # 답답함을 해소하기 위한 실시간 진행 바(Progress Bar) 추가
                progress_text = st.empty()
                progress_bar = st.progress(0)
                
                def update_progress(current, total):
                    progress_text.text(f"🏃 {total}장 중 {current}번째 이미지 분석 중... (AI가 꼼꼼히 읽고 있습니다!)")
                    progress_bar.progress(current / total)
                    
                # 대기 시간을 4초에서 1.5초로 확 줄여서 속도 향상
                combined_html = extractor.process_multiple_images(image_paths, delay_seconds=1.5, progress_callback=update_progress)
                
                progress_text.text("✨ 거의 다 왔습니다! 워드 문서로 조립하는 중...")
                # 2, 3단계: 워드 변환
                output_path = os.path.join(temp_dir, "converted.docx")
                converter = HtmlToDocxConverter()
                converter.parse_and_convert(combined_html, output_path)
                
                st.success("🎉 변환이 완료되었습니다! 아래 버튼을 눌러 다운로드하세요.")
                
                # 비용 계산 로직 추가
                in_tokens = extractor.total_input_tokens
                out_tokens = extractor.total_output_tokens
                is_pro = "pro" in extractor.model_name.lower()
                in_price = 3.50 if is_pro else 0.35
                out_price = 10.50 if is_pro else 1.05
                
                est_cost_usd = (in_tokens / 1_000_000) * in_price + (out_tokens / 1_000_000) * out_price
                est_cost_krw = est_cost_usd * 1350 # 환율 대략 1350원 기준
                
                st.info(f"💰 **이번 변환 예상 API 비용:** 약 ${est_cost_usd:.4f} (한화 약 {int(est_cost_krw)}원)\n\n"
                        f"📊 **소모된 토큰:** 이미지 분석(입력) {in_tokens:,}개 / 결과 생성(출력) {out_tokens:,}개")
                
                # 확실하게 보이도록 미리보기 상자 밖으로 빼기 (expander 제거)
                st.markdown("### 👀 추출된 내용 미리보기")
                with st.container():
                    # 표 테두리와 구분선이 잘 보이도록 기본 스타일(CSS) 적용
                    preview_style = """
                    <style>
                        body { font-family: 'Malgun Gothic', sans-serif; line-height: 1.6; color: #333; padding: 10px; }
                        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
                        th, td { border: 1px solid #777; padding: 8px; text-align: left; }
                        hr.page-break { border: 0; border-top: 3px dashed #bbb; margin: 30px 0; }
                    </style>
                    """
                    components.html(preview_style + combined_html, height=500, scrolling=True)
                
                # 완료 시 다운로드 버튼 생성
                with open(output_path, "rb") as f:
                    st.download_button(
                        label="📥 워드(.docx) 파일 다운로드",
                        data=f,
                        file_name="converted_document.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            except Exception as e:
                st.error(f"오류가 발생했습니다: {e}")