import os
import tempfile
import streamlit as st
import streamlit.components.v1 as components
import fitz  # PyMuPDF (PDF 처리용)
from PIL import Image
from streamlit_sortables import sort_items
from vision_extractor import DocumentVisionExtractor
from html_to_word import HtmlToDocxConverter

# 페이지 기본 설정
st.set_page_config(page_title="문서 변환기", layout="centered")

st.title("📄 문서 구조 보존형 디지털 변환기 (✨업데이트됨)")
st.write("이미지 및 **PDF 파일**을 아래에 **드래그 앤 드롭**하거나 클릭해서 업로드하세요.")

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

# 업로드 칸을 비우기 위한 세션 상태 키 초기화
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

# 파일 드래그 앤 드롭 업로드 칸 (여러 장 지원)
uploaded_files = st.file_uploader("이미지 및 PDF 파일 업로드", type=['png', 'jpg', 'jpeg', 'pdf'], accept_multiple_files=True, key=str(st.session_state["uploader_key"]))

# --- [추가된 기능 1] 업로드된 파일 순서 확인 및 드래그(선택) 변경 ---
ordered_files = []
if uploaded_files:
    st.markdown("### 🔄 작업 순서 확인 및 변경")
    if len(uploaded_files) > 1:
        st.caption("💡 **순서가 맞지 않나요?** 아래 목록에서 마우스로 직접 항목을 **드래그(위아래로 끌기)**하여 순서를 변경해 보세요!")
        
        # 동일한 파일명이 있을 수 있으므로 번호를 붙여 고유하게 만듦
        file_dict = {f"{i+1}. {f.name}": f for i, f in enumerate(uploaded_files)}
        
        # 드래그 앤 드롭 방식의 정렬 UI 제공
        selected_order = sort_items(list(file_dict.keys()))
        if not selected_order: # 로딩 중일 때의 안전장치
            selected_order = list(file_dict.keys())
            
        ordered_files = [file_dict[k] for k in selected_order]
    else:
        ordered_files = uploaded_files

    # 화면에 작은 썸네일(미리보기) 띄워주기 (최대 5열)
    if ordered_files:
        cols = st.columns(min(len(ordered_files), 5))
        for idx, file in enumerate(ordered_files):
            col = cols[idx % 5]
            with col:
                if file.name.lower().endswith(('.png', '.jpg', '.jpeg')):
                    st.image(file, caption=f"[{idx+1}번]", use_container_width=True)
                    file.seek(0)  # 이미지를 읽고 난 후, 변환 작업을 위해 버퍼 위치를 초기화(매우 중요!)
                else:
                    st.info(f"📄 [{idx+1}번] PDF")

# 업데이트 기능 안내 (사용자 혼란 방지용)
st.info("💡 **안내:** 이미지를 올리고 **[🚀 변환 시작하기]**를 누르시면, 작업 완료 후 하단에 **'미리보기'**와 **'API 요금 계산기'**가 나타납니다!")

# 사용자로부터 API 키 직접 입력받기 (보안 강화)
user_api_key = st.text_input("🔑 본인의 Gemini API Key를 입력하세요", type="password", help="Google AI Studio에서 발급받은 API 키를 입력하세요. 입력하신 키는 서버에 저장되지 않고 변환 즉시 폐기됩니다.")

col1, col2 = st.columns(2)
with col1:
    start_btn = st.button("🚀 변환 시작하기", use_container_width=True)
with col2:
    if st.button("🗑️ 파일 목록 비우기", use_container_width=True):
        st.session_state["uploader_key"] += 1
        st.rerun()

if start_btn:
    if not user_api_key:
        st.warning("API 키를 입력해 주세요!")
    elif not ordered_files:
        st.warning("변환할 이미지나 PDF 파일을 업로드하고 순서를 지정해 주세요!")
    else:
        # UI를 가장 먼저 화면에 띄워서 멈춰있는 느낌 원천 차단
        st.markdown("---")
        progress_title = st.empty()
        progress_bar = st.progress(0)
        progress_status = st.empty()
        st.markdown("---")
        
        temp_dir = tempfile.mkdtemp()
        image_paths = []
        
        progress_title.markdown("### ⚙️ 파일 변환 준비 중...")
        
        for f_idx, file in enumerate(ordered_files):
            progress_status.info(f"🏃 '{file.name}' 파일을 AI가 읽을 수 있도록 최적화하고 있습니다...")
            ext = os.path.splitext(file.name)[1].lower()
            
            if ext == '.pdf':
                pdf_document = fitz.open(stream=file.read(), filetype="pdf")
                for page_num in range(len(pdf_document)):
                    page = pdf_document.load_page(page_num)
                    # 해상도를 150으로 살짝 낮춰 분할 속도 2배 향상
                    pix = page.get_pixmap(dpi=150) 
                    img_path = os.path.join(temp_dir, f"{os.path.splitext(file.name)[0]}_page_{page_num+1}.jpg")
                    pix.save(img_path)
                    image_paths.append(img_path)
            else:
                path = os.path.join(temp_dir, f"{os.path.splitext(file.name)[0]}_safe.jpg")
                img = Image.open(file)
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                # 용량 다이어트(quality=85)를 통해 구글 API 전송 속도 대폭 향상
                img.save(path, format='JPEG', optimize=True, quality=85)
                image_paths.append(path)
            
            progress_bar.progress((f_idx + 1) / len(ordered_files))
        
        try:
            with st.spinner("AI 분석을 시작합니다..."):
                # 1단계: HTML 추출
                extractor = DocumentVisionExtractor(api_key=user_api_key)
                st.info(f"🤖 **자동 적용된 AI 모델:** `{extractor.model_name}`")
                
                progress_bar.progress(0) # AI 분석용으로 게이지 리셋
                
                def update_progress(current, total):
                    percent = int((current / total) * 100)
                    progress_title.markdown(f"### 🚀 분석 진행률: {percent}% ({current}/{total}장 완료)")
                    progress_bar.progress(current / total)
                    progress_status.info(f"🏃 현재 {current}번째 페이지 처리가 완료되었습니다. (병렬 처리 중)")
                    
                combined_html = extractor.process_multiple_images(image_paths, delay_seconds=1.5, progress_callback=update_progress)
                
                progress_title.markdown("### ✨ 거의 다 왔습니다! 워드 문서로 조립하는 중...")
                progress_status.empty()
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
