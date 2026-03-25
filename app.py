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

st.title("📄 문서 변환기 : 이미지, pdf -> 워드")
st.markdown("##### (프린트물을 워드로 변환할 때 유용합니다)")

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
st.info("💡 **안내:** 이미지를 올리고 **[🚀 변환 시작하기]**를 누르시면, 작업 완료 후 하단에 **'미리보기'**가 나타납니다!")

col1, col2 = st.columns(2)
with col1:
    start_btn = st.button("🚀 변환 시작하기", use_container_width=True)
with col2:
    if st.button("🗑️ 파일 목록 비우기", use_container_width=True):
        st.session_state["uploader_key"] += 1
        st.rerun()

if start_btn:
    if not ordered_files:
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
                # 1단계: HTML 추출 (서버 금고에 숨겨둔 내 API 키를 안전하게 꺼내 쓰기)
                extractor = DocumentVisionExtractor(api_key=st.secrets["GOOGLE_API_KEY"])
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
                
                # [신규] 저작권/보안 차단된 페이지가 있다면 사용자에게 명시적으로 알림
                if extractor.blocked_pages:
                    blocked_msg = "**🚨 주의: 일부 페이지가 저작권 및 보안 정책에 의해 차단되었습니다!**\n\n"
                    blocked_msg += "아래의 파일들은 내용이 추출되지 않았습니다:\n"
                    for path, reason in extractor.blocked_pages:
                        # 사용자 친화적인 파일명으로 다듬기 (예: file_page_1.jpg -> file_page_1)
                        clean_name = os.path.basename(path).replace('_safe.jpg', '').replace('.jpg', '')
                        blocked_msg += f"- 📄 **{clean_name}** ({reason})\n"
                    blocked_msg += "\n*(차단된 페이지 외의 나머지 문서는 정상적으로 변환되어 워드 파일에 포함되었습니다.)*"
                    st.error(blocked_msg, icon="🚨")
                
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
