import streamlit as st
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as ReportLabImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import inch, mm
import os
from io import BytesIO
from datetime import datetime
from PIL import Image as PILImage

# ==============================================================================
# [관리자 설정 구역]
# ==============================================================================
# 1. 발신자 이메일 (요청하신 주소로 변경함)
# ※ 주의: korea.kr 메일은 보안 문제로 발송이 안 될 수 있습니다. 
#    실패 시 지메일(Gmail) 사용을 권장합니다.
ADMIN_EMAIL = "rlaekdgb@korea.kr" 

# 2. 구글 앱 비밀번호 (이메일 발송을 원할 때만 입력!)
# 입력하지 않으면(비워두면) 이메일 전송은 생략하고 PDF 다운로드만 가능합니다.
ADMIN_PASSWORD = ""  # 예: "abcd efgh ijkl mnop"

# 3. 수신자 이메일 (보고서를 받을 주소 - 고정됨)
TARGET_EMAIL = "rlaekdgb@korea.kr"
# ==============================================================================

# 페이지 기본 설정 (와이드 모드 적용)
st.set_page_config(page_title="함정 장비 상태 접수", layout="centered", page_icon="⚓")

# 한글 폰트 설정
FONT_NAME = 'NanumGothic'
FONT_FILE = 'NanumGothic.ttf'

def register_korean_font():
    """한글 폰트 등록 함수"""
    if os.path.exists(FONT_FILE):
        pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_FILE))
        return True
    return False

# ------------------------------------------------------------------
# 1. PDF 보고서 생성 함수 (HWP 양식 반영)
# ------------------------------------------------------------------
def generate_official_pdf(data, image_buffer=None):
    """HWP 양식에 맞춘 표 형태의 PDF 생성 (가로 방향 추천)"""
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), 
                            rightMargin=15*mm, leftMargin=15*mm, 
                            topMargin=15*mm, bottomMargin=15*mm)
    story = []
    
    has_font = register_korean_font()
    font_main = FONT_NAME if has_font else 'Helvetica'

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('Title', fontName=font_main, fontSize=20, alignment=1, spaceAfter=20)
    cell_style_center = ParagraphStyle('CellCenter', fontName=font_main, fontSize=10, alignment=1, leading=14)
    cell_style_left = ParagraphStyle('CellLeft', fontName=font_main, fontSize=10, alignment=0, leading=14)

    # (1) 문서 제목
    story.append(Paragraph("함정 장비 상태 현황", title_style))
    story.append(Spacer(1, 10))

    # (2) 표 데이터 구성
    headers = [
        Paragraph("연번", cell_style_center),
        Paragraph("함정(파출소)", cell_style_center),
        Paragraph("구분/기기", cell_style_center),
        Paragraph("제품명(model)", cell_style_center),
        Paragraph("지원 요청 항목", cell_style_center),
        Paragraph("담당자", cell_style_center),
        Paragraph("연락처", cell_style_center)
    ]

    dept_equip = f"{data['department']}-{data['equip_name']}"
    manager = f"{data['rank']} {data['name']}"
    
    row_data = [
        Paragraph("1", cell_style_center),
        Paragraph(data['ship_name'], cell_style_center),
        Paragraph(dept_equip, cell_style_center),
        Paragraph(data['model'], cell_style_center),
        Paragraph(data['action_req'], cell_style_left),
        Paragraph(manager, cell_style_center),
        Paragraph(data['phone'], cell_style_center)
    ]

    headers_sub = [
        Paragraph("기기 상태", cell_style_center),
        Paragraph("함정 점검 사항", cell_style_center), '', '', '', 
        Paragraph("문제점 사진", cell_style_center), ''
    ]

    photo_cell = Paragraph("사진 없음", cell_style_center)
    if image_buffer:
        try:
            img = PILImage.open(image_buffer)
            w, h = img.size
            aspect = h / float(w)
            display_width = 2.0 * inch
            display_height = display_width * aspect
            photo_cell = ReportLabImage(image_buffer, width=display_width, height=display_height)
        except:
            photo_cell = Paragraph("이미지 오류", cell_style_center)

    content_sub = [
        Paragraph(data['condition'], cell_style_center),
        Paragraph(data['status'].replace('\n', '<br/>'), cell_style_left), '', '', '',
        photo_cell, ''
    ]

    table_data = [headers, row_data, headers_sub, content_sub]
    col_widths = [15*mm, 30*mm, 45*mm, 35*mm, 60*mm, 30*mm, 35*mm]

    t = Table(table_data, colWidths=col_widths)

    tbl_style = [
        ('FONTNAME', (0, 0), (-1, -1), font_main),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('BACKGROUND', (0, 2), (-1, 2), colors.lightgrey),
        ('SPAN', (1, 2), (4, 2)), 
        ('SPAN', (5, 2), (6, 2)),
        ('SPAN', (1, 3), (4, 3)),
        ('SPAN', (5, 3), (6, 3)),
        ('MINROWHEIGHT', (3, 3), 50*mm), 
        ('VALIGN', (0, 3), (-1, 3), 'TOP'),
        ('PADDING', (0, 0), (-1, -1), 6),
    ]
    t.setStyle(TableStyle(tbl_style))
    
    story.append(t)
    
    story.append(Spacer(1, 10))
    story.append(Paragraph(f"작성일시: {datetime.now().strftime('%Y-%m-%d %H:%M')}", 
                           ParagraphStyle('Date', fontName=font_main, fontSize=9, alignment=2)))

    doc.build(story)
    buffer.seek(0)
    return buffer

# ------------------------------------------------------------------
# 2. 이메일 자동 전송 함수
# ------------------------------------------------------------------
def send_email_auto(data, pdf_buffer):
    try:
        msg = MIMEMultipart()
        msg['From'] = ADMIN_EMAIL
        msg['To'] = TARGET_EMAIL
        msg['Subject'] = f"[{data['ship_name']}] {data['equip_name']} 상태 현황 보고 ({data['name']})"

        body = f"""
        [함정 장비 상태 접수 알림]
        
        ■ 함정명: {data['ship_name']} ({data['department']})
        ■ 장비명: {data['equip_name']} (모델: {data['model']})
        ■ 작성자: {data['rank']} {data['name']}
        ■ 연락처: {data['phone']}
        ■ 지원 요청 항목: {data['action_req']}
        
        ※ 상세 내용은 첨부된 PDF 파일을 확인해 주세요.
        """
        msg.attach(MIMEText(body, 'plain'))

        filename = f"Report_{data['ship_name']}_{data['equip_name']}.pdf"
        part = MIMEApplication(pdf_buffer.read(), Name=filename)
        part.add_header('Content-Disposition', 'attachment', filename=filename)
        msg.attach(part)

        # SMTP 서버 설정 (korea.kr 사용 시 이 부분을 변경해야 할 수 있음)
        smtp_server = 'smtp.gmail.com' if 'gmail' in ADMIN_EMAIL else 'smtp.korea.kr' # 예시
        
        # 지메일이 아닐 경우 경고
        if 'gmail' not in ADMIN_EMAIL:
            print("주의: Gmail이 아닌 메일 주소가 설정되었습니다. SMTP 설정 확인이 필요합니다.")
            # 일단 Gmail 서버로 시도해봅니다 (실패 가능성 높음)
            smtp_server = 'smtp.gmail.com'

        server = smtplib.SMTP(smtp_server, 587)
        server.starttls()
        server.login(ADMIN_EMAIL, ADMIN_PASSWORD)
        server.sendmail(ADMIN_EMAIL, TARGET_EMAIL, msg.as_string())
        server.quit()
        
        return True, "전송 성공"
    except Exception as e:
        return False, str(e)

# ------------------------------------------------------------------
# 3. 메인 화면 UI (직원용)
# ------------------------------------------------------------------
def main():
    st.title("⚓ 함정 장비 상태 접수")
    st.markdown("아래 양식을 작성하여 제출하면 담당자 이메일로 자동 전송됩니다.")
    st.divider()

    with st.form("report_form"):
        st.subheader("1. 기본 정보")
        
        ship_list = [
            "1007함", "516함", "517함", "117정", "123정", "216정", 
            "P-22정", "P-55정", "P-62정", "P-76정", "P-98정", "P-115정", 
            "방제15호함", "방제26호정", "화학방제2함"
        ]
        
        col1, col2 = st.columns(2)
        with col1:
            ship_name = st.selectbox("함정(파출소)", ["선택하세요"] + ship_list)
        with col2:
            department = st.selectbox("소속 부서", ["항해", "안전", "통신", "기관"])
            
        col3, col4, col5 = st.columns(3)
        with col3:
            rank = st.selectbox("계급", ["순경", "경장", "경사", "경위", "경감", "경정"])
        with col4:
            name = st.text_input("성명", placeholder="홍길동")
        with col5:
            phone = st.text_input("연락처", placeholder="010-0000-0000")

        st.divider()
        st.subheader("2. 장비 정보 및 상태")
        
        col6, col7 = st.columns(2)
        with col6:
            equip_name = st.text_input("장비명 (기기)", placeholder="예: 주기관, 발전기")
        with col7:
            model = st.text_input("제품명 (Model)", placeholder="예: MTU 12V 1163TB93")
        
        action_req = st.text_input("지원 요청 항목", placeholder="예: NO.2 주기관, 가스켓 교체 필요 등")
        condition = st.text_input("기기 상태 (요약)", placeholder="예: 작동 불가, 누유, 소음 발생")
        status = st.text_area("함정 점검 사항 (상세 문제점)", height=150, 
                            placeholder="점검 결과 발견된 문제점, 고장 증상 등을 상세히 기록하세요.")
        
        uploaded_file = st.file_uploader("문제점 사진 첨부", type=['jpg', 'png', 'jpeg'])

        submitted = st.form_submit_button("접수 제출하기", type="primary")

    if submitted:
        if ship_name == "선택하세요" or not name or not equip_name or not status:
            st.error("⚠️ [함정명], [성명], [장비명], [점검사항]은 필수 입력 항목입니다.")
            return
        
        with st.spinner("보고서 생성 중..."):
            report_data = {
                'ship_name': ship_name,
                'department': department,
                'rank': rank,
                'name': name,
                'phone': phone,
                'equip_name': equip_name,
                'model': model,
                'action_req': action_req,
                'condition': condition,
                'status': status,
                'report_time': datetime.now().strftime("%Y-%m-%d %H:%M")
            }

            img_buffer = None
            if uploaded_file:
                img_buffer = BytesIO(uploaded_file.getvalue())

            # 1. PDF 생성
            pdf_result = generate_official_pdf(report_data, img_buffer)
            
            # 2. 이메일 전송 여부 결정
            # 비밀번호가 비어있으면 이메일 전송을 생략합니다.
            email_success = False
            email_msg = "비밀번호 미입력으로 전송 생략"

            if not ADMIN_PASSWORD:
                st.warning("⚠️ 관리자 비밀번호가 입력되지 않아 이메일 발송은 건너뜁니다.")
            else:
                # 비밀번호가 있을 때만 전송 시도
                with st.spinner("이메일 전송 중..."):
                    pdf_result.seek(0)
                    email_success, email_msg = send_email_auto(report_data, pdf_result)
            
            # 3. 결과 표시
            if email_success:
                st.success(f"✅ 접수 완료! {TARGET_EMAIL}로 보고서가 전송되었습니다.")
                st.balloons()
            elif ADMIN_PASSWORD:
                # 비밀번호를 넣었는데 실패한 경우
                st.error(f"❌ 이메일 전송 실패: {email_msg}")
                st.info("💡 보안 설정 문제일 수 있습니다. 아래 버튼으로 보고서를 직접 다운로드하세요.")
            
            # 4. 다운로드 버튼 (항상 표시)
            st.download_button(
                label="📄 생성된 보고서 다운로드 (PDF)",
                data=pdf_result.getvalue(),
                file_name=f"Report_{ship_name}_{equip_name}.pdf",
                mime="application/pdf"
            )

if __name__ == "__main__":
    main()
