import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# 네이버 SMTP 설정 (모듈 내부 상수)
SMTP_SERVER = "smtp.naver.com"
SMTP_PORT = 465

def send_email(user_id, user_pwd, to_email, subject, html_content):
    """
    네이버 SMTP를 통해 이메일을 전송하는 함수입니다.
    
    Args:
        user_id (str): 네이버 아이디 (앞부분만)
        user_pwd (str): 애플리케이션 비밀번호
        to_email (str): 받는 사람 이메일 주소
        subject (str): 메일 제목
        html_content (str): HTML 형식의 메일 본문
        
    Returns:
        bool: 전송 성공 시 True, 실패 시 False
    """
    
    # 메일 객체 생성
    msg = MIMEMultipart()
    msg['From'] = f"{user_id}@naver.com"
    msg['To'] = to_email
    msg['Subject'] = subject

    # 본문 추가
    msg.attach(MIMEText(html_content, 'html'))

    try:
        # SMTP 서버 연결 (SSL 보안 연결)
        server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
        server.login(user_id, user_pwd)
        
        # 메일 전송
        server.send_message(msg)
        server.quit()
        return True
        
    except Exception as e:
        print(f"❌ [mailer.py] 전송 오류 ({to_email}): {e}")
        return False

# 테스트용 코드 (이 파일을 직접 실행했을 때만 작동)
if __name__ == "__main__":
    print("이 파일은 모듈입니다. 다른 코드에서 import해서 사용하세요.")