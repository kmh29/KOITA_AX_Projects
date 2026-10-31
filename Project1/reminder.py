"""정부사업 리마인더 — DB 및 로그 저장 없는 MVP 버전.
실행할 때마다 schedule.xlsx의 알림 대상 일정을 발송하고 종료합니다.
중복 발송 방지가 없으므로 같은 날 재실행하면 같은 메일이 다시 발송됩니다.
콘솔 print 출력만 유지하며 로그 파일이나 DB는 생성하지 않습니다.
"""


"""정부사업 리마인더 — DB 및 로그 저장 없는 MVP 버전.
실행할 때마다 schedule.xlsx의 알림 대상 일정을 발송하고 종료합니다.
중복 발송 방지가 없으므로 같은 날 재실행하면 같은 메일이 다시 발송됩니다.
콘솔 print 출력만 유지하며 로그 파일이나 DB는 생성하지 않습니다.
"""


import os
import smtplib
import ssl
from pathlib import Path
import pandas as pd
from dotenv import load_dotenv
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from openai import OpenAI

# .py에서는 파일 폴더, 주피터에 붙여 넣으면 현재 작업 폴더 기준
BASE_DIR = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
load_dotenv(BASE_DIR / "reminder_info.env", override=True)

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-6-luna")

GMAIL_ADDRESS = os.getenv("GMAIL_ADDRESS", "").strip()
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD", "").replace(" ", "").strip()

REMINDER_DAYS = [30, 14, 7, 5, 3, 1]
EXCEL_PATH = BASE_DIR / "schedule.xlsx"
client = OpenAI(api_key=OPENAI_API_KEY)



# 1. GPT 메일 생성 (기존 방식 유지)
def generate_email(event, days_left):
    event_date = event["기준일"].strftime("%Y년 %m월 %d일")
    prompt = f"""
너는 정부 R&D 사업 운영 담당자를 지원하는 업무비서야.

아래 일정을 담당자에게 알려주는 내부 업무메일을 다음과 같이 작성해줘.

[일정정보]
사업명: {event["사업명"]}
일정유형: {event["일정유형"]}
일정명: {event["일정명"]}
일정일: {event_date}
남은기간: D-{days_left}
참고사항: {event["참고사항"]}

작성원칙:
- 전문기관이 연구수행기관 담당자에게 메일을 보내는 것처럼 정중하게 작성할 것. (500자 이내)
- 첫 줄은 반드시 '제목: 간결한 메일 제목'으로 작성
- 다음 줄에 '본문:'을 작성
- 사업명, 일정유형, 일정명, 일정일, 남은기간, 참고사항은 본문 맨 위에 제시하고, 그 아래에 안내문 작성
- 마크다운 표나 HTML 태그 없이 일반 텍스트로 작성
- 입력된 정보에 없는 사실을 임의로 만들지 말 것
- 불확실한 것은 '확인 필요'라고 표시
- 아래 형식으로 작성

제목: ...
본문:
...
"""
    response = client.responses.create(
        model=OPENAI_MODEL,
        input=prompt,
        extra_headers={"Accept-Encoding": "identity"}
    )

    text = response.output_text.strip()

    # 아주 단순한 제목/본문 분리
    if "본문:" in text:
        subject_part, body = text.split("본문:", 1)
        subject = subject_part.replace("제목:", "").strip()
    else:
        subject = f"[D-{days_left}] {event['사업명']} {event['일정명']} 안내"
        body = text

    return subject, body.strip()


# 2. Gmail 발송
def send_email(recipient, subject, body):

    message = MIMEMultipart()
    message["From"] = GMAIL_ADDRESS
    message["To"] = recipient
    message["Subject"] = subject

    message.attach(
        MIMEText(body, "plain", "utf-8")
    )


    with smtplib.SMTP_SSL(
        "smtp.gmail.com",
        465,
        context=ssl.create_default_context(),
        timeout=30
    ) as server:

        server.login(
            GMAIL_ADDRESS,
            GMAIL_APP_PASSWORD
        )

        server.send_message(message)


# 3. 엑셀 확인 및 메일 발송

def run_reminder():
    today = pd.Timestamp.today().normalize()
    print(f"일정 확인 시작: {today.date()}", flush=True)
    print(f"엑셀 파일: {EXCEL_PATH}", flush=True)

    df = pd.read_excel(EXCEL_PATH)
    df["기준일"] = pd.to_datetime(df["기준일"])
    print(f"전체 일정: {len(df)}건", flush=True)
    sent_count = 0
    failed_count = 0
    skipped_count = 0

    for _, event in df.iterrows():
        event_id = str(event["event_id"])

        if str(event["활성화"]).strip().upper() != "Y":
            print(f"{event_id}: 비활성 일정 → 제외", flush=True)
            skipped_count += 1
            continue

        if pd.isna(event["기준일"]):
            print(f"{event_id}: 기준일이 비어 있음 → 제외", flush=True)
            skipped_count += 1
            continue

        days_left = (event["기준일"].normalize() - today).days
        if days_left not in REMINDER_DAYS:
            print(f"{event_id}: 남은 일수 {days_left}일 → 알림 대상 아님", flush=True)
            skipped_count += 1
            continue

        try:
            if pd.isna(event["수신자"]) or not str(event["수신자"]).strip():
                raise ValueError("수신자 주소가 비어 있습니다.")
            recipient = str(event["수신자"]).strip()
            if pd.isna(event["참고사항"]):
                event = event.copy()
                event["참고사항"] = "없음"

            print(f"{event_id}: D-{days_left} 메일 생성 중", flush=True)
            subject, body = generate_email(event, days_left)
            send_email(recipient, subject, body)
            sent_count += 1
            print(f"{event_id}: D-{days_left} 전송 완료 → {recipient}", flush=True)

        except Exception as e:
            failed_count += 1
            print(f"{event_id}: 오류 — {type(e).__name__}: {e}", flush=True)

    print(
        f"일정 확인 완료: 전송 {sent_count}건 / 오류 {failed_count}건 / 제외 {skipped_count}건",
        flush=True
    )
    if failed_count:
        raise RuntimeError(f"메일 처리 중 {failed_count}건의 오류가 발생했습니다.")


if __name__ == "__main__":
    run_reminder()
