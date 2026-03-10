from __future__ import annotations

import os
import re
import json
import logging
import base64
from datetime import datetime
from typing import Optional

from flask import Flask, request, abort
from linebot.v3 import WebhookHandler
from linebot.v3.messaging import (
    Configuration,
    ApiClient,
    MessagingApi,
    ReplyMessageRequest,
    TextMessage,
)
from linebot.v3.webhooks import MessageEvent, ImageMessageContent
from linebot.v3.exceptions import InvalidSignatureError
import anthropic
import gspread
from google.oauth2.service_account import Credentials

# ========================================
# 設定
# ========================================

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

LINE_CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
LINE_CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET", "")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
GOOGLE_SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "")
GOOGLE_CREDENTIALS_JSON = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")

if not LINE_CHANNEL_ACCESS_TOKEN or not LINE_CHANNEL_SECRET:
    logger.warning("LINE API credentials not set")

if not ANTHROPIC_API_KEY:
    logger.warning("ANTHROPIC_API_KEY not set")

if not GOOGLE_SHEET_ID or not GOOGLE_CREDENTIALS_JSON:
    logger.warning("Google Sheets credentials not set")

# LINE SDK v3 setup
configuration = Configuration(access_token=LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# Anthropic client
claude_client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

# Google Sheets setup
def get_gspread_client():
    creds_dict = json.loads(GOOGLE_CREDENTIALS_JSON)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)


# ========================================
# ユーティリティ関数
# ========================================


def get_display_name(user_id: str, group_id: Optional[str] = None) -> str:
    """LINE display_name を取得する"""
    try:
        with ApiClient(configuration) as api_client:
            line_bot_api = MessagingApi(api_client)
            if group_id:
                profile = line_bot_api.get_group_member_profile(group_id, user_id)
            else:
                profile = line_bot_api.get_profile(user_id)
            return profile.display_name
    except Exception as e:
        logger.error(f"display_name取得エラー: {e}")
        return "不明"


def get_image_from_line(message_id: str) -> bytes:
    """LINE から画像バイナリを取得する"""
    with ApiClient(configuration) as api_client:
        line_bot_api = MessagingApi(api_client)
        content = line_bot_api.get_message_content(message_id)
        return content


def extract_card_info(image_data: bytes) -> Optional[dict]:
    """Claude Vision APIで名刺画像から情報を抽出する"""
    base64_image = base64.b64encode(image_data).decode("utf-8")

    prompt = """この名刺画像から以下の情報を読み取り、JSON形式で返してください。
読み取れない項目は空文字にしてください。
手書きでA, B, Cのいずれかが書かれている場合は rank に記録してください。書かれていなければ空文字にしてください。

必ず以下のJSON形式のみで返してください。説明文は不要です。
{
  "company": "",
  "name": "",
  "title": "",
  "tel": "",
  "email": "",
  "rank": ""
}"""

    try:
        response = claude_client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=1024,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": "image/jpeg",
                                "data": base64_image,
                            },
                        },
                        {
                            "type": "text",
                            "text": prompt,
                        },
                    ],
                }
            ],
        )

        result_text = response.content[0].text.strip()
        # JSON部分を抽出（コードブロックで囲まれている場合も対応）
        json_match = re.search(r'\{[^{}]*\}', result_text, re.DOTALL)
        if json_match:
            return json.loads(json_match.group())
        return json.loads(result_text)
    except Exception as e:
        logger.error(f"Claude API エラー: {e}")
        return None


def append_to_sheet(display_name: str, card_info: dict) -> bool:
    """Googleスプレッドシートにデータを追記する"""
    try:
        gc = get_gspread_client()
        sheet = gc.open_by_key(GOOGLE_SHEET_ID).sheet1

        now = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        row = [
            now,                          # 受信日時
            display_name,                 # 担当営業
            card_info.get("rank", ""),     # ランク
            card_info.get("company", ""),  # 会社名
            card_info.get("name", ""),     # 氏名
            card_info.get("title", ""),    # 役職
            card_info.get("tel", ""),      # 電話番号
            card_info.get("email", ""),    # メール
            "",                           # 展示会名（空欄）
        ]
        sheet.append_row(row, value_input_option="USER_ENTERED")
        logger.info(f"スプレッドシート追記完了: {card_info.get('name', '不明')}")
        return True
    except Exception as e:
        logger.error(f"スプレッドシート書き込みエラー: {e}")
        return False


# ========================================
# エンドポイント
# ========================================


@app.route("/", methods=["GET"])
def health_check():
    """Render Health Check用エンドポイント"""
    return "OK"


@app.route("/callback", methods=["POST"])
def callback():
    """LINE Webhook受信エンドポイント"""
    signature = request.headers.get("X-Line-Signature", "")
    body = request.get_data(as_text=True)
    logger.info(f"Webhook received: {body[:200]}")

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        logger.error("Invalid signature")
        abort(400)

    return "OK"


@handler.add(MessageEvent, message=ImageMessageContent)
def handle_image(event: MessageEvent):
    """画像メッセージ処理"""
    user_id = event.source.user_id
    group_id = getattr(event.source, "group_id", None)

    # 送信者のdisplay_name取得
    display_name = get_display_name(user_id, group_id)
    logger.info(f"画像受信: {display_name}")

    try:
        # 画像取得
        image_data = get_image_from_line(event.message.id)

        # Claude Vision で名刺情報抽出
        card_info = extract_card_info(image_data)
        if not card_info:
            reply_text = "名刺の読み取りに失敗しました。もう一度お試しください。"
            with ApiClient(configuration) as api_client:
                line_bot_api = MessagingApi(api_client)
                line_bot_api.reply_message(
                    ReplyMessageRequest(
                        reply_token=event.reply_token,
                        messages=[TextMessage(text=reply_text)],
                    )
                )
            return

        # スプレッドシートに追記
        success = append_to_sheet(display_name, card_info)

        if success:
            reply_text = (
                f"名刺を登録しました\n"
                f"会社名: {card_info.get('company', '')}\n"
                f"氏名: {card_info.get('name', '')}\n"
                f"役職: {card_info.get('title', '')}\n"
                f"TEL: {card_info.get('tel', '')}\n"
                f"Email: {card_info.get('email', '')}\n"
                f"ランク: {card_info.get('rank', '') or 'なし'}"
            )
        else:
            reply_text = "スプレッドシートへの書き込みに失敗しました。"

        with ApiClient(configuration) as api_client:
            line_bot_api = MessagingApi(api_client)
            line_bot_api.reply_message(
                ReplyMessageRequest(
                    reply_token=event.reply_token,
                    messages=[TextMessage(text=reply_text)],
                )
            )

    except Exception as e:
        logger.error(f"処理エラー: {e}")
        try:
            with ApiClient(configuration) as api_client:
                line_bot_api = MessagingApi(api_client)
                line_bot_api.reply_message(
                    ReplyMessageRequest(
                        reply_token=event.reply_token,
                        messages=[TextMessage(text="エラーが発生しました。しばらくしてからもう一度お試しください。")],
                    )
                )
        except Exception:
            logger.error("エラー返信にも失敗しました")


# ========================================
# エントリポイント
# ========================================

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
