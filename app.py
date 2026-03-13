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
    MessagingApiBlob,
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
        blob_api = MessagingApiBlob(api_client)
        content = blob_api.get_message_content(message_id)
        return content


PREFECTURES = [
    "北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県",
    "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県",
    "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県", "岐阜県",
    "静岡県", "愛知県", "三重県", "滋賀県", "京都府", "大阪府", "兵庫県",
    "奈良県", "和歌山県", "鳥取県", "島根県", "岡山県", "広島県", "山口県",
    "徳島県", "香川県", "愛媛県", "高知県", "福岡県", "佐賀県", "長崎県",
    "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県",
]


def extract_prefecture(address: str) -> str:
    """住所文字列から都道府県を抽出する"""
    for pref in PREFECTURES:
        if address.startswith(pref):
            return pref
    return ""


def postprocess_card(card_info: dict) -> dict:
    """都道府県の自動補完を行う"""
    prefecture = card_info.get("prefecture", "")
    address = card_info.get("address", "")
    if not prefecture and address:
        prefecture = extract_prefecture(address)
        card_info["prefecture"] = prefecture
    # 住所に都道府県が含まれている場合、addressから除去
    if prefecture and address.startswith(prefecture):
        card_info["address"] = address[len(prefecture):]
    return card_info


def extract_card_info(image_data: bytes) -> Optional[list]:
    """Claude Vision APIで名刺画像から情報を抽出する（複数枚対応）"""
    base64_image = base64.b64encode(image_data).decode("utf-8")

    prompt = """この画像に写っている「全ての名刺」を1枚ずつ読み取り、JSON配列で返してください。
名刺が1枚でも必ず配列で返してください。

【読み取りルール】
- 読み取れない項目は空文字にする
- 住所が複数ある場合、「本社」と記載されている住所を優先する。本社の記載がない場合は最初の住所を使用する
- 都道府県は住所から必ず単独で抽出し、prefectureに入れる（例：東京都、大阪府、北海道など）
- addressには都道府県を除いた残りの住所を入れる

【電話番号の読み取りルール】
- 印刷された電話番号だけでなく、手書きで書き加えられた電話番号・携帯番号も必ず読み取ること
- 電話番号が複数ある場合、代表電話（固定電話）と携帯電話（090/080/070始まり）を区別する
- 手書きの番号が携帯番号（090/080/070始まり）の場合はtel_mobileに入れる
- 手書きの番号が固定電話の場合はtel_mainに入れる
- 印刷された番号より手書きで追記された番号を優先する
- 数字が複数行に分かれて書かれている場合も1つの番号として結合して読み取る（例：080 / -2489 / -9661 → 080-2489-9661）

【ランク（手書きABC）の読み取りルール】
- 名刺の余白・端・空きスペースに手書きで書かれた単独のアルファベット「A」「B」「C」をランクとして読み取る
- 数字と組み合わさっていても（例：「7 A」「7B」）A/B/Cの部分だけをランクとして抽出する
- 複数文字ある場合（例：「BA」）は最初のアルファベットを優先する
- ランクが見つからない場合は空文字にする

必ず以下のJSON配列形式のみで返してください。説明文は不要です。
[
  {
    "company": "会社名",
    "title": "役職",
    "name": "顧客名",
    "tel_main": "代表電話",
    "tel_mobile": "携帯電話",
    "email": "メールアドレス",
    "prefecture": "都道府県",
    "address": "住所（都道府県を除く）",
    "rank": "手書きのABC（A, B, Cのいずれか1文字）"
  }
]"""

    try:
        response = claude_client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=4096,
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
        # JSON配列を抽出（コードブロックで囲まれている場合も対応）
        json_match = re.search(r'\[.*\]', result_text, re.DOTALL)
        if json_match:
            result = json.loads(json_match.group())
        else:
            result = json.loads(result_text)
        # 単一オブジェクトが返された場合も配列に変換
        if isinstance(result, dict):
            result = [result]
        # 都道府県の自動補完
        result = [postprocess_card(card) for card in result]
        return result
    except Exception as e:
        logger.error(f"Claude API エラー: {e}")
        return None


def append_to_sheet(display_name: str, card_list: list) -> int:
    """Googleスプレッドシートにデータを追記する（複数枚対応）。追記した件数を返す。"""
    try:
        gc = get_gspread_client()
        sheet = gc.open_by_key(GOOGLE_SHEET_ID).sheet1
        now = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        count = 0

        for card_info in card_list:
            row = [
                now,                               # 名刺取得日
                display_name,                      # 営業担当
                card_info.get("company", ""),       # 会社名
                card_info.get("title", ""),         # 役職
                card_info.get("name", ""),          # 顧客名
                card_info.get("tel_main", ""),      # 代表電話
                card_info.get("tel_mobile", ""),    # 携帯電話
                card_info.get("email", ""),         # メールアドレス
                card_info.get("prefecture", ""),    # 都道府県
                card_info.get("address", ""),       # 住所
                card_info.get("rank", ""),          # ランク
            ]
            sheet.append_row(row, value_input_option="USER_ENTERED")
            logger.info(f"スプレッドシート追記完了: {card_info.get('name', '不明')}")
            count += 1

        return count
    except Exception as e:
        logger.error(f"スプレッドシート書き込みエラー: {e}")
        return 0


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
        card_list = extract_card_info(image_data)
        if not card_list:
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
        count = append_to_sheet(display_name, card_list)

        if count > 0:
            reply_text = f"✅ 登録完了（{count}件）\n"
            warnings = []
            for i, card_info in enumerate(card_list[:count], 1):
                reply_text += (
                    f"\n【{i}件目】\n"
                    f"会社名：{card_info.get('company', '')}\n"
                    f"役職：{card_info.get('title', '')}\n"
                    f"顧客名：{card_info.get('name', '')}\n"
                    f"代表電話：{card_info.get('tel_main', '')}\n"
                    f"携帯電話：{card_info.get('tel_mobile', '')}\n"
                    f"メール：{card_info.get('email', '')}\n"
                    f"都道府県：{card_info.get('prefecture', '')}\n"
                    f"住所：{card_info.get('address', '')}\n"
                    f"ランク：{card_info.get('rank', '') or 'なし'}"
                )
                # 警告チェック
                prefix = f"【{i}件目】" if count > 1 else ""
                if not card_info.get("company"):
                    warnings.append(f"⚠️ {prefix}会社名が読み取れませんでした")
                if not card_info.get("name"):
                    warnings.append(f"⚠️ {prefix}顧客名が読み取れませんでした")
                if not card_info.get("tel_main") and not card_info.get("tel_mobile"):
                    warnings.append(f"⚠️ {prefix}電話番号が読み取れませんでした")
            if warnings:
                reply_text += "\n\n" + "\n".join(warnings)
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
