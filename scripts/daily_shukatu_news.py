#!/usr/bin/env python3
"""
daily_shukatu_news.py

大学3〜4年生の就活本番期向けに、前日公開の経済・ビジネス・社会ニュースを
Claude + web_searchツールで収集し、1本を選定して「就活ニュース」Slackチャンネルへ投稿するスクリプト。

- 環境変数:
  - ANTHROPIC_API_KEY: Anthropic APIキー
  - SLACK_WEBHOOK_URL: Slack Incoming Webhook URL

- 重複防止:
  - スクリプトと同じディレクトリの post_history.log を参照・追記する

SKILL.md（daily-shukatu-news）の書式・トーン規定に従う。
"""

from __future__ import annotations

import json
import os
import re
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

import anthropic
import requests


SCRIPT_DIR = Path(__file__).resolve().parent
HISTORY_FILE = SCRIPT_DIR / "post_history.log"

JST = timezone(timedelta(hours=9))

WEEKDAY_JP = ["月", "火", "水", "木", "金", "土", "日"]

MODEL = "claude-sonnet-4-5"
MAX_TOKENS = 4096


def load_history() -> str:
    """過去の投稿履歴を文字列として返す。未作成の場合は空文字。"""
    if not HISTORY_FILE.exists():
        return ""
    try:
        return HISTORY_FILE.read_text(encoding="utf-8")
    except OSError:
        return ""


def append_history(date_str: str, topic_keyword: str) -> None:
    """post_history.log に1行追記する。"""
    line = f"{date_str} | {topic_keyword}\n"
    with HISTORY_FILE.open("a", encoding="utf-8") as f:
        f.write(line)


def build_target_dates(today_jst: datetime) -> list[str]:
    """
    検索対象日（前日）を返す。
    月曜日なら金〜日の3日間を対象、それ以外は前日のみ。
    返却は 'YYYY年M月D日' 形式のリスト。
    """
    weekday = today_jst.weekday()  # 月=0 ... 日=6
    if weekday == 0:  # 月曜: 金土日
        days = [3, 2, 1]
    else:
        days = [1]
    return [
        (today_jst - timedelta(days=d)).strftime("%Y年%-m月%-d日")
        for d in days
    ]


def generate_post(today_jst: datetime, history_text: str) -> dict:
    """
    Claude + web_search を使って投稿テキストを生成する。
    戻り値: {"text": 投稿本文, "topic_keyword": 履歴ログ用キーワード}
    """
    client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

    target_dates = build_target_dates(today_jst)
    date_label = f"{today_jst.year}年{today_jst.month}月{today_jst.day}日({WEEKDAY_JP[today_jst.weekday()]})"

    history_block = history_text.strip() if history_text.strip() else "（履歴なし。初回実行）"

    system_prompt = """あなたはWorCry(ワークライ)の就活ニュース編集者です。
大学3〜4年生の就活本番期の学生向けに、前日公開の経済・ビジネス・社会関連ニュースを1本選び、
Slack投稿用テキストを生成します。

【厳守事項】
- 絵文字は一切使わないでください（顔文字・記号絵文字・装飾記号すべて禁止）。
- 文体はすべて「ですます調」で統一してください。だ・である調は使いません。
- Slack mrkdwn形式: 太字はアスタリスク1つで囲みます（*テキスト*）。
- 区切り線は「—————————————」を使います。
- トピックタイトルは参照元記事のタイトルをそのまま転記し、要約・改変しません。
- 焦りを煽らず背中を押す、熱量があって親しみやすく本質を伝えるトーン。
- 誹謗中傷、差別的表現は禁止。
- 「就活」という言葉はそのまま使ってかまいません。

【選定基準】（どちらか一方を満たすこと）
1. 内定取得・面接力向上に役立つ: 採用で重視される能力・姿勢の変化、面接で話せる時事知識、差がつく視点。
2. 企業選びに役立つ: 業界・企業の成長性・安定性・将来性、働き方・給与・社風の変化、注目企業戦略・市場動向。

【重複回避】
- 過去の投稿履歴に含まれるテーマ（同一カテゴリ・同一企業の同種ニュース・春闘/賃上げ等）は避けてください。

【出力形式】
以下の厳密なJSONオブジェクトのみを出力してください。説明文やコードブロック記号は不要です。
{
  "post_text": "Slackへ投稿する本文全体（改行は \\n）",
  "topic_keyword": "履歴ログ用の短いキーワード（10〜30字）"
}

post_text は以下のフォーマットに正確に従ってください:

就活ニュース｜<日付ラベル>

*<参照元記事タイトル>*
<参照元サイト名>（<URL>）

—————————————

*ひとことSummary*
<1〜2文の概要。ですます調>

*読み解きPoint*
<選定基準に沿った2〜4文。ですます調>

*おすすめAction*
<今日〜今週中にできる具体的な行動提案1〜2つ。ですます調>

—————————————

<就活生への一言。1〜2文。ですます調で背中を押す>
"""

    user_prompt = f"""本日は{date_label}です。

以下の期間に公開されたニュース（{', '.join(target_dates)}）から、web_searchツールを積極的に使って
経済・ビジネス・社会・採用・業界トレンドの最新動向を幅広く調べてください。
就活専門メディアに限定せず、日本経済新聞・東洋経済オンライン・Diamond Online・
ビジネスインサイダージャパン・NHKニュース経済・日経ビジネス・マイナビキャリアリサーチLab・
HR総研などの一般経済紙・ビジネスメディアも積極的に対象にしてください。

【過去に投稿済みのトピック（避けるべきテーマ）】
{history_block}

選定基準（内定取得・面接力向上 または 企業選び）のいずれかに強く該当するものを
1本だけ選定し、SKILL.mdのフォーマット・書式・トーン規定に完全準拠した投稿本文を
生成してください。日付ラベルは「{date_label}」を使います。

最終出力は指定のJSONオブジェクトのみにしてください。"""

    response = client.messages.create(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        system=system_prompt,
        tools=[
            {
                "type": "web_search_20250305",
                "name": "web_search",
                "max_uses": 8,
            }
        ],
        messages=[{"role": "user", "content": user_prompt}],
    )

    # 最終テキストブロック（アシスタントのテキスト回答）を抽出
    final_text = ""
    for block in response.content:
        if getattr(block, "type", None) == "text":
            final_text += block.text

    final_text = final_text.strip()
    if not final_text:
        raise RuntimeError("Claudeから空のレスポンスが返されました。")

    parsed = _extract_json_object(final_text)
    post_text = parsed.get("post_text", "").strip()
    topic_keyword = parsed.get("topic_keyword", "").strip()

    if not post_text or not topic_keyword:
        raise RuntimeError(
            "生成結果に post_text または topic_keyword が含まれていません。\n"
            f"raw:\n{final_text}"
        )

    _validate_post_text(post_text)
    return {"text": post_text, "topic_keyword": topic_keyword}


def _extract_json_object(text: str) -> dict:
    """
    テキストからJSONオブジェクトを抽出する。
    コードブロックで囲まれている場合も対応。
    """
    # ```json ... ``` ブロックを剥がす
    fenced = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, re.DOTALL)
    if fenced:
        candidate = fenced.group(1)
    else:
        # 最初の { から最後の } までを抽出
        start = text.find("{")
        end = text.rfind("}")
        if start == -1 or end == -1 or end <= start:
            raise RuntimeError(f"JSONオブジェクトを抽出できませんでした。raw:\n{text}")
        candidate = text[start : end + 1]

    try:
        return json.loads(candidate)
    except json.JSONDecodeError as e:
        raise RuntimeError(f"JSONパースに失敗しました: {e}\ncandidate:\n{candidate}")


# 絵文字判定用: 代表的なUnicodeレンジ
_EMOJI_PATTERN = re.compile(
    "["
    "\U0001F300-\U0001FAFF"  # Misc Symbols and Pictographs, Emoticons, Transport, Supplemental, etc.
    "\U00002600-\U000027BF"  # Miscellaneous Symbols, Dingbats
    "\U0001F1E6-\U0001F1FF"  # Regional Indicator Symbols
    "\U00002B00-\U00002BFF"  # Miscellaneous Symbols and Arrows
    "\U0000FE00-\U0000FE0F"  # Variation Selectors
    "\U0001F900-\U0001F9FF"  # Supplemental Symbols and Pictographs
    "]",
    flags=re.UNICODE,
)


def _validate_post_text(text: str) -> None:
    """投稿テキストの形式ルールを検証する。"""
    if _EMOJI_PATTERN.search(text):
        raise RuntimeError("投稿テキストに絵文字が含まれています（SKILL.md違反）。")

    required_markers = [
        "就活ニュース｜",
        "*ひとことSummary*",
        "*読み解きPoint*",
        "*おすすめAction*",
        "—————————————",
    ]
    for marker in required_markers:
        if marker not in text:
            raise RuntimeError(f"投稿テキストに必須要素が含まれていません: {marker}")


def post_to_slack(text: str) -> None:
    """Slack Webhookへ投稿する。"""
    webhook_url = os.environ["SLACK_WEBHOOK_URL"]
    payload = {
        "unfurl_links": False,
        "unfurl_media": False,
        "text": text,
    }
    resp = requests.post(
        webhook_url,
        data={"payload": json.dumps(payload, ensure_ascii=False)},
        timeout=30,
    )
    if resp.status_code != 200 or resp.text.strip() != "ok":
        raise RuntimeError(
            f"Slack投稿に失敗しました: status={resp.status_code}, body={resp.text!r}"
        )


def main() -> int:
    for var in ("ANTHROPIC_API_KEY", "SLACK_WEBHOOK_URL"):
        if not os.environ.get(var):
            print(f"ERROR: 環境変数 {var} が未設定です。", file=sys.stderr)
            return 2

    today_jst = datetime.now(tz=JST)
    date_str = today_jst.strftime("%Y-%m-%d")

    history_text = load_history()
    print(f"[info] 今日(JST)={date_str} / 履歴行数={len(history_text.splitlines())}")

    try:
        result = generate_post(today_jst, history_text)
    except Exception as e:
        print(f"ERROR: 投稿テキスト生成に失敗しました: {e}", file=sys.stderr)
        return 3

    post_text = result["text"]
    topic_keyword = result["topic_keyword"]

    print("---- 生成された投稿テキスト ----")
    print(post_text)
    print("---- topic_keyword:", topic_keyword, "----")

    try:
        post_to_slack(post_text)
    except Exception as e:
        print(f"ERROR: Slack投稿に失敗しました: {e}", file=sys.stderr)
        return 4

    try:
        append_history(date_str, topic_keyword)
    except Exception as e:
        # 投稿自体は成功しているので警告のみ
        print(f"WARN: 履歴ログへの追記に失敗しました: {e}", file=sys.stderr)

    print("[ok] Slack投稿と履歴追記が完了しました。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
