#!/usr/bin/env python3
"""
AOHARUインターン マーケ数値 ダッシュボード 自動集計スクリプト
スプレッドシートの「マーケ数値」タブ依存を完全排除し、ソースシートから直接集計する。

【データソース】
  シート1        → 面談実施（出欠=出席、流入月×流入ルートで集計）
  LINE登録一覧   → LINE登録数, PD（アンケートタグ）, 面談予約（申込タグ）
  学生管理       → セット, 決定（流入月+ルートあり）
  成約管理       → セット, 決定（流入月なし→名前でJOINまたは面談日月）

【新機能】
  リスティングLPルートを自動集計対象に追加
  月・ルートが増えても自動対応（スプシのメンテ不要）
"""

import csv
import io
import json
import re
import urllib.parse
import urllib.request
from collections import defaultdict
from datetime import datetime

# ──────────────────────────────────────────
# 設定
# ──────────────────────────────────────────
SHEET_ID = "1U9kPp0tU-dF0ZmIuee085yyBvKvUIWMgqx42e__DyGo"
DASHBOARD_HTML = "/Users/ynegishi/WorCry-agents/dashboard_AOHARU_マーケ数値.html"

# LINE登録一覧の流入経路 → ダッシュボードカテゴリ
ROUTE_MAP = {
    "サークル協賛":               "協賛系",
    "学生協賛":                   "協賛系",
    "ペンマーク":                 "Web広告",
    "ペンマーク（wework用）":      "Web広告",
    "ガクシー":                   "アライアンス",
    "ガクシー（wework用）":        "アライアンス",
    "HP":                         "自然流入",
    "学生向けLP_MV":              "Meta",
    "学生向けLP_フッター":         "Meta",
    "学生向けLP_選ばれる理由":     "Meta",
    "AOHARUインターン_詳細ページ": "自然流入",
    "AOHARUインターン_FV":         "自然流入",
    "紹介":                       "リアル",
    "Ref_長谷川":                 "リアル",
    "学生イベント出展_チラシ":     "リアル",
    "Instagram広告":              "Meta",
    "リスティングLP_MV":          "リスティング",   # ★新規
    "リスティングLP_フッター":     "リスティング",   # ★新規
    "リスティングLP_選ばれる理由": "リスティング",   # ★新規
}

CATEGORY_ORDER = ["協賛系", "Meta", "Web広告", "アライアンス", "リアル", "自然流入", "リスティング"]

MONTH_ORDER = ["9月","10月","11月","12月","1月","2月","3月","4月","5月","6月","7月","8月"]


# ──────────────────────────────────────────
# CSV取得
# ──────────────────────────────────────────
def fetch_csv(sheet_name: str) -> list[list[str]]:
    """Google SheetsからCSVを取得（認証不要の公開エクスポート）"""
    encoded = urllib.parse.quote(sheet_name)
    url = (
        f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
        f"/gviz/tq?tqx=out:csv&sheet={encoded}"
    )
    try:
        with urllib.request.urlopen(url, timeout=20) as resp:
            text = resp.read().decode("utf-8")
    except Exception as e:
        print(f"  ⚠️  {sheet_name} の取得失敗: {e}")
        return []
    reader = csv.reader(io.StringIO(text))
    return list(reader)


def to_int(v) -> int:
    if v is None or v == "":
        return 0
    try:
        return int(float(str(v).replace(",", "")))
    except Exception:
        return 0


def date_to_month(date_str: str) -> str | None:
    """'2025/12/08' → '12月'、'2026/03/06' → '3月'"""
    if not date_str:
        return None
    m = re.match(r"\d{4}/(\d{1,2})/", date_str)
    return (str(int(m.group(1))) + "月") if m else None


# ──────────────────────────────────────────
# 1. シート1 → 面談実施
# ──────────────────────────────────────────
def aggregate_sheet1(rows: list[list[str]]) -> dict:
    """
    シート1: 面談予約管理（booking system）
    出欠=出席 かつ 流入月あり → 面談実施1カウント

    Returns: {流入月: {流入ルート: jissi_count}}
    """
    if not rows:
        return {}
    header = rows[0]
    attend_idx     = 18  # 出欠
    route_idx      = 19  # 流入ルート
    line_month_idx = 21  # 流入月
    jissi_month_idx = 20  # 実施月（流入月がない場合のフォールバック）

    result = defaultdict(lambda: defaultdict(int))

    for row in rows[1:]:
        if len(row) <= attend_idx:
            continue
        attend = row[attend_idx].strip() if len(row) > attend_idx else ""
        route  = row[route_idx].strip()  if len(row) > route_idx  else ""
        lineM  = row[line_month_idx].strip() if len(row) > line_month_idx else ""
        jissiM = row[jissi_month_idx].strip() if len(row) > jissi_month_idx else ""

        month = lineM if lineM and lineM != "#N/A" else jissiM
        if not month or month == "#N/A":
            continue

        # 面談実施 = 出席のみ
        if attend == "出席":
            result[month][route] += 1

    return result


# ──────────────────────────────────────────
# 2. LINE登録一覧 → LINE, PD, 面談予約
# ──────────────────────────────────────────
def aggregate_line_sheet(rows: list[list[str]]) -> tuple[dict, dict]:
    """
    Returns:
      monthly: {月: {route: {line, pd, yoyaku}}}
      users:   {システム表示名: {route, month}}  ← 成約管理JOIN用
    """
    if not rows:
        return {}, {}

    monthly = defaultdict(lambda: defaultdict(lambda: {"line": 0, "pd": 0, "yoyaku": 0}))
    users: dict = {}

    pd_tag_idx     = 4   # タグ_無料個別面談申込済み → 面談予約
    survey_tag_idx = 5   # タグ_友達追加後アンケート_回答済み → PD
    month_idx      = 17  # 流入月
    route_idx      = 3   # 流入経路
    sysname_idx    = 7   # システム表示名

    for row in rows[1:]:
        if len(row) <= month_idx:
            continue
        route  = row[route_idx].strip()
        month  = row[month_idx].strip()
        pd_tag = row[pd_tag_idx].strip() if len(row) > pd_tag_idx else ""
        survey = row[survey_tag_idx].strip() if len(row) > survey_tag_idx else ""
        name   = row[sysname_idx].strip() if len(row) > sysname_idx else ""

        if not month or route == "テスト":
            continue

        monthly[month][route]["line"]   += 1
        if survey == "1": monthly[month][route]["pd"]     += 1
        if pd_tag == "1": monthly[month][route]["yoyaku"] += 1

        if name:
            users[name] = {"route": route, "month": month}

    return monthly, users


# ──────────────────────────────────────────
# 3. 学生管理 → セット, 決定
# ──────────────────────────────────────────
def aggregate_gakusei_sheet(rows: list[list[str]]) -> dict:
    """
    流入月 + ルートあり
    Returns: {流入月: {route: {set, ketei}}}
    """
    if not rows:
        return {}

    result = defaultdict(lambda: defaultdict(lambda: {"set": 0, "ketei": 0}))
    route_idx  = 13
    set_idx    = 16  # セット日
    ketei_idx  = 21  # 決定日
    month_idx  = 25  # 流入月

    for row in rows[2:]:  # 1行目:補助ヘッダー, 2行目:本ヘッダー
        if len(row) <= month_idx or not row[1].strip():
            continue
        route  = row[route_idx].strip() if len(row) > route_idx else ""
        month  = row[month_idx].strip()
        setd   = row[set_idx].strip()   if len(row) > set_idx   else ""
        ketei  = row[ketei_idx].strip() if len(row) > ketei_idx else ""

        if not month:
            continue
        if setd:   result[month][route]["set"]   += 1
        if ketei:  result[month][route]["ketei"] += 1

    return result


# ──────────────────────────────────────────
# 4. 成約管理 → セット, 決定
# ──────────────────────────────────────────
def aggregate_seiyaku_sheet(rows: list[list[str]], users: dict) -> dict:
    """
    流入月なし → システム表示名でJOIN、なければ面談日月
    Returns: {流入月: {route: {set, ketei}}}
    """
    if not rows:
        return {}

    header = rows[0]
    result = defaultdict(lambda: defaultdict(lambda: {"set": 0, "ketei": 0}))

    try:
        route_idx  = header.index("ルート")
        mendan_idx = header.index("面談日")
        set_idx    = header.index("セット日")
        ketei_idx  = header.index("決定日")
    except ValueError:
        print("  ⚠️  成約管理のヘッダー構造が想定外です")
        return result

    name_idx = 1  # 氏名

    for row in rows[1:]:
        if len(row) <= ketei_idx or not row[name_idx].strip():
            continue

        name   = row[name_idx].strip()
        route  = row[route_idx].strip() if len(row) > route_idx else ""
        mendan = row[mendan_idx].strip() if len(row) > mendan_idx else ""
        setd   = row[set_idx].strip()   if len(row) > set_idx   else ""
        ketei  = row[ketei_idx].strip() if len(row) > ketei_idx else ""

        # 流入月をLINE登録一覧からJOIN
        info = users.get(name)
        if info:
            month = info["month"]
            route = info["route"]   # LINE登録時のルートを優先
        else:
            month = date_to_month(mendan) or date_to_month(setd)

        if not month:
            continue
        if setd:  result[month][route]["set"]   += 1
        if ketei: result[month][route]["ketei"] += 1

    return result


# ──────────────────────────────────────────
# マージ & ルートカテゴリ適用
# ──────────────────────────────────────────
def build_monthly_and_route_data(
    line_data: dict,
    sheet1_data: dict,
    gakusei_data: dict,
    seiyaku_data: dict,
) -> tuple[list, dict, list]:
    """
    Returns:
      monthly_data  : [{m, line, pd, yoyaku, jissi, obo, set, ketei}, ...]
      cum_total     : {line, pd, yoyaku, jissi, obo, set, ketei}
      route_cum_list: [{cat, route, line, jissi, set, ketei}, ...]
    """

    # 全月を収集
    all_months = set()
    all_months.update(line_data.keys())
    all_months.update(sheet1_data.keys())
    all_months.update(gakusei_data.keys())
    all_months.update(seiyaku_data.keys())

    sorted_months = sorted(
        [m for m in all_months if m in MONTH_ORDER],
        key=lambda m: MONTH_ORDER.index(m)
    )

    # ルート累計
    route_cum: dict = defaultdict(lambda: defaultdict(lambda: {"line": 0, "jissi": 0, "set": 0, "ketei": 0}))

    monthly_data: list = []

    for month in sorted_months:
        # LINE指標（LINE登録一覧ベース）
        line_total = pd_total = yoyaku_total = 0
        for route, vals in line_data.get(month, {}).items():
            cat = ROUTE_MAP.get(route, "その他")
            line_total   += vals["line"]
            pd_total     += vals["pd"]
            yoyaku_total += vals["yoyaku"]
            route_cum[cat][route]["line"] += vals["line"]

        # 面談実施（シート1ベース）
        jissi_total = 0
        for route, cnt in sheet1_data.get(month, {}).items():
            cat = ROUTE_MAP.get(route, "その他")
            jissi_total += cnt
            route_cum[cat][route]["jissi"] += cnt

        # セット / 決定（学生管理のみ）
        # ※ 成約管理は学生管理のサブセットのため二重カウント回避
        set_total = ketei_total = 0
        for route, vals in gakusei_data.get(month, {}).items():
            cat = ROUTE_MAP.get(route, "その他")
            set_total   += vals["set"]
            ketei_total += vals["ketei"]
            route_cum[cat][route]["set"]   += vals["set"]
            route_cum[cat][route]["ketei"] += vals["ketei"]

        monthly_data.append({
            "m":      month,
            "line":   line_total,
            "pd":     pd_total,
            "yoyaku": yoyaku_total,
            "jissi":  jissi_total,
            "obo":    0,   # 応募意思取得は将来対応
            "set":    set_total,
            "ketei":  ketei_total,
        })

    # 累計
    cum_total = {k: sum(d[k] for d in monthly_data)
                 for k in ("line", "pd", "yoyaku", "jissi", "obo", "set", "ketei")}

    # ルートリスト（カテゴリ順・route内はline降順）
    route_cum_list: list = []
    for cat in CATEGORY_ORDER:
        cat_routes = [(r, v) for r, v in route_cum.get(cat, {}).items()]
        cat_routes.sort(key=lambda x: -x[1]["line"])
        for route, vals in cat_routes:
            route_cum_list.append({
                "cat":   cat,
                "route": route,
                "line":  vals["line"],
                "jissi": vals["jissi"],
                "set":   vals["set"],
                "ketei": vals["ketei"],
            })

    return monthly_data, cum_total, route_cum_list


# ──────────────────────────────────────────
# 今月データ抽出
# ──────────────────────────────────────────
def build_month_route_data(
    month: str,
    line_data: dict,
    sheet1_data: dict,
    gakusei_data: dict,
    seiyaku_data: dict,
) -> list:
    """指定月のルート別データを返す"""
    route_month: dict = defaultdict(lambda: defaultdict(lambda: {"line": 0, "jissi": 0, "set": 0, "ketei": 0}))

    for route, vals in line_data.get(month, {}).items():
        cat = ROUTE_MAP.get(route, "その他")
        route_month[cat][route]["line"] += vals["line"]

    for route, cnt in sheet1_data.get(month, {}).items():
        cat = ROUTE_MAP.get(route, "その他")
        route_month[cat][route]["jissi"] += cnt

    for route, vals in gakusei_data.get(month, {}).items():
        cat = ROUTE_MAP.get(route, "その他")
        route_month[cat][route]["set"]   += vals["set"]
        route_month[cat][route]["ketei"] += vals["ketei"]

    result = []
    for cat in CATEGORY_ORDER:
        cat_routes = [(r, v) for r, v in route_month.get(cat, {}).items()]
        cat_routes.sort(key=lambda x: -x[1]["line"])
        for route, vals in cat_routes:
            result.append({
                "cat": cat, "route": route,
                "line": vals["line"], "jissi": vals["jissi"],
                "set": vals["set"], "ketei": vals["ketei"],
            })
    return result


# ──────────────────────────────────────────
# HTML更新
# ──────────────────────────────────────────
def fmt_month_data(monthly: list) -> str:
    lines = []
    for d in monthly:
        lines.append(
            f"  {{m:'{d['m']}', line:{d['line']}, pd:{d['pd']}, "
            f"yoyaku:{d['yoyaku']}, jissi:{d['jissi']}, "
            f"obo:{d['obo']}, set:{d['set']}, ketei:{d['ketei']}}}"
        )
    return "[\n" + ",\n".join(lines) + "\n]"


def fmt_cum(cum: dict) -> str:
    return (
        f"{{line:{cum['line']}, pd:{cum['pd']}, yoyaku:{cum['yoyaku']}, "
        f"jissi:{cum['jissi']}, obo:{cum['obo']}, set:{cum['set']}, ketei:{cum['ketei']}}}"
    )


def fmt_route_list(routes: list, key: str) -> str:
    lines = []
    for r in routes:
        lines.append(
            f"    {{cat:'{r['cat']}', route:'{r['route']}', "
            f"line:{r['line']}, jissi:{r['jissi']}, set:{r['set']}, ketei:{r['ketei']}}}"
        )
    return f"  {key}: [\n" + ",\n".join(lines) + "\n  ]"


def update_html(monthly_data, cum_total, route_cum, route_month, route_quarter):
    with open(DASHBOARD_HTML, encoding="utf-8") as f:
        html = f.read()

    now_str = datetime.now().strftime("%Y年%m月%d日 %H:%M")

    # 最終更新日
    html = re.sub(
        r"最終更新:.*?(?=</span>)",
        f"最終更新: {now_str}",
        html
    )

    # monthData
    html = re.sub(
        r"const monthData = \[[\s\S]*?\];",
        f"const monthData = {fmt_month_data(monthly_data)};",
        html
    )

    # cumTotal
    html = re.sub(
        r"const cumTotal = \{[\s\S]*?\};",
        f"const cumTotal = {fmt_cum(cum_total)};",
        html
    )

    # routeData（cum / month / quarter）
    cum_js    = fmt_route_list(route_cum,     "cum")
    month_js  = fmt_route_list(route_month,   "month")
    quarter_js = fmt_route_list(route_quarter, "quarter")

    html = re.sub(
        r"const routeData = \{[\s\S]*?\};",
        f"const routeData = {{\n{cum_js},\n{month_js},\n{quarter_js}\n}};",
        html
    )

    # quarter = month と同値を示すコメント行を削除
    html = re.sub(r"routeData\.quarter = routeData\.month.*\n", "", html)

    with open(DASHBOARD_HTML, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"✅ ダッシュボード更新完了: {now_str}")


# ──────────────────────────────────────────
# メイン
# ──────────────────────────────────────────
def main():
    print("📊 Google Sheetsからデータ取得中...")

    print("  ① シート1（面談実施）を取得...")
    sheet1_rows = fetch_csv("シート1")
    sheet1_data = aggregate_sheet1(sheet1_rows)

    print("  ② LINE登録一覧（LINE/PD/面談予約）を取得...")
    line_rows, users = fetch_csv("LINE登録一覧"), {}
    line_data, users = aggregate_line_sheet(line_rows)

    print("  ③ 学生管理（セット/決定）を取得...")
    gakusei_rows = fetch_csv("学生管理")
    gakusei_data = aggregate_gakusei_sheet(gakusei_rows)

    print("  ④ 成約管理（セット/決定）を取得...")
    seiyaku_rows = fetch_csv("成約管理")
    seiyaku_data = aggregate_seiyaku_sheet(seiyaku_rows, users)

    # 集計
    print("📐 集計中...")
    monthly_data, cum_total, route_cum = build_monthly_and_route_data(
        line_data, sheet1_data, gakusei_data, seiyaku_data
    )

    # 今月（最終月）
    cur_month = monthly_data[-1]["m"] if monthly_data else "3月"
    route_month = build_month_route_data(cur_month, line_data, sheet1_data, gakusei_data, seiyaku_data)
    route_quarter = route_month  # 四半期 ≒ 今月（期中）

    # 結果表示
    print(f"\n📈 集計結果 ({len(monthly_data)}ヶ月分):")
    for d in monthly_data:
        lpd = round(d["pd"]     / d["line"]   * 100) if d["line"]   else 0
        pdy = round(d["yoyaku"] / d["pd"]     * 100) if d["pd"]     else 0
        yj  = round(d["jissi"]  / d["yoyaku"] * 100) if d["yoyaku"] else 0
        print(f"  {d['m']:4s}  LINE:{d['line']:4d}  PD:{d['pd']:4d}({lpd}%)  "
              f"予約:{d['yoyaku']:3d}({pdy}%)  実施:{d['jissi']:3d}({yj}%)  "
              f"セット:{d['set']:3d}  決定:{d['ketei']:3d}")

    print(f"\n🔢 累計: LINE={cum_total['line']}  PD={cum_total['pd']}  "
          f"予約={cum_total['yoyaku']}  実施={cum_total['jissi']}  "
          f"セット={cum_total['set']}  決定={cum_total['ketei']}")

    print(f"\n✏️  ダッシュボードHTML更新中...")
    update_html(monthly_data, cum_total, route_cum, route_month, route_quarter)


if __name__ == "__main__":
    main()
