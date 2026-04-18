#!/usr/bin/env python3
"""
AOHARUインターン 週次KPIダッシュボード 自動更新スクリプト
毎週月曜 9:00 に実行（スケジュール設定済み）
"""
import json
import re
import urllib.request
from datetime import datetime

SHEET_ID = "1U9kPp0tU-dF0ZmIuee085yyBvKvUIWMgqx42e__DyGo"
DASHBOARD = "/Users/ynegishi/WorCry-agents/dashboard_AOHARU_週次ステップ率.html"

# ─── データ取得 ────────────────────────────────────────────
def fetch_sheet(sheet_name: str) -> dict:
    encoded = urllib.parse.quote(sheet_name)
    url = (
        f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
        f"/gviz/tq?tq=select+*&sheet={encoded}"
    )
    with urllib.request.urlopen(url, timeout=15) as resp:
        raw = resp.read().decode("utf-8")
    json_str = re.sub(r"^/\*O_o\*/\ngoogle\.visualization\.Query\.setResponse\(", "", raw)
    json_str = re.sub(r"\);$", "", json_str.strip())
    return json.loads(json_str)

def parse_rows(data: dict) -> list[list]:
    def cell_val(c):
        if c is None:
            return None
        f = c.get("f")
        v = c.get("v")
        return f if f is not None else v
    return [[cell_val(c) for c in row["c"]] for row in data["table"]["rows"]]

# ─── マーケ数値パース ──────────────────────────────────────
def parse_marke(rows):
    """
    月別データ: LINE, PD, 面談予約, 面談実施, 応募意思, セット, 決定
    月名行 → 次行がステップ率（スキップ）
    """
    month_order = ["10月","11月","12月","1月","2月","3月"]
    months = {}
    i = 0
    while i < len(rows):
        lbl = rows[i][0]
        if lbl in month_order:
            r = rows[i]
            def n(v): return int(float(v)) if v not in (None, "") else 0
            months[lbl] = dict(
                m=lbl,
                line=n(r[1]), pd=n(r[2]), yoyaku=n(r[3]),
                jissi=n(r[4]), obo=n(r[5]), set=n(r[6]), ketei=n(r[7])
            )
        i += 1
    # 対象月を12月以降に絞る
    target = [months[m] for m in ["12月","1月","2月","3月"] if m in months]
    cum = {k: sum(d[k] for d in target)
           for k in ("line","pd","yoyaku","jissi","obo","set","ketei")}
    return target, cum

# ─── KPI週次パース ─────────────────────────────────────────
def parse_weekly(rows):
    """
    週データ: 日付, 週ラベル, LINE目標, LINE実績, PD目標, PD実績,
             LINE→PD目標%, LINE→PD実績%, 面談目標, 面談予約, 面談実施,
             予→実目標%, 予→実実績%, Uセット目標, Uセット実績, 決定目標, 決定実績
    """
    month_map = {"1月": "1月", "2月": "2月", "3月": "3月", "4月": "4月"}
    week_data = []
    current_month = None
    for r in rows:
        lbl = r[1]
        if lbl in month_map:
            current_month = lbl
            continue
        if lbl in ("合計", None) or r[0] is None:
            continue
        if current_month is None:
            continue
        def n(v): return int(float(v)) if v not in (None, "") else None
        def pct(v):
            if v is None or v == "": return None
            s = str(v).replace("%","")
            try: return round(float(s))
            except: return None
        date_str = r[0].split("/") if r[0] else ["?","?"]
        week_label = f"{date_str[0]}/{date_str[1]}" if len(date_str) >= 2 else r[0]
        week_data.append(dict(
            m=current_month, w=week_label,
            lt=n(r[2]), lr=n(r[3]),
            pt=n(r[4]), pr=n(r[5]),
            lpd=pct(r[7]),
            myo=n(r[9]), mji=n(r[10]),
            yj=pct(r[12]),
            ut=n(r[13]), ur=n(r[14]),
            dt=n(r[15]), dr=n(r[16])
        ))
    return week_data

# ─── 月別週次合計 ──────────────────────────────────────────
def weekly_month_totals(week_data):
    totals = {}
    for d in week_data:
        m = d["m"]
        if m not in totals:
            totals[m] = {"lr":0,"pr":0,"myo":0,"mji":0,"ur":0,"dr":0}
        for k in ("lr","pr","myo","ur","dr"):
            if d[k] is not None: totals[m][k] += d[k]
        if d["mji"] is not None: totals[m]["mji"] += d["mji"]
    return totals

# ─── クロスステップ率（当月） ──────────────────────────────
def cross_rates(month_data_list):
    """直近月・前月の LINE→面談, 面談→セット, セット→決定"""
    by_m = {d["m"]: d for d in month_data_list}
    result = {}
    for label, d in by_m.items():
        l2f = round(d["jissi"]/d["line"]*100) if d["line"] else 0
        f2s = round(d["set"]/d["jissi"]*100)  if d["jissi"] else 0
        s2k = round(d["ketei"]/d["set"]*100)  if d["set"] else 0
        result[label] = {"l2f": l2f, "f2s": f2s, "s2k": s2k}
    return result

# ─── HTMLへの書き込み ──────────────────────────────────────
def to_js_array(week_data):
    lines = []
    for d in week_data:
        def q(v): return "null" if v is None else str(v)
        lines.append(
            f"  {{m:'{d['m']}', w:'{d['w']}', "
            f"lt:{q(d['lt'])}, lr:{q(d['lr'])}, "
            f"pt:{q(d['pt'])}, pr:{q(d['pr'])}, "
            f"myo:{q(d['myo'])}, mji:{q(d['mji'])}, "
            f"ut:{q(d['ut'])}, ur:{q(d['ur'])}, "
            f"dt:{q(d['dt'])}, dr:{q(d['dr'])}, "
            f"lpd:{q(d['lpd'])}, yj:{q(d['yj'])}}}"
        )
    return "[\n" + ",\n".join(lines) + "\n]"

def to_month_js(month_list):
    lines = []
    for d in month_list:
        lines.append(
            f"  {{m:'{d['m']}', line:{d['line']}, pd:{d['pd']}, "
            f"yoyaku:{d['yoyaku']}, jissi:{d['jissi']}, "
            f"obo:{d['obo']}, set:{d['set']}, ketei:{d['ketei']}}}"
        )
    return "[\n" + ",\n".join(lines) + "\n]"

def to_cum_js(cum):
    return (
        f"{{line:{cum['line']}, pd:{cum['pd']}, yoyaku:{cum['yoyaku']}, "
        f"jissi:{cum['jissi']}, obo:{cum['obo']}, set:{cum['set']}, ketei:{cum['ketei']}}}"
    )

def to_wmonth_js(totals):
    parts = []
    for m, t in totals.items():
        parts.append(
            f"  '{m}': {{lr:{t['lr']}, pr:{t['pr']}, "
            f"myo:{t['myo']}, mji:{t['mji']}, "
            f"ur:{t['ur']}, dr:{t['dr']}}}"
        )
    return "{\n" + ",\n".join(parts) + "\n}"

def update_html(month_data, cum, week_data, w_totals, cross):
    with open(DASHBOARD, encoding="utf-8") as f:
        html = f.read()

    now_str = datetime.now().strftime("%Y年%m月%d日 %H:%M")

    # 最終更新日
    html = re.sub(
        r'<!--LAST_UPDATED-->.*?(?=<)',
        f"<!--LAST_UPDATED-->{now_str}",
        html, flags=re.DOTALL
    )
    # 最終更新（span直接置換）
    html = html.replace(
        re.search(r'最終更新: <!--LAST_UPDATED-->.*?(?=</span>)', html, re.DOTALL).group(0),
        f"最終更新: {now_str}"
    ) if re.search(r'最終更新: <!--LAST_UPDATED-->', html) else html

    # monthData
    html = re.sub(
        r"const monthData = \[[\s\S]*?\];",
        f"const monthData = {to_month_js(month_data)};",
        html
    )
    # cumTotal
    html = re.sub(
        r"const cumTotal = \{[\s\S]*?\};",
        f"const cumTotal = {to_cum_js(cum)};",
        html
    )
    # weekData
    html = re.sub(
        r"const weekData = \[[\s\S]*?\];",
        f"const weekData = {to_js_array(week_data)};",
        html
    )
    # wMonthTotals
    html = re.sub(
        r"const wMonthTotals = \{[\s\S]*?\};",
        f"const wMonthTotals = {to_wmonth_js(w_totals)};",
        html
    )

    # クロスステップ率カード（当月3月）
    cur_m = month_data[-1]["m"]
    cr = cross.get(cur_m, {})
    prev_m = month_data[-2]["m"] if len(month_data) >= 2 else None
    pr = cross.get(prev_m, {}) if prev_m else {}

    # LINE→面談 当月値
    html = re.sub(r'(LINE → 面談実施.*?累計実績.*?)<div[^>]*font-size:36px[^>]*>(\d+)',
        lambda m: m.group(0), html)  # 静的値維持（カードは別途更新対象外）

    with open(DASHBOARD, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"✅ ダッシュボード更新完了: {now_str}")
    print(f"   マーケ数値: {len(month_data)}ヶ月分, 週次: {len(week_data)}週分")
    print(f"   当月({cur_m}): LINE→面談 {cr.get('l2f','-')}%  面談→セット {cr.get('f2s','-')}%  セット→決定 {cr.get('s2k','-')}%")

# ─── メイン ────────────────────────────────────────────────
def main():
    import urllib.parse
    print("📊 Google Sheetsからデータ取得中...")
    try:
        marke_data = fetch_sheet("マーケ数値")
        marke_rows = parse_rows(marke_data)
        month_data, cum = parse_marke(marke_rows)

        kpi_data = fetch_sheet("KPI（週次）")
        kpi_rows = parse_rows(kpi_data)
        week_data = parse_weekly(kpi_rows)
        w_totals  = weekly_month_totals(week_data)

        cross = cross_rates(month_data)
        update_html(month_data, cum, week_data, w_totals, cross)

    except Exception as e:
        print(f"❌ エラー: {e}")
        raise

if __name__ == "__main__":
    import urllib.parse
    main()
