"""
GGPL新宿 ランキングダッシュボード 生成スクリプト
GitHub Actionsで実行し、index.htmlを生成してGitHub Pagesで公開する。
"""

import os
import sys
import json
import re
import base64
from datetime import datetime, timezone, timedelta, date

import jpholiday
import gspread
from google.oauth2.service_account import Credentials

JST = timezone(timedelta(hours=9))

RING_SHEET_ID = os.environ.get("RING_SHEET_ID", "")
TONAME_SHEET_ID = os.environ.get("TONAME_SHEET_ID", "")
SA_JSON_STR = os.environ.get("SERVICE_ACCOUNT_JSON", "")

SEASON_TAB_PATTERN = re.compile(r"^\d{1,2}/\d{1,2}-\d{1,2}/\d{1,2}$")
MONTH_TAB_PATTERN = re.compile(r"^\d{4}年\d{1,2}月$|^\d{1,2}月$|^\d{4}/\d{1,2}$")

def get_gc():
    sa_json = SA_JSON_STR.strip()
    try:
        sa_info = json.loads(sa_json)
    except (json.JSONDecodeError, ValueError):
        try:
            fixed = re.sub(
                r'("private_key"\s*:\s*")(.*?)(")',
                lambda m: m.group(1) + m.group(2).replace('\n', '\\n').replace('\r', '') + m.group(3),
                sa_json,
                flags=re.DOTALL
            )
            sa_info = json.loads(fixed)
        except (json.JSONDecodeError, ValueError):
            sa_info = json.loads(base64.b64decode(sa_json.encode()).decode('utf-8'))
    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
    return gspread.authorize(creds)


def _detect_rank(cell):
    """Returns rank number 1-30 or None."""
    r = cell.strip().replace(" ", "")
    for i in range(1, 31):
        suffix = "st" if i == 1 else "nd" if i == 2 else "rd" if i == 3 else "th"
        if r in (f"{i}{suffix}", f"{i}位", str(i)):
            return i
    return None


def _rank_label(n):
    suffix = "st" if n == 1 else "nd" if n == 2 else "rd" if n == 3 else "th"
    return f"{n}{suffix}"


# ── Ring season tabs ──
# Each tab has raw game data AND a ranking summary section.
# The summary section is identified by a header row with "名前" and "獲得ポイント".
# We find ALL such headers and return the section with the most entries.

def get_ring_ranking_from_sheet(ws):
    rows = ws.get_all_values()
    print(f"    {ws.title}: {len(rows)}行")

    # Find all "名前 | 獲得ポイント" header rows
    headers = []
    for i, row in enumerate(rows):
        cells = [str(c).strip() for c in row]
        if "名前" in cells and "獲得ポイント" in cells:
            name_col = cells.index("名前")
            pts_col  = cells.index("獲得ポイント")
            rank_col = max(0, name_col - 1)
            headers.append((i, rank_col, name_col, pts_col))

    if not headers:
        print(f"    ランキングテーブルが見つかりません")
        return []

    best = []
    for header_i, rank_col, name_col, pts_col in headers:
        ranking = []
        auto_rank = 0
        for row in rows[header_i + 1:]:
            cells = [str(c).strip() for c in row]
            # Stop at next "名前 | 獲得ポイント" header
            if "名前" in cells and "獲得ポイント" in cells:
                break
            if len(row) <= name_col:
                continue
            a_cell   = cells[rank_col] if rank_col < len(cells) else ""
            name_raw = cells[name_col]
            pts_raw  = cells[pts_col] if pts_col < len(cells) else ""

            if not name_raw or name_raw in ("名前", "#N/A") or ":-:" in name_raw:
                continue
            name = re.sub(r"様$", "", name_raw).strip()
            if not name:
                continue
            pts_num = re.sub(r"[^\d]", "", pts_raw)
            if not pts_num or int(pts_num) == 0:
                continue

            rank_num = _detect_rank(a_cell)
            if rank_num is not None:
                auto_rank = rank_num
            else:
                auto_rank += 1
                rank_num = auto_rank

            ranking.append({
                "rank": _rank_label(rank_num),
                "name": name,
                "points": f"{int(pts_num):,}pt",
            })

        if len(ranking) > len(best):
            best = ranking

    print(f"    → {len(best)}件")
    return best


def fetch_ring_seasons(gc, sheet_id):
    try:
        sh = gc.open_by_key(sheet_id)
    except Exception as e:
        print(f"  リングシート読み込みエラー: {e}")
        return []

    seasons = []
    for ws in sh.worksheets():
        if SEASON_TAB_PATTERN.match(ws.title):
            print(f"  タブ読み込み中: {ws.title}")
            ranking = get_ring_ranking_from_sheet(ws)
            seasons.append({"title": ws.title, "ranking": ranking})
    return seasons


def fetch_non_game_dates(gc, sheet_id) -> set:
    """デイリー運用タブのK列から非開催日を取得し、金・土・祝前日も追加する。"""
    non_game = set()
    try:
        sh = gc.open_by_key(sheet_id)
        ws = sh.worksheet("デイリー運用")
        col_k = ws.col_values(11)  # K列
        for cell in col_k:
            cell = cell.strip()
            if re.match(r"\d{4}/\d{2}/\d{2}", cell):
                non_game.add(cell)
        print(f"  非開催日リスト: {len(non_game)}件読み込み")
    except Exception as e:
        print(f"  非開催日リスト読み込みスキップ: {e}")
    return non_game


def is_non_game_day(date_str: str, non_game_dates: set) -> bool:
    """金・土・祝前日・スプシ記載日は非開催。"""
    if date_str in non_game_dates:
        return True
    try:
        d = date(int(date_str[:4]), int(date_str[5:7]), int(date_str[8:10]))
    except ValueError:
        return False
    # 金曜(4)・土曜(5)
    if d.weekday() in (4, 5):
        return True
    # 祝前日（翌日が祝日）
    next_day = d + timedelta(days=1)
    if jpholiday.is_holiday(next_day):
        return True
    return False


# ── Daily ring ranking: aggregate col C=name, col I=net points, col J=date
#    from all season tabs (デイリー運用 is just an ops manual, not data)

def fetch_daily_ranking(gc, sheet_id):
    try:
        sh = gc.open_by_key(sheet_id)
    except Exception as e:
        print(f"  デイリー読み込みエラー: {e}")
        return []

    non_game_dates = fetch_non_game_dates(gc, sheet_id)

    by_date = {}
    for ws in sh.worksheets():
        if not SEASON_TAB_PATTERN.match(ws.title):
            continue
        print(f"  デイリー集計: {ws.title}")
        rows = ws.get_all_values()
        for row in rows:
            if len(row) < 10:
                continue
            name = str(row[2]).strip()
            raw = str(row[8]).strip()
            date_str = str(row[9]).strip()

            if not name or not date_str or not re.match(r"\d{4}/\d{2}/\d{2}", date_str):
                continue
            # デイリーランキングは2026年6月以降のみ表示
            if date_str < "2026/06/01":
                continue
            if is_non_game_day(date_str, non_game_dates):
                continue
            minus = raw.startswith("-") or raw.startswith("−")
            points_str = re.sub(r"[^\d.]", "", raw)
            if not points_str:
                continue
            try:
                points = int(float(points_str))
            except ValueError:
                continue
            if minus:
                points = -points
            if points <= 0:
                continue

            by_date.setdefault(date_str, {})
            by_date[date_str][name] = by_date[date_str].get(name, 0) + points

    seasons = []
    for date_str in sorted(by_date.keys(), reverse=True):
        players = sorted(by_date[date_str].items(), key=lambda x: x[1], reverse=True)
        ranking = [
            {"rank": _rank_label(i + 1), "name": n, "points": f"{p:,}pt"}
            for i, (n, p) in enumerate(players)
        ]
        m, d = int(date_str[5:7]), int(date_str[8:10])
        seasons.append({"title": f"{m}/{d}", "ranking": ranking})

    print(f"  デイリー: {len(seasons)}日分")
    return seasons


# ── Tournament monthly tabs: A=1位/2位, C=name, L(idx11)=獲得skill → sum per player ──

def fetch_toname_seasons(gc, sheet_id):
    try:
        sh = gc.open_by_key(sheet_id)
    except Exception as e:
        print(f"  トナメシート読み込みエラー: {e}")
        return []

    seasons = []
    for ws in sh.worksheets():
        title = ws.title
        if not MONTH_TAB_PATTERN.match(title):
            continue

        print(f"  タブ読み込み中: {title}")
        rows = ws.get_all_values()
        player_pts = {}

        for row in rows:
            if len(row) < 12:
                continue
            a_cell = str(row[0]).strip()
            name   = str(row[2]).strip()
            skill_raw = str(row[11]).strip().replace(",", "").replace(" ", "")

            if not _detect_rank(a_cell):
                continue
            if not name or name in ("名前", "#N/A", ""):
                continue
            try:
                skill = int(float(skill_raw))
            except (ValueError, TypeError):
                continue
            if skill <= 0:
                continue

            player_pts[name] = player_pts.get(name, 0) + skill

        ranking = [
            {"rank": _rank_label(i + 1), "name": n, "points": f"{p:,}pt"}
            for i, (n, p) in enumerate(
                sorted(player_pts.items(), key=lambda x: x[1], reverse=True)
            )
        ]
        seasons.append({"title": title, "ranking": ranking})

    return seasons


# ── HTML generation ───────────────────────────────────────────────────────────

def _podium_html(ranking, color_key):
    if not ranking:
        return '<p class="no-data">データなし</p>'

    top3 = (ranking + [None, None, None])[:3]
    first, second, third = top3[0], top3[1], top3[2]

    def esc(s):
        return s.replace("&", "&amp;").replace('"', "&quot;")

    def place_html(player, place, block_cls, block_num):
        name_attr = f' data-name="{esc(player["name"])}"' if player else ""
        h = f'<div class="podium-place {place}"{name_attr}>'
        if player:
            if place == "first":
                h += '<div class="podium-stars">✦ ✦ ✦</div>'
            h += f'<div class="pod-rank-badge badge-{place}">{block_num}</div>'
            h += f'<div class="pod-name">{player["name"]}</div>'
            h += f'<div class="pod-pts {color_key}-pts">{player["points"]}</div>'
        h += f'<div class="pod-block {block_cls}"></div>'
        h += '</div>'
        return h

    html = '<div class="podium-wrap">'
    html += place_html(second, "second", "block-silver", "2")
    html += place_html(first,  "first",  "block-gold",   "1")
    html += place_html(third,  "third",  "block-bronze",  "3")
    html += '</div>'

    rest = ranking[3:]
    if rest:
        html += '<table class="rest-table"><tbody>'
        for r in rest:
            html += (f'<tr data-name="{esc(r["name"])}">'
                     f'<td class="rest-rank">{r["rank"]}</td>'
                     f'<td class="rest-name">{r["name"]}</td>'
                     f'<td class="rest-pts">{r["points"]}</td>'
                     f'</tr>')
        html += '</tbody></table>'

    return html


def _season_tabs_html(seasons, color_key, no_data_msg):
    if not seasons:
        return f'<p class="no-data">{no_data_msg}</p>'

    tabs = '<div class="season-tabs">'
    for i, s in enumerate(seasons):
        active = "active" if i == 0 else ""
        past   = 'data-past="true"' if i > 0 else ""
        tabs += (f'<button class="season-btn {active} {color_key}-tab-btn" '
                 f'{past} onclick="switchSeason(this,\'{color_key}-sp-{i}\')">{s["title"]}</button>')
    tabs += '</div>'

    panels = ''
    for i, s in enumerate(seasons):
        active = "active" if i == 0 else ""
        panels += f'<div class="season-panel {active}" id="{color_key}-sp-{i}">'
        panels += _podium_html(s["ranking"], color_key)
        panels += '</div>'

    return tabs + panels


def _build_search_js(ring_seasons, daily_seasons, toname_seasons):
    entries = []
    cats = [
        ("ring",   "リング",   ring_seasons),
        ("daily",  "デイリー", daily_seasons),
        ("toname", "トナメ",   toname_seasons),
    ]
    for cat_key, cat_label, seasons in cats:
        for i, s in enumerate(seasons):
            for r in s["ranking"]:
                entries.append({
                    "cat": cat_key, "catLabel": cat_label,
                    "catPanel": f"{cat_key}-panel",
                    "panelId": f"{cat_key}-sp-{i}",
                    "tabIdx": i,
                    "season": s["title"],
                    "rank": r["rank"], "name": r["name"], "points": r["points"],
                })
    return json.dumps(entries, ensure_ascii=False)


def generate_html(ring_seasons, daily_seasons, toname_seasons, updated_at):
    ring_html     = _season_tabs_html(ring_seasons,   "ring",   "シーズンデータなし")
    daily_html    = _season_tabs_html(daily_seasons,  "daily",  "デイリーデータなし")
    toname_html   = _season_tabs_html(toname_seasons, "toname", "シーズンデータなし")
    search_data_js = _build_search_js(ring_seasons, daily_seasons, toname_seasons)

    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>GGP LIVE SHINJUKU - RANKING</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Cinzel:wght@600;900&family=Noto+Sans+JP:wght@400;700;900&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/canvas-confetti@1.9.3/dist/confetti.browser.min.js"></script>
<style>
/* ── Variables ── */
:root {{
  --gold:      #b8922a;
  --gold-lt:   #d4af37;
  --ring-c:    #c0392b;
  --ring-lt:   #e74c3c;
  --daily-c:   #1e8449;
  --daily-lt:  #27ae60;
  --toname-c:  #1a5276;
  --toname-lt: #2980b9;
  --bg:        #f7f4f0;
  --card:      #ffffff;
  --border:    #e8e0d4;
  --text:      #1a1a1a;
  --muted:     #888;
  --shadow:    0 2px 16px rgba(0,0,0,0.08);
}}
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{
  font-family: 'Noto Sans JP', sans-serif;
  background: var(--bg);
  color: var(--text);
  min-height: 100vh;
}}

/* ── Header ── */
header {{
  background: #0d0d1a;
  border-bottom: 3px solid var(--gold-lt);
  padding: 0 20px;
  display: flex;
  align-items: center;
  gap: 20px;
  min-height: 80px;
  position: relative;
}}
.header-logo {{
  display: flex;
  align-items: center;
  gap: 10px;
  flex-shrink: 0;
}}
.logo-img {{
  height: 60px;
  width: auto;
  object-fit: contain;
  filter:
    drop-shadow(0 0 1px rgba(255,255,255,0.9))
    drop-shadow(0 0 3px rgba(255,255,255,0.5));
}}
.logo-fallback {{
  font-family: 'Cinzel', serif;
  font-size: 1.1rem;
  font-weight: 900;
  color: var(--gold-lt);
  letter-spacing: 0.1em;
  line-height: 1.2;
  white-space: nowrap;
}}
.logo-fallback small {{
  display: block;
  font-size: 0.5rem;
  letter-spacing: 0.4em;
  color: rgba(212,175,55,0.6);
  font-weight: 600;
}}
.header-center {{
  flex: 1;
  text-align: center;
}}
.header-title {{
  font-family: 'Cinzel', serif;
  font-size: clamp(0.9rem, 2.5vw, 1.4rem);
  font-weight: 900;
  color: #fff;
  letter-spacing: 0.25em;
}}
.header-sub {{
  font-size: 0.6rem;
  letter-spacing: 0.4em;
  color: rgba(212,175,55,0.6);
  margin-top: 3px;
}}
.header-updated {{
  font-size: 0.65rem;
  color: rgba(255,255,255,0.35);
  text-align: right;
  white-space: nowrap;
  flex-shrink: 0;
}}

/* ── Category nav ── */
.cat-nav {{
  display: flex;
  justify-content: center;
  gap: 0;
  background: #fff;
  border-bottom: 1px solid var(--border);
  overflow-x: auto;
}}
.cat-btn {{
  flex: 1;
  max-width: 200px;
  padding: 14px 20px;
  border: none;
  border-bottom: 3px solid transparent;
  font-family: 'Noto Sans JP', sans-serif;
  font-size: 0.82rem;
  font-weight: 700;
  cursor: pointer;
  background: transparent;
  color: var(--muted);
  transition: all 0.2s;
  white-space: nowrap;
  letter-spacing: 0.04em;
}}
.cat-btn:hover {{ background: #fafafa; }}
.cat-btn.ring-cat.active   {{ color: var(--ring-c);   border-color: var(--ring-c); }}
.cat-btn.daily-cat.active  {{ color: var(--daily-c);  border-color: var(--daily-c); }}
.cat-btn.toname-cat.active {{ color: var(--toname-c); border-color: var(--toname-c); }}

/* ── Category panels ── */
.category-panel {{ display: none; }}
.category-panel.active {{ display: block; }}

/* ── Hero banner ── */
.hero {{
  padding: 6px 20px 4px;
  text-align: center;
  position: relative;
  overflow: hidden;
}}
.ring-hero   {{ background: linear-gradient(135deg, #2c0000 0%, #7b1010 50%, #2c0000 100%); }}
.daily-hero  {{ background: linear-gradient(135deg, #001a0a 0%, #145a32 50%, #001a0a 100%); }}
.toname-hero {{ background: linear-gradient(135deg, #000d1a 0%, #1a3a5c 50%, #000d1a 100%); }}
.hero::before {{
  content: '♠ ♥ ♦ ♣';
  position: absolute; inset: 0;
  display: flex; align-items: center; justify-content: center;
  font-size: 6rem; letter-spacing: 1em;
  color: rgba(255,255,255,0.04);
  pointer-events: none;
}}
.hero-label {{
  font-family: 'Cinzel', serif;
  font-size: 0.68rem;
  letter-spacing: 0.55em;
  color: rgba(212,175,55,0.75);
  margin-bottom: 8px;
  position: relative;
  white-space: nowrap;
}}
.hero-title-img {{
  display: block;
  width: min(94%, 820px);
  height: auto;
  margin: 0 auto;
  position: relative;
  filter:
    drop-shadow(0 0 1px rgba(212,175,55,1.0))
    drop-shadow(0 0 3px rgba(212,175,55,0.85))
    drop-shadow(0 3px 8px rgba(0,0,0,0.6));
}}

/* ── Content area ── */
.content-wrap {{
  max-width: 720px;
  margin: 0 auto;
  padding: 24px 16px 48px;
}}

/* ── Season tabs ── */
.season-tabs {{
  display: flex;
  flex-wrap: nowrap;
  overflow-x: auto;
  gap: 3px;
  margin-bottom: 24px;
  padding: 4px;
  background: rgba(0,0,0,0.06);
  border-radius: 12px;
  -webkit-overflow-scrolling: touch;
  scrollbar-width: none;
}}
.season-tabs::-webkit-scrollbar {{ display: none; }}
.season-btn {{
  flex-shrink: 0;
  padding: 8px 18px;
  border-radius: 9px;
  border: none;
  font-size: 0.78rem;
  font-weight: 700;
  cursor: pointer;
  background: transparent;
  color: #888;
  transition: all 0.2s;
  white-space: nowrap;
  letter-spacing: 0.02em;
}}
.season-btn.active {{
  background: #fff;
  box-shadow: 0 1px 8px rgba(0,0,0,0.12);
}}
.ring-tab-btn.active   {{ color: var(--ring-c); }}
.daily-tab-btn.active  {{ color: var(--daily-c); }}
.toname-tab-btn.active {{ color: var(--toname-c); }}

.season-panel {{ display: none; }}
.season-panel.active {{ display: block; }}

/* ── Podium ── */
.podium-wrap {{
  display: flex;
  align-items: flex-end;
  justify-content: center;
  gap: 8px;
  margin: 4px 0 20px;
}}
.podium-place {{
  display: flex;
  flex-direction: column;
  align-items: center;
  flex: 1;
  max-width: 200px;
}}
.podium-place.first  {{ order: 2; background: radial-gradient(ellipse at 50% 55%, rgba(212,175,55,0.13) 0%, transparent 70%); border-radius: 12px; padding-top: 6px; }}
.podium-place.second {{ order: 1; }}
.podium-place.third  {{ order: 3; }}

.podium-stars {{
  font-size: 0.6rem;
  color: #d4af37;
  letter-spacing: 0.5em;
  margin-bottom: 5px;
  opacity: 0.9;
}}
.pod-rank-badge {{
  border-radius: 50%;
  display: flex; align-items: center; justify-content: center;
  font-family: 'Cinzel', serif;
  font-weight: 900;
  color: #fff;
  margin-bottom: 7px;
  letter-spacing: 0;
}}
.badge-first {{
  width: 48px; height: 48px; font-size: 1.3rem;
  background: linear-gradient(145deg, #ffe066 0%, #d4a017 40%, #8b5e00 100%);
  box-shadow: 0 0 0 3px rgba(212,175,55,0.35), 0 0 16px rgba(212,175,55,0.7), 0 3px 10px rgba(0,0,0,0.3);
  animation: gold-pulse 2.5s ease-in-out infinite;
}}
.badge-second {{
  width: 40px; height: 40px; font-size: 1.05rem;
  background: linear-gradient(145deg, #e8e8e8 0%, #a0a0a0 50%, #606060 100%);
  box-shadow: 0 0 0 2px rgba(180,180,180,0.3), 0 2px 10px rgba(0,0,0,0.25);
}}
.badge-third {{
  width: 38px; height: 38px; font-size: 1rem;
  background: linear-gradient(145deg, #f0a858 0%, #c07020 50%, #7a3c00 100%);
  box-shadow: 0 0 0 2px rgba(200,120,50,0.3), 0 2px 10px rgba(0,0,0,0.25);
}}
@keyframes gold-pulse {{
  0%, 100% {{ box-shadow: 0 0 0 3px rgba(212,175,55,0.35), 0 0 16px rgba(212,175,55,0.7), 0 3px 10px rgba(0,0,0,0.3); }}
  50%       {{ box-shadow: 0 0 0 5px rgba(212,175,55,0.2),  0 0 32px rgba(212,175,55,1.0), 0 3px 10px rgba(0,0,0,0.3); }}
}}

.pod-name {{
  font-size: 1.0rem;
  font-weight: 900;
  text-align: center;
  color: var(--text);
  margin: 6px 4px 3px;
  word-break: break-all;
  line-height: 1.2;
}}
.podium-place.first .pod-name {{
  font-size: 1.15rem;
}}
.pod-pts {{
  font-size: 0.88rem;
  font-weight: 700;
  margin-bottom: 10px;
}}
.podium-place.first .pod-pts {{
  font-size: 1.0rem;
}}
.ring-pts   {{ color: var(--ring-c); }}
.daily-pts  {{ color: var(--daily-c); }}
.toname-pts {{ color: var(--toname-c); }}

.pod-block {{
  width: 100%;
  border-radius: 6px 6px 0 0;
  position: relative;
  overflow: hidden;
}}
.pod-block::after {{
  content: '';
  position: absolute;
  top: 0; left: -120%;
  width: 55%; height: 100%;
  background: linear-gradient(105deg, transparent 20%, rgba(255,255,255,0.28) 50%, transparent 80%);
  animation: block-shine 4s ease-in-out infinite;
}}
.block-gold   {{ height: 90px; background: linear-gradient(180deg, #f0d060 0%, #c9950a 40%, #7a5000 100%); }}
.block-silver {{ height: 64px; background: linear-gradient(180deg, #d8d8d8 0%, #909090 50%, #505050 100%); }}
.block-bronze {{ height: 46px; background: linear-gradient(180deg, #e8a060 0%, #b06020 50%, #6a3000 100%); }}
@keyframes block-shine {{
  0%    {{ left: -120%; }}
  40%, 100% {{ left: 160%; }}
}}

/* ── Rest table (4th+) ── */
.rest-table {{
  width: 100%;
  border-collapse: collapse;
  margin-top: 4px;
  background: var(--card);
  border-radius: 8px;
  overflow: hidden;
  box-shadow: var(--shadow);
}}
.rest-table tr {{
  border-bottom: 1px solid #f0ebe4;
  transition: background 0.15s;
}}
.rest-table tr:last-child {{ border-bottom: none; }}
.rest-table tr:hover {{ background: #faf7f4; }}
.rest-rank {{
  padding: 11px 12px;
  width: 52px;
  color: var(--muted);
  font-size: 0.78rem;
  font-weight: 700;
}}
.rest-name {{ padding: 11px 8px; font-weight: 700; }}
.rest-pts  {{
  padding: 11px 14px;
  text-align: right;
  font-weight: 700;
  color: var(--gold);
  white-space: nowrap;
  font-variant-numeric: tabular-nums;
}}

.no-data {{ color: var(--muted); font-size: 0.88rem; padding: 24px 0; text-align: center; }}

/* ── Search ── */
.search-btn {{
  background: none; border: none; cursor: pointer;
  color: rgba(255,255,255,0.65); padding: 8px;
  display: flex; align-items: center; flex-shrink: 0;
  transition: color 0.2s;
}}
.search-btn:hover {{ color: var(--gold-lt); }}
.search-btn svg {{ width: 24px; height: 24px; }}
.search-overlay {{
  display: none; position: fixed; inset: 0;
  background: rgba(0,0,0,0.55); z-index: 1000;
  align-items: flex-start; justify-content: center;
  padding: 72px 16px 24px;
}}
.search-overlay.open {{ display: flex; }}
.search-modal {{
  background: #fff; border-radius: 16px;
  width: min(100%, 600px); max-height: 72vh;
  display: flex; flex-direction: column; overflow: hidden;
  box-shadow: 0 20px 60px rgba(0,0,0,0.4);
}}
.search-input-wrap {{
  display: flex; align-items: center;
  padding: 12px 16px; gap: 10px;
  border-bottom: 1px solid var(--border);
}}
.search-input {{
  flex: 1; border: none; outline: none;
  font-size: 1rem; font-family: 'Noto Sans JP', sans-serif;
  color: var(--text); background: transparent;
}}
.search-input::placeholder {{ color: #bbb; }}
.search-close {{
  background: none; border: none; cursor: pointer;
  color: var(--muted); font-size: 1.1rem; padding: 4px 6px;
  line-height: 1;
}}
.search-results {{ overflow-y: auto; }}
.search-empty {{ text-align: center; color: var(--muted); padding: 32px; font-size: 0.88rem; }}
.search-hint  {{ text-align: center; color: #ccc; padding: 28px; font-size: 0.82rem; }}
.search-result-item {{
  display: flex; align-items: center; gap: 10px;
  padding: 10px 16px; border-bottom: 1px solid #f5f0eb;
  transition: background 0.12s;
}}
.search-result-item:last-child {{ border-bottom: none; }}
.search-result-item:hover {{ background: #faf7f4; }}
.s-cat {{
  flex-shrink: 0; font-size: 0.62rem; font-weight: 700;
  padding: 3px 7px; border-radius: 20px; color: #fff; white-space: nowrap;
}}
.s-cat.ring   {{ background: var(--ring-c); }}
.s-cat.daily  {{ background: var(--daily-c); }}
.s-cat.toname {{ background: var(--toname-c); }}
.s-season {{ font-size: 0.7rem; color: var(--muted); white-space: nowrap; min-width: 48px; flex-shrink: 0; }}
.s-rank   {{ font-size: 0.72rem; font-weight: 700; color: var(--muted); min-width: 28px; flex-shrink: 0; }}
.s-name   {{ flex: 1; font-weight: 700; font-size: 0.92rem; }}
.s-pts    {{ font-weight: 700; color: var(--gold); font-size: 0.82rem; white-space: nowrap; }}
.s-arrow  {{ color: #ccc; font-size: 1.1rem; flex-shrink: 0; margin-left: 4px; }}
.search-result-item {{ cursor: pointer; }}

/* ── Player flash highlight ── */
@keyframes flash-player {{
  0%   {{ box-shadow: none; background: transparent; }}
  12%  {{ box-shadow: 0 0 0 3px rgba(212,175,55,0.9), 0 0 24px rgba(212,175,55,0.5); background: rgba(212,175,55,0.18); }}
  24%  {{ box-shadow: none; background: transparent; }}
  36%  {{ box-shadow: 0 0 0 3px rgba(212,175,55,0.9), 0 0 24px rgba(212,175,55,0.5); background: rgba(212,175,55,0.18); }}
  48%  {{ box-shadow: none; background: transparent; }}
  60%  {{ box-shadow: 0 0 0 3px rgba(212,175,55,0.7), 0 0 16px rgba(212,175,55,0.35); background: rgba(212,175,55,0.12); }}
  76%  {{ box-shadow: none; background: transparent; }}
  88%  {{ box-shadow: 0 0 0 2px rgba(212,175,55,0.4), 0 0 8px rgba(212,175,55,0.2); background: rgba(212,175,55,0.06); }}
  100% {{ box-shadow: none; background: transparent; }}
}}
.flash-player {{
  animation: flash-player 2.2s ease-out;
  border-radius: 8px;
  position: relative;
  z-index: 2;
}}

@media (max-width: 480px) {{
  header {{ min-height: 64px; gap: 10px; padding: 0 12px; }}
  .logo-fallback {{ font-size: 0.9rem; }}
  .hero-label {{ font-size: 0.5rem; letter-spacing: 0.18em; }}
}}
</style>
</head>
<body>

<!-- ── Header ── -->
<header>
  <div class="header-logo">
    <img src="logo.png" class="logo-img" alt="GGP LIVE SHINJUKU">
  </div>
  <div class="header-center">
    <div class="header-title">R A N K I N G</div>
    <div class="header-sub">最終更新: {updated_at}</div>
  </div>
  <button class="search-btn" onclick="openSearch()" aria-label="検索">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
      <circle cx="11" cy="11" r="7"/><line x1="16.5" y1="16.5" x2="22" y2="22"/>
    </svg>
  </button>
</header>

<!-- ── Search overlay ── -->
<div id="searchOverlay" class="search-overlay">
  <div class="search-modal">
    <div class="search-input-wrap">
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#aaa" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
        <circle cx="11" cy="11" r="7"/><line x1="16.5" y1="16.5" x2="22" y2="22"/>
      </svg>
      <input id="searchInput" class="search-input" type="text" placeholder="名前で検索..." oninput="doSearch(this.value)" autocomplete="off">
      <button class="search-close" onclick="closeSearch()">✕</button>
    </div>
    <div id="searchResults" class="search-results">
      <div class="search-hint">名前を入力して検索</div>
    </div>
  </div>
</div>

<!-- ── Category nav ── -->
<nav class="cat-nav">
  <button class="cat-btn ring-cat active"   onclick="switchCat(this,'ring-panel')">♠ リングポイント</button>
  <button class="cat-btn daily-cat"         onclick="switchCat(this,'daily-panel')">♥ デイリーリング</button>
  <button class="cat-btn toname-cat"        onclick="switchCat(this,'toname-panel')">♦ トナメポイント</button>
</nav>

<!-- ── Ring panel ── -->
<div class="category-panel ring-panel active" id="ring-panel">
  <div class="hero ring-hero">
    <img src="title_ring.png" class="hero-title-img" alt="リングゲームポイントランキング">
  </div>
  <div class="content-wrap">{ring_html}</div>
</div>

<!-- ── Daily panel ── -->
<div class="category-panel daily-panel" id="daily-panel">
  <div class="hero daily-hero">
    <img src="title_daily.png" class="hero-title-img" alt="デイリーリングポイントランキング">
  </div>
  <div class="content-wrap">{daily_html}</div>
</div>

<!-- ── Tournament panel ── -->
<div class="category-panel toname-panel" id="toname-panel">
  <div class="hero toname-hero">
    <img src="title_toname.png" class="hero-title-img" alt="MONTHLY DEEP RANKING">
  </div>
  <div class="content-wrap">{toname_html}</div>
</div>

<script>
const SEARCH_DATA = {search_data_js};

function switchCat(btn, id) {{
  document.querySelectorAll('.cat-btn').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.category-panel').forEach(p => p.classList.remove('active'));
  btn.classList.add('active');
  var el = document.getElementById(id);
  if (el) el.classList.add('active');
}}
function switchSeason(btn, id) {{
  var wrap = btn.closest('.content-wrap');
  wrap.querySelectorAll('.season-btn').forEach(b => b.classList.remove('active'));
  wrap.querySelectorAll('.season-panel').forEach(p => p.classList.remove('active'));
  btn.classList.add('active');
  var el = document.getElementById(id);
  if (el) el.classList.add('active');
  if (btn.dataset.past === 'true') fireConfetti();
}}
function fireConfetti() {{
  var colors = ['#d4af37','#ffe066','#fff','#c0392b','#2980b9','#8e44ad'];
  confetti({{ particleCount: 80, angle: 60, spread: 55, origin: {{x: 0, y: 0.65}}, colors: colors }});
  confetti({{ particleCount: 80, angle: 120, spread: 55, origin: {{x: 1, y: 0.65}}, colors: colors }});
}}

function openSearch() {{
  document.getElementById('searchOverlay').classList.add('open');
  setTimeout(() => document.getElementById('searchInput').focus(), 50);
}}
function closeSearch() {{
  document.getElementById('searchOverlay').classList.remove('open');
  document.getElementById('searchInput').value = '';
  document.getElementById('searchResults').innerHTML = '<div class="search-hint">名前を入力して検索</div>';
}}
function doSearch(q) {{
  q = q.trim().toLowerCase();
  var el = document.getElementById('searchResults');
  if (!q) {{ el.innerHTML = '<div class="search-hint">名前を入力して検索</div>'; return; }}
  var matches = SEARCH_DATA.filter(d => d.name.toLowerCase().includes(q));
  if (!matches.length) {{ el.innerHTML = '<div class="search-empty">見つかりませんでした</div>'; return; }}
  el.innerHTML = matches.map((d, i) =>
    '<div class="search-result-item" onclick="goToResult(' + i + ')" data-idx="' + i + '">' +
    '<span class="s-cat ' + d.cat + '">' + d.catLabel + '</span>' +
    '<span class="s-season">' + d.season + '</span>' +
    '<span class="s-rank">' + d.rank + '</span>' +
    '<span class="s-name">' + d.name + '</span>' +
    '<span class="s-pts">' + d.points + '</span>' +
    '<span class="s-arrow">›</span>' +
    '</div>'
  ).join('');
  window._searchMatches = matches;
}}
function goToResult(idx) {{
  var d = window._searchMatches[idx];
  if (!d) return;
  closeSearch();
  // カテゴリ切り替え
  document.querySelectorAll('.cat-btn').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.category-panel').forEach(p => p.classList.remove('active'));
  var catPanel = document.getElementById(d.catPanel);
  if (catPanel) catPanel.classList.add('active');
  document.querySelectorAll('.cat-btn.' + d.cat + '-cat').forEach(b => b.classList.add('active'));
  // シーズンタブ切り替え
  var seasonPanel = document.getElementById(d.panelId);
  if (seasonPanel) {{
    var wrap = seasonPanel.closest('.content-wrap');
    wrap.querySelectorAll('.season-btn').forEach(b => b.classList.remove('active'));
    wrap.querySelectorAll('.season-panel').forEach(p => p.classList.remove('active'));
    seasonPanel.classList.add('active');
    var tabs = wrap.querySelectorAll('.season-btn');
    if (tabs[d.tabIdx]) tabs[d.tabIdx].classList.add('active');
  }}
  // 該当選手を中央スクロール＋ちかちかハイライト
  setTimeout(function() {{
    var panel = document.getElementById(d.panelId);
    if (!panel) return;
    var playerEl = null;
    panel.querySelectorAll('[data-name]').forEach(function(el) {{
      if (el.getAttribute('data-name') === d.name) playerEl = el;
    }});
    if (playerEl) {{
      playerEl.scrollIntoView({{ behavior: 'smooth', block: 'center' }});
      playerEl.classList.remove('flash-player');
      void playerEl.offsetWidth;
      playerEl.classList.add('flash-player');
      playerEl.addEventListener('animationend', function() {{
        playerEl.classList.remove('flash-player');
      }}, {{ once: true }});
    }}
  }}, 150);
}}
document.getElementById('searchOverlay').addEventListener('click', function(e) {{
  if (e.target === this) closeSearch();
}});
document.addEventListener('keydown', function(e) {{
  if (e.key === 'Escape') closeSearch();
}});
</script>
</body>
</html>"""


def main():
    if not SA_JSON_STR:
        print("❌ SERVICE_ACCOUNT_JSON が設定されていません")
        sys.exit(1)

    gc = get_gc()
    updated_at = datetime.now(JST).strftime("%Y/%m/%d %H:%M JST")

    print("リングポイントランキング読み込み中...")
    ring_seasons = fetch_ring_seasons(gc, RING_SHEET_ID) if RING_SHEET_ID else []

    print("デイリーリングランキング読み込み中...")
    daily_seasons = fetch_daily_ranking(gc, RING_SHEET_ID) if RING_SHEET_ID else []

    print("トナメポイントランキング読み込み中...")
    toname_seasons = fetch_toname_seasons(gc, TONAME_SHEET_ID) if TONAME_SHEET_ID else []

    print("HTML生成中...")
    html = generate_html(ring_seasons, daily_seasons, toname_seasons, updated_at)

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)

    print("✅ index.html 生成完了！")


if __name__ == "__main__":
    main()
