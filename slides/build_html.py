"""slides-spec.json から GitHub Pages 用のスライド閲覧ページ slides.html を生成する。

    python build_html.py            # 生成
    python build_html.py --verify   # 生成物を機械検証
    python build_html.py --all      # 生成 → 検証

出力は完全自己完結（外部CDN・外部fetchなし、CSS/JS はインライン）。
レイアウト定数は build_pptx.py と同じ 13.333 x 7.5 inch 座標系で持ち、
CSS へ出す際にカード幅基準の相対値（% / cqw）へ変換する。
"""

import argparse
import base64
import html
import json
import math
import re
import sys
import unicodedata
from pathlib import Path

BASE = Path(__file__).resolve().parent
SPEC_PATH = BASE / "slides-spec.json"
OUT_PATH = BASE.parent / "slides.html"        # リポジトリルート直下

PPTX_HREF = "slides/deepfake-talk.pptx"
INDEX_HREF = "./"

# ---------------------------------------------------------------- レイアウト定数
# （build_pptx.py と同値。単位は inch）
SLIDE_W, SLIDE_H = 13.333, 7.5
ML = 0.75                       # 左右マージン
CW = SLIDE_W - ML * 2           # 本文幅 11.833
INSET = 0.2                     # テキストボックス左右インセット合計

CONTENT_TOP = 1.80
CONTENT_BOTTOM = 6.55
FOOTER_Y = 6.40
BIG_TOP = 1.30
BIG_W = 12.233

# image_slot（あとから画像を貼るための予約領域）。CSS 側の .slot と対応。
SLOT_LABEL = "イメージ画像スペース"
SLOT_LABEL_PT, SLOT_DESC_PT = 15, 12
SLOT_MAIN_R = 0.62              # pos:right のときの本文幅（CW 比）
SLOT_BAND_H, SLOT_BAND_GAP = 1.35, 0.22   # pos:bottom の帯高さとすき間

# sources（出典）: 表示名＋URL の2行1組。CSS の .src / .src-url の line-height と同値。
SRC_NAME_LS, SRC_URL_LS, SRC_GAP_R = 1.25, 1.15, 0.30
SRC_SIZES = (15, 14, 13, 12)    # 表示名の候補pt（15pt が全名を1行に収める上限）


# ---------------------------------------------------------------- 文字幅の見積り
def em_width(text):
    """全角=1.0em / 半角=0.55em として文字列の幅を em で返す。"""
    w = 0.0
    for ch in text:
        w += 1.0 if unicodedata.east_asian_width(ch) in ("W", "F", "A") else 0.55
    return w


def est_lines(text, size_pt, width_in):
    """size_pt の文字を width_in に流し込んだときの行数見積り。"""
    per_line = width_in * 72.0 / size_pt
    if per_line <= 0:
        return 1
    return max(1, math.ceil(em_width(text) / per_line))


# ---------------------------------------------------------------- 出力ヘルパ
def esc(text):
    return html.escape(str(text), quote=True)


def num(v):
    return "%g" % v


def fs(pt):
    """カード幅に追従するフォントサイズ（1pt = カード幅の 0.10417%）。

    実際の font-size は CSS 側の .fz が var(--size) から計算する。変数で持たせるのは、
    狭い画面で大見出しだけ上限を掛けられるようにするため。
    """
    return "--size:%s" % num(pt)


def esc_multiline(text):
    """spec 内の \\n を <br> にした上でエスケープする。"""
    return "<br>".join(esc(line) for line in str(text).split("\n"))


def title_html(title):
    """アクセント色の見出し＋下線。サイズは build_pptx.py と同じ判定。"""
    size = 40 if em_width(title) <= 18 else 36
    return ('<div class="head"><p class="title fz" style="%s">%s</p></div>'
            '<div class="rule"></div>' % (fs(size), esc(title)))


def footer_html(footer, center=False):
    cls = "s-footer center fz" if center else "s-footer fz"
    return '<p class="%s" style="%s">%s</p>' % (cls, fs(15), esc(footer))


def emph_html(text, size=26, danger=False):
    cls = "emph danger fz" if danger else "emph fz"
    return '<p class="%s" style="%s">%s</p>' % (cls, fs(size), esc(text))


def image_slot(s, pos):
    """指定位置の image_slot を返す（無ければ None）。"""
    slot = s.get("image_slot")
    return slot if slot and slot.get("pos") == pos else None


def slot_html(slot, bottom=False):
    """あとから画像を貼るための予約領域（破線の仮置き枠）。"""
    out = ['<div class="slot%s">' % (" slot-bottom" if bottom else ""),
           '<p class="slot-lab fz" style="%s">%s</p>' % (fs(SLOT_LABEL_PT), esc(SLOT_LABEL))]
    if slot.get("desc"):
        out.append('<p class="slot-desc fz" style="%s">%s</p>'
                   % (fs(SLOT_DESC_PT), esc(slot["desc"])))
    out.append("</div>")
    return "".join(out)


def source_sizes(items, avail_h):
    """(表示名, URL) の2行1組を avail_h に収める (表示名pt, URLpt) を返す。"""
    size, url_size = SRC_SIZES[-1], 10
    for cand in SRC_SIZES:
        u = 11 if cand >= 14 else 10
        text_w = CW - INSET - cand / 72.0      # ぶら下げインデント分を差し引く
        name_lines = sum(est_lines(n, cand, text_w) for n, _ in items)
        url_lines = sum(est_lines(url, u, text_w) for _, url in items)
        need = (name_lines * cand * SRC_NAME_LS + url_lines * u * SRC_URL_LS
                + (len(items) - 1) * cand * SRC_GAP_R) / 72
        if need <= avail_h:
            size, url_size = cand, u
            break
    return size, url_size


def source_list_html(items, size, url_size):
    """出典を「表示名／URL」の2行1組のリンクにしたリスト。"""
    lis = []
    for name, url in items:
        lis.append('<li><a class="src-a" href="%s" target="_blank" rel="noopener">'
                   '<span class="src-name">%s</span>'
                   '<span class="src-url fz" style="%s">%s</span></a></li>'
                   % (esc(url), esc(name), fs(url_size), esc(url)))
    style = "%s;row-gap:calc(var(--pt) * %s)" % (fs(size), num(round(size * SRC_GAP_R, 2)))
    return '<ul class="bul src fz" style="%s">%s</ul>' % (style, "".join(lis))


def bullet_list_html(items, size, gap_ratio):
    """「・」つきのぶら下げリスト。gap_ratio は pptx の space_after 比。"""
    lis = "".join("<li>%s</li>" % esc(t) for t in items)
    style = "%s;row-gap:calc(var(--pt) * %s)" % (fs(size), num(round(size * gap_ratio, 2)))
    return '<ul class="bul fz" style="%s">%s</ul>' % (style, lis)


# ---------------------------------------------------------------- 各タイプの描画
def render_cover(s):
    body = ['<div class="cover">',
            '<p class="c-title fz" style="%s">%s</p>' % (fs(48), esc(s["title"]))]
    if s.get("sub"):
        body.append('<p class="c-sub fz" style="%s">%s</p>' % (fs(28), esc(s["sub"])))
    body.append("</div>")
    if s.get("footer"):
        body.append(footer_html(s["footer"], center=True))
    return "".join(body), ""


def render_section(s):
    body = []
    if s.get("kicker"):
        body.append('<p class="kicker fz" style="%s">%s</p>' % (fs(20), esc(s["kicker"])))
    title = s["title"]
    size = 54 if em_width(title) * 54 / 72 <= CW - INSET else 44
    body.append('<div class="sec-body">')
    body.append('<p class="sec-title fz" style="%s">%s</p>' % (fs(size), esc(title)))
    if s.get("sub"):
        body.append('<p class="sec-sub fz" style="%s">%s</p>' % (fs(24), esc(s["sub"])))
    body.append("</div>")
    return "".join(body), ""


def render_big(s):
    bottom = CONTENT_BOTTOM
    if s.get("warn"):
        bottom = bottom - 1.05 - 0.25
    slot = image_slot(s, "bottom")
    if slot:
        bottom = bottom - SLOT_BAND_H - SLOT_BAND_GAP
    if s.get("sub"):
        bottom = bottom - 1.10 - 0.15

    lines = str(s["big"]).split("\n")
    max_em = max(em_width(x) for x in lines)
    avail_w = BIG_W - INSET
    avail_h = bottom - BIG_TOP
    size = 44
    for cand in range(54, 43, -1):
        if (max_em * cand / 72 <= avail_w
                and len(lines) * cand * 1.28 / 72 <= avail_h):
            size = cand
            break

    body = ['<div class="bigwrap">',
            '<p class="big fz" style="%s">%s</p>' % (fs(size), esc_multiline(s["big"]))]
    if s.get("sub"):
        body.append('<p class="big-sub fz" style="%s">%s</p>' % (fs(22), esc(s["sub"])))
    body.append("</div>")
    if slot:
        body.append(slot_html(slot, bottom=True))

    tail = ""
    if s.get("warn"):
        tail = '<div class="warn fz" style="%s">%s</div>' % (fs(28), esc(s["warn"]))
    return "".join(body), tail


def render_bullets(s):
    body = [title_html(s["title"])]
    bottom = CONTENT_BOTTOM
    if s.get("footer"):
        bottom = FOOTER_Y - 0.08
    # プレースホルダは本文領域（emphasis を含む）と同じ高さ帯に置く。
    slot = image_slot(s, "right")
    body_w = CW * SLOT_MAIN_R if slot else CW
    if s.get("emphasis"):
        emph_h = 0.85
        if slot:
            # 幅が狭まって折り返す分、強調ボックスを縦に伸ばす（文字はみ出し防止）。
            emph_h = max(emph_h, est_lines(s["emphasis"], 26, body_w - 0.6)
                         * 26 * 1.35 / 72 + 0.22)
        bottom = bottom - emph_h - 0.18

    items = s["bullets"]
    avail_h = bottom - CONTENT_TOP
    size = 24
    fit = False
    for cand in (tuple(range(28, 17, -1)) if slot else (28, 26, 24)):
        text_w = body_w - INSET - cand / 72.0  # ぶら下げインデント分を差し引く
        lines = sum(est_lines(t, cand, text_w) for t in items)
        need = lines * cand * 1.35 / 72 + (len(items) - 1) * (0.35 * cand / 72)
        if need <= avail_h:
            size, fit = cand, True
            break
    if slot and not fit:
        raise ValueError("n=%s: バレットが幅 %.2fin の本文領域に収まらない" % (s["n"], body_w))

    if slot:
        main = [bullet_list_html(items, size, 0.35)]
        if s.get("emphasis"):
            main.append(emph_html(s["emphasis"], 26))
        body.append('<div class="content split"><div class="split-main">%s</div>%s</div>'
                    % ("".join(main), slot_html(slot)))
    else:
        body.append('<div class="content">%s</div>' % bullet_list_html(items, size, 0.35))
        if s.get("emphasis"):
            body.append(emph_html(s["emphasis"], 26))
    if s.get("footer"):
        body.append(footer_html(s["footer"]))
    return "".join(body), ""


_VALUE_ACCENT = r"(「[^」]*」|『[^』]*』|約?[0-9０-９][0-9０-９,，\.]*[万億兆]?円?)"


def value_html(value):
    """値の中の金額・鍵カッコ語をアクセント色にする（build_pptx.py と同じ規則）。"""
    out = []
    for part in [x for x in re.split(_VALUE_ACCENT, value) if x]:
        hit = bool(re.fullmatch(_VALUE_ACCENT, part))
        out.append('<span class="%s">%s</span>' % ("acc" if hit else "v", esc(part)))
    return "".join(out)


def render_stat(s):
    body = [title_html(s["title"])]
    rows = []
    for label, value in s["stats"]:
        rows.append('<div class="stat-row">'
                    '<p class="stat-lab fz" style="%s">%s</p>'
                    '<p class="stat-val fz" style="%s">%s</p></div>'
                    % (fs(20), esc(label), fs(28), value_html(value)))
    body.append('<div class="content center"><div class="stat">%s</div></div>' % "".join(rows))
    if s.get("emphasis"):
        body.append(emph_html(s["emphasis"], 28, danger=True))
    if s.get("footer"):
        body.append(footer_html(s["footer"]))
    return "".join(body), ""


def render_steps(s):
    body = [title_html(s["title"])]
    bottom = CONTENT_BOTTOM
    if s.get("footer"):
        bottom = FOOTER_Y - 0.08
    if s.get("emphasis"):
        bottom = bottom - 0.85 - 0.18

    rows = s["steps"]
    wide_label = max(em_width(r[0]) for r in rows) > 5
    main_x = 5.35 if wide_label else 1.65
    main_w = SLIDE_W - ML - main_x
    avail_h = bottom - CONTENT_TOP

    def row_heights(main_s, note_s):
        hs = []
        for r in rows:
            h = est_lines(r[1], main_s, main_w - INSET) * main_s * 1.20 / 72
            if len(r) > 2 and r[2]:
                h += est_lines(r[2], note_s, main_w - INSET) * note_s * 1.30 / 72 + 0.06
            hs.append(h)
        return hs

    combos = [(m, n) for m in (30, 28, 26) for n in (20, 19, 18)]
    main_s, note_s = combos[-1]
    for cand_main, cand_note in combos:
        hs = row_heights(cand_main, cand_note)
        if sum(hs) + (len(rows) - 1) * 0.14 <= avail_h:
            main_s, note_s = cand_main, cand_note
            break

    lab_s = 20 if wide_label else 28
    cells = []
    for r in rows:
        cells.append('<p class="lab fz" style="%s">%s</p>' % (fs(lab_s), esc(r[0])))
        cell = ['<div class="step">',
                '<p class="main fz" style="%s">%s</p>' % (fs(main_s), esc(r[1]))]
        if len(r) > 2 and r[2]:
            cell.append('<p class="note fz" style="%s">%s</p>' % (fs(note_s), esc(r[2])))
        cell.append("</div>")
        cells.append("".join(cell))

    body.append('<div class="content center"><div class="steps %s">%s</div></div>'
                % ("wide" if wide_label else "narrow", "".join(cells)))
    if s.get("emphasis"):
        body.append(emph_html(s["emphasis"], 26))
    if s.get("footer"):
        body.append(footer_html(s["footer"]))
    return "".join(body), ""


def render_chart(s):
    """見出し＋グラフ画像＋footer。画像は data URI で埋め込む（外部参照ゼロを保つ）。"""
    path = BASE / s["image"]
    if not path.exists():
        raise FileNotFoundError(
            "グラフ画像がない: %s（python build_chart.py で生成する）" % path)
    uri = "data:image/png;base64," + base64.b64encode(path.read_bytes()).decode("ascii")
    alt = s.get("alt") or CHART_ALT.get(s["image"]) or s["title"]

    body = [title_html(s["title"]),
            '<div class="content center">'
            '<img class="chart-img" src="%s" alt="%s"></div>' % (uri, esc(alt))]
    if s.get("footer"):
        body.append(footer_html(s["footer"]))
    return "".join(body), ""


def render_sources(s):
    body = [title_html(s["title"])]
    bottom = FOOTER_Y - 0.08 if s.get("footer") else CONTENT_BOTTOM
    avail_h = bottom - CONTENT_TOP

    if s.get("sources"):
        size, url_size = source_sizes(s["sources"], avail_h)
        body.append('<div class="content">%s</div>'
                    % source_list_html(s["sources"], size, url_size))
        if s.get("footer"):
            body.append(footer_html(s["footer"]))
        return "".join(body), ""

    items = s["bullets"]        # 旧形式（URL なしの文字列リスト）
    size = 16
    for cand in (18, 17, 16):
        text_w = CW - INSET - cand / 72.0
        lines = sum(est_lines(t, cand, text_w) for t in items)
        need = lines * cand * 1.30 / 72 + (len(items) - 1) * (0.45 * cand / 72)
        if need <= avail_h:
            size = cand
            break

    body.append('<div class="content">%s</div>' % bullet_list_html(items, size, 0.45))
    if s.get("footer"):
        body.append(footer_html(s["footer"]))
    return "".join(body), ""


RENDERERS = {
    "cover": render_cover,
    "section": render_section,
    "big": render_big,
    "bullets": render_bullets,
    "stat": render_stat,
    "steps": render_steps,
    "chart": render_chart,
    "sources": render_sources,
}

# グラフ画像の代替テキスト（spec に "alt" があればそちらが優先）。
CHART_ALT = {
    "charts/nise-keisatsu.png":
        "円グラフ: 特殊詐欺被害額の70.6%がニセ警察詐欺、犯行電話の75.5%が国際電話",
}


# ---------------------------------------------------------------- ページ組み立て
def slide_label(s):
    """目次に出す短いラベル。"""
    if s.get("title"):
        return s["title"]
    if s.get("big"):
        return str(s["big"]).split("\n")[0]
    return s.get("kicker", "")


def render_slide(s, total):
    body, tail = RENDERERS[s["type"]](s)
    return "".join([
        '<section class="slide" id="s%d">' % s["n"],
        '<div class="stage t-%s">' % s["type"],
        '<span class="pageno fz" style="%s">%d / %d</span>' % (fs(13), s["n"], total),
        '<div class="body">%s</div>' % body,
        tail,
        "</div>",
        '<div class="under">',
        '<span class="part">%s</span>' % esc(s.get("part", "")),
        "<details><summary>\U0001F4DD 台本・発表者ノート</summary>",
        '<p class="notes">%s</p></details>' % esc(s.get("notes") or ""),
        "</div>",
        "</section>",
    ])


def render_toc(slides):
    groups = []
    for s in slides:
        part = s.get("part", "")
        if not groups or groups[-1][0] != part:
            groups.append((part, []))
        groups[-1][1].append(s)

    out = ['<nav class="toc" aria-label="目次"><h2>目次</h2><div class="toc-groups">']
    for part, items in groups:
        out.append('<div class="toc-group"><p class="toc-part">%s</p><ul>' % esc(part))
        for s in items:
            out.append('<li><a href="#s%d"><span class="num">%d</span>%s</a></li>'
                       % (s["n"], s["n"], esc(slide_label(s))))
        out.append("</ul></div>")
    out.append("</div></nav>")
    return "".join(out)


CSS = """
:root{
  --bg:#eef0f3; --panel:#ffffff; --ink:#20242c; --muted:#5c6472;
  --line:#d9dce2; --line-soft:#e7e9ee; --chip:#f1f2f5;
  --accent:#b9791f; --accent-soft:#f3e6cf;
  --shadow:0 1px 2px rgba(20,24,34,.06),0 8px 24px -16px rgba(20,24,34,.25);
  --serif:"Hiragino Mincho ProN","Yu Mincho","YuMincho",serif;
  --sans:"Hiragino Kaku Gothic ProN","Yu Gothic","YuGothic","Noto Sans JP",system-ui,sans-serif;
  --mono:ui-monospace,"SFMono-Regular","Consolas",monospace;
}
@media (prefers-color-scheme:dark){
  :root{
    --bg:#14171d; --panel:#1c2029; --ink:#e9ebef; --muted:#9aa3b2;
    --line:#2b3038; --line-soft:#242932; --chip:#242a33;
    --accent:#e0a63e; --accent-soft:#3a2f18;
    --shadow:0 1px 2px rgba(0,0,0,.4),0 10px 28px -18px rgba(0,0,0,.7);
  }
}

*{box-sizing:border-box}
html,body{margin:0;padding:0}
body{background:var(--bg);color:var(--ink);font-family:var(--sans);
  line-height:1.8;-webkit-font-smoothing:antialiased}
.wrap{max-width:1040px;margin:0 auto;padding:clamp(18px,4vw,44px) clamp(12px,3vw,28px) 96px}
a{color:var(--accent);text-decoration:underline;text-underline-offset:3px;text-decoration-thickness:1px}
a:focus-visible,summary:focus-visible{outline:2px solid var(--accent);outline-offset:3px}

/* ---------- ページヘッダー ---------- */
.eyebrow{font-family:var(--mono);font-size:.72rem;letter-spacing:.2em;
  text-transform:uppercase;color:var(--accent);margin:0 0 12px}
h1{font-family:var(--serif);font-weight:600;font-size:clamp(1.45rem,4.2vw,2.2rem);
  line-height:1.35;margin:0 0 16px;text-wrap:balance;letter-spacing:.01em}
.actions{display:flex;flex-wrap:wrap;gap:10px;margin:0}
.btn{display:inline-block;font-size:.88rem;text-decoration:none;color:var(--ink);
  background:var(--panel);border:1px solid var(--line);border-radius:999px;
  padding:7px 16px;box-shadow:var(--shadow)}
.btn:hover{background:var(--accent-soft)}
.btn.primary{border-color:var(--accent);color:var(--accent);font-weight:600}
.hint{font-size:.85rem;color:var(--muted);margin:14px 0 0}

/* ---------- 目次 ---------- */
.toc{margin:32px 0 40px;background:var(--panel);border:1px solid var(--line);
  border-radius:14px;box-shadow:var(--shadow);padding:18px clamp(14px,3vw,24px)}
.toc h2{font-family:var(--serif);font-size:1.08rem;font-weight:600;margin:0 0 14px}
.toc-groups{display:grid;gap:16px}
.toc-part{font-family:var(--mono);font-size:.7rem;letter-spacing:.14em;
  color:var(--accent);margin:0 0 7px}
.toc ul{list-style:none;margin:0;padding:0;display:flex;flex-wrap:wrap;gap:6px}
.toc li a{display:inline-flex;align-items:baseline;gap:7px;max-width:100%;
  font-size:.84rem;line-height:1.5;text-decoration:none;color:var(--ink);
  background:var(--chip);border-radius:8px;padding:5px 11px}
.toc li a:hover{background:var(--accent-soft)}
.toc .num{font-family:var(--mono);font-size:.72rem;font-weight:700;color:var(--accent);
  font-variant-numeric:tabular-nums}

/* ---------- スライドカード ---------- */
.deck{display:grid;gap:34px}
.slide{scroll-margin-top:16px}

/* スライド内部は「スクリーン投影の再現」として常にライト配色（ダークモードでも変えない）。
   --pt は 1pt 相当の長さ。1pt = 13.333in カード幅の 0.10417%。 */
.stage{
  --s-bg:#ffffff; --s-ink:#20242c; --s-muted:#5c6472; --s-dim:#5c6472;
  --s-accent:#b9791f; --s-danger:#a83b3b; --s-emph:#f3e6cf; --s-rule:#e6e1d6;
  --pt:max(0.10417cqw,0.7px);
  container-type:inline-size;
  position:relative;
  aspect-ratio:16 / 9;
  display:flex;flex-direction:column;
  background:var(--s-bg);color:var(--s-ink);
  border:1px solid var(--line);border-radius:14px;overflow:hidden;box-shadow:var(--shadow)}
.stage.t-section{--s-bg:#20242c;--s-ink:#ffffff;--s-muted:#bcbdc0;--s-dim:#9aa3b2}
/* --size は pptx と同じ pt 値。カード幅に比例して拡縮する。 */
.fz{font-size:calc(var(--pt) * var(--size))}

.pageno{position:absolute;z-index:2;top:calc(var(--pt) * 12);right:calc(var(--pt) * 16);
  font-family:var(--mono);color:var(--s-dim);letter-spacing:.06em;line-height:1;
  font-variant-numeric:tabular-nums}
.body{flex:1;min-height:0;display:flex;flex-direction:column;
  padding:max(3.375%,14px) max(5.625%,16px) max(5.6%,18px)}
.body>p{margin:0}

.head{display:flex;align-items:flex-end;min-height:calc(var(--pt) * 76);
  padding-right:max(9%,46px)}
.title{margin:0;font-weight:700;color:var(--s-accent);line-height:1.25;letter-spacing:.01em}
.rule{flex:none;width:19.5%;height:calc(var(--pt) * 4);margin-top:calc(var(--pt) * 7);
  background:var(--s-accent);border-radius:1px}
/* bullets/sources は pptx と同じく上詰め、stat/steps は上下中央（.center）。 */
.content{flex:1;min-height:0;margin-top:calc(var(--pt) * 10);
  display:flex;flex-direction:column;justify-content:flex-start}
.content.center{justify-content:center}
.s-footer{margin-top:auto;padding-top:calc(var(--pt) * 8);
  color:var(--s-muted);line-height:1.4}
.s-footer.center{text-align:center}

/* cover / big：游明朝の大文字中央 */
.cover,.bigwrap{flex:1;display:flex;flex-direction:column;justify-content:center;
  align-items:center;text-align:center}
.cover{gap:calc(var(--pt) * 22)}
.bigwrap{gap:calc(var(--pt) * 18)}
.c-title,.big{margin:0;font-family:var(--serif);font-weight:700;line-height:1.3;
  letter-spacing:.01em;text-wrap:balance}
.big{line-height:1.28}
.c-sub{margin:0;color:var(--s-accent);line-height:1.4}
.big-sub{margin:0;color:var(--s-muted);line-height:1.35}

/* warn：#a83b3b の白文字帯（フルブリード） */
.warn{flex:none;min-height:calc(var(--pt) * 76);padding:calc(var(--pt) * 12) 6%;
  background:var(--s-danger);color:#fff;font-weight:700;line-height:1.35;
  display:flex;align-items:center;justify-content:center;text-align:center}

/* section：濃色背景＋白文字＋アクセント kicker */
.kicker{margin:0;font-weight:700;color:var(--s-accent);line-height:1.3;letter-spacing:.08em}
.sec-body{flex:1;display:flex;flex-direction:column;justify-content:center;
  align-items:center;text-align:center;gap:calc(var(--pt) * 16)}
.sec-title{margin:0;font-weight:700;color:var(--s-ink);line-height:1.3;letter-spacing:.02em}
.sec-sub{margin:0;color:var(--s-muted);line-height:1.4}

/* bullets / sources：「・」つきぶら下げリスト */
.bul{list-style:none;margin:0;padding:0;display:grid}
.bul li{position:relative;padding-left:1em;line-height:1.35;color:var(--s-ink)}
.bul li::before{content:"\\30FB";position:absolute;left:0;top:0}

/* sources：表示名＋URL の2行1組リンク。line-height と row-gap は
   build_pptx.py / build_html.py の収まり判定（SRC_NAME_LS / SRC_URL_LS / SRC_GAP_R）と同値。 */
.src li{line-height:1.25}
.src-a{display:block;color:var(--s-ink);text-decoration:none;overflow-wrap:anywhere}
.src-a:hover .src-name{text-decoration:underline}
.src-url{display:block;color:var(--s-muted);line-height:1.15;overflow-wrap:anywhere}

/* emphasis：#f3e6cf のボックス */
.emph{flex:none;margin:calc(var(--pt) * 13) 0 0;min-height:calc(var(--pt) * 61);
  padding:calc(var(--pt) * 8) 3%;background:var(--s-emph);color:var(--s-ink);
  border-radius:calc(var(--pt) * 10);font-weight:700;line-height:1.35;
  display:flex;align-items:center;justify-content:center;text-align:center}
.emph.danger{color:var(--s-danger)}

/* image_slot：あとから画像を貼るための予約領域（破線の仮置き枠。本文より目立たせない） */
.slot{flex:none;display:flex;flex-direction:column;align-items:center;justify-content:center;
  text-align:center;gap:calc(var(--pt) * 3);color:var(--s-muted);
  background:#f1f2f5;border:1px dashed #5c6472;border-radius:calc(var(--pt) * 10);
  padding:calc(var(--pt) * 10) 4%}
.slot-lab{margin:0;font-weight:700;line-height:1.3}
.slot-desc{margin:0;line-height:1.4}
/* pos:right（bullets）：本文 62% ＋ すき間 4% ＋ プレースホルダ 34%。
   見出しが2行に折り返すカードでは本文帯が pptx より下から始まるため、
   flex-basis を auto にして帯を縮めない（emphasis が footer に重なるのを防ぐ）。 */
.content.split{flex:1 0 auto;flex-direction:row;column-gap:4%;align-items:stretch}
.split-main{flex:0 0 62%;min-width:0;display:flex;flex-direction:column;
  row-gap:calc(var(--pt) * 13)}
.split-main .emph{margin-top:auto}
.content.split .slot{flex:0 0 34%}
/* pos:bottom（big）：全幅 1.35in ＝ 97pt 相当の横長帯 */
.slot-bottom{min-height:calc(var(--pt) * 97);margin-top:calc(var(--pt) * 16)}

/* chart：グラフ画像（data URI）。本文領域に収める＝縦横とも上限を掛ける */
.chart-img{display:block;margin:auto;max-width:100%;max-height:100%;
  width:auto;height:auto;object-fit:contain}

/* stat：ラベル＋太字値 */
.stat{display:grid}
.stat-row{display:grid;grid-template-columns:31.3% 1fr;column-gap:2.5%;
  align-items:baseline;padding:calc(var(--pt) * 9) 0;border-bottom:1px solid var(--s-rule)}
.stat-row:last-child{border-bottom:0}
.stat-lab{margin:0;color:var(--s-muted);line-height:1.25}
.stat-val{margin:0;font-weight:700;line-height:1.2}
.stat-val .acc{color:var(--s-accent)}

/* steps：帯ラベル＋主文＋補足 */
.steps{display:grid;row-gap:calc(var(--pt) * 14);align-items:baseline}
.steps.wide{grid-template-columns:36.3% 1fr;column-gap:2.5%}
.steps.narrow{grid-template-columns:6.8% 1fr;column-gap:.9%}
.steps .lab{margin:0;color:var(--s-accent);font-weight:700;line-height:1.15}
.steps.narrow .lab{text-align:right}
.steps .main{margin:0;font-weight:700;color:var(--s-ink);line-height:1.2}
.steps .note{margin:calc(var(--pt) * 4) 0 0;color:var(--s-muted);line-height:1.3}

/* ---------- カード下（part ラベル＋発表者ノート） ---------- */
.under{display:flex;flex-direction:column;gap:8px;margin-top:12px}
.part{font-family:var(--mono);font-size:.72rem;letter-spacing:.14em;color:var(--muted)}
details{background:var(--panel);border:1px solid var(--line);border-radius:12px;
  box-shadow:var(--shadow)}
summary{cursor:pointer;list-style:none;padding:10px 15px;
  font-size:.87rem;font-weight:600;color:var(--muted)}
summary::-webkit-details-marker{display:none}
summary::before{content:"\\25B8";margin-right:8px;color:var(--accent)}
details[open] summary::before{content:"\\25BE"}
details[open] summary{border-bottom:1px solid var(--line-soft)}
.notes{margin:0;padding:13px 16px 17px;font-size:.9rem;line-height:1.9;
  color:var(--ink);overflow-wrap:anywhere}

footer.page{margin-top:48px;padding-top:18px;border-top:1px solid var(--line);
  font-size:.82rem;color:var(--muted)}
footer.page p{margin:0 0 6px}

/* ---------- 狭い画面：16:9 を解除して縦に伸ばす ---------- */
@media (max-width:759px){
  /* 16:9 は外すが、文字数の少ないスライドが薄くなりすぎないよう下限は残す */
  .stage{aspect-ratio:auto;min-height:52vw}
  .head{min-height:0}
  /* 大見出しだけ上限を掛け、狭い幅で1文字だけ折り返すのを防ぐ */
  .c-title,.big,.sec-title{font-size:calc(var(--pt) * min(var(--size),38))}
  .warn{font-size:calc(var(--pt) * min(var(--size),22))}
  .steps.wide{grid-template-columns:1fr;row-gap:calc(var(--pt) * 16)}
  .steps.wide .lab{margin-bottom:calc(var(--pt) * 3)}
  .stat-row{grid-template-columns:1fr;row-gap:calc(var(--pt) * 3)}
  /* right スロットは本文の下へ回す */
  .content.split{flex-direction:column;row-gap:calc(var(--pt) * 16)}
  .split-main{flex:auto}
  .content.split .slot{flex:none;min-height:calc(var(--pt) * 120)}
  .toc li a{font-size:.8rem}
}

/* コンテナクエリ非対応ブラウザ向けの退避 */
@supports not (container-type:inline-size){
  .stage{--pt:0.8px;aspect-ratio:auto}
}
"""


def build():
    spec = json.loads(SPEC_PATH.read_text(encoding="utf-8"))
    slides = spec["slides"]
    total = len(slides)
    deck_title = "%s ― スライド(全%d枚)" % (spec["meta"]["title"], total)

    page = "".join([
        "<!doctype html>\n",
        '<html lang="ja">\n<head>\n',
        '<meta charset="utf-8">\n',
        '<meta name="viewport" content="width=device-width,initial-scale=1">\n',
        "<title>%s</title>\n" % esc(deck_title),
        '<meta name="description" content="%s">\n'
        % esc("シニア向け60分セミナー「%s」の投影スライド全%d枚。各スライドに台本・発表者ノートつき。"
              % (spec["meta"]["title"], total)),
        "<style>%s</style>\n" % CSS,
        "</head>\n<body>\n",
        '<div class="wrap">\n',
        "<header>",
        '<p class="eyebrow">Slides ・ %s ・ 全%d枚</p>' % (esc(spec["meta"]["aspect"]), total),
        "<h1>%s</h1>" % esc(deck_title),
        '<div class="actions">',
        '<a class="btn" href="%s">← 構成案ページへ戻る</a>' % INDEX_HREF,
        '<a class="btn primary" href="%s" download>PowerPoint (.pptx) をダウンロード</a>'
        % PPTX_HREF,
        "</div>",
        '<p class="hint">各カードは投影時の見た目（16:9）です。カード下の'
        '「\U0001F4DD 台本・発表者ノート」を開くと、そのスライドで話す内容と注意点が読めます。</p>',
        "</header>\n",
        render_toc(slides),
        "\n",
        '<main class="deck">\n',
        "\n".join(render_slide(s, total) for s in slides),
        "\n</main>\n",
        '<footer class="page">',
        "<p>スライドの正本は <code>slides/slides-spec.json</code>。"
        "このページは <code>slides/build_html.py</code> で、"
        "PowerPoint は <code>slides/build_pptx.py</code> で生成しています。</p>",
        "<p>数値・出典はスライド化の直前に最新の一次資料で再確認してください。</p>",
        "</footer>\n",
        "</div>\n</body>\n</html>\n",
    ])

    OUT_PATH.write_text(page, encoding="utf-8")
    return total


# ---------------------------------------------------------------- 検証
def verify():
    spec = json.loads(SPEC_PATH.read_text(encoding="utf-8"))
    expected = len(spec["slides"])
    doc = OUT_PATH.read_text(encoding="utf-8")
    ok = True

    # (a) スライドカードが expected 個
    cards = re.findall(r'<section class="slide" id="s(\d+)">', doc)
    stages = len(re.findall(r'<div class="stage t-', doc))
    a_ok = (len(cards) == expected and stages == expected
            and [int(x) for x in cards] == [s["n"] for s in spec["slides"]])
    print("[a] slide cards : %d sections / %d stages (expected %d) -> %s"
          % (len(cards), stages, expected, "OK" if a_ok else "NG"))
    ok &= a_ok

    # (b) 全カードに notes の details がある（中身が空でないこと）
    notes = re.findall(r"<details><summary>[^<]*台本・発表者ノート</summary>"
                       r'<p class="notes">(.*?)</p></details>', doc, re.S)
    empty = [i + 1 for i, t in enumerate(notes) if not t.strip()]
    b_ok = len(notes) == expected and not empty
    print("[b] notes <details> : %d (expected %d)%s -> %s"
          % (len(notes), expected,
             "" if not empty else " empty=%s" % empty, "OK" if b_ok else "NG"))
    ok &= b_ok

    # (c) 外部URL参照が 0 件（self-contained）
    # <a href> の画面遷移はリソース読み込みではないので、出典リンク（アンカー＋表示URL）は除いて判定する。
    probe = re.sub(r"<a\b[^>]*href=\"https?://[^\"]*\"[^>]*>.*?</a>", "", doc, flags=re.S)
    attrs = re.findall(r'(?:src|href)\s*=\s*"(https?://[^"]*)"', probe)
    schemes = re.findall(r"https?://", probe)
    imports = re.findall(r"@import", probe)
    c_ok = not attrs and not schemes and not imports
    print("[c] external refs : %d src/href, %d url tokens, %d @import -> %s"
          % (len(attrs), len(schemes), len(imports), "OK" if c_ok else "NG"))
    ok &= c_ok

    # (d) pptx リンクと戻るリンク
    has_pptx = ('href="%s"' % PPTX_HREF) in doc
    has_back = ('href="%s"' % INDEX_HREF) in doc
    d_ok = has_pptx and has_back
    print("[d] links : pptx=%s back=%s -> %s"
          % (has_pptx, has_back, "OK" if d_ok else "NG"))
    ok &= d_ok

    # (e) 画像プレースホルダが image_slot 指定のカードだけにある
    want = [s["n"] for s in spec["slides"] if s.get("image_slot")]
    got = [int(n) for n, blk in
           re.findall(r'<section class="slide" id="s(\d+)">(.*?)</section>', doc, re.S)
           if '<div class="slot' in blk]
    e_ok = bool(want) and want == got
    print("[e] image placeholder cards : %s (expected %s) -> %s"
          % (got, want, "OK" if e_ok else "NG"))
    ok &= e_ok

    # (f) sources カードの出典リンクが spec の URL と一致し、target/rel つきか
    blocks = dict(re.findall(r'<section class="slide" id="s(\d+)">(.*?)</section>', doc, re.S))
    src_slides = [s for s in spec["slides"] if s.get("sources")]
    bad_src, n_links = [], 0
    for s in src_slides:
        anchors = re.findall(r'<a class="src-a" href="([^"]+)" target="_blank" rel="noopener">',
                             blocks.get(str(s["n"]), ""))
        n_links += len(anchors)
        if anchors != [esc(u) for _, u in s["sources"]]:
            bad_src.append(s["n"])
    f_ok = not bad_src                          # 旧 bullets 形式だけなら 0 件で OK
    print("[f] source links on %s : %d <a href> (expected %d) -> %s%s"
          % ([s["n"] for s in src_slides], n_links,
             sum(len(s["sources"]) for s in src_slides),
             "OK" if f_ok else "NG", "" if f_ok else " bad=%s" % bad_src))
    ok &= f_ok

    print("RESULT: %s" % ("PASS" if ok else "FAIL"))
    return 0 if ok else 1


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("--verify", action="store_true", help="生成物を検証する")
    ap.add_argument("--all", action="store_true", help="生成してから検証する")
    args = ap.parse_args()

    if args.verify and not args.all:
        return verify()

    n = build()
    print("built: %s (%d slides)" % (OUT_PATH, n))
    if args.all:
        return verify()
    return 0


if __name__ == "__main__":
    sys.exit(main())
