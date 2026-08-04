"""slides-spec.json の type:"chart" スライドに載せるグラフ画像を生成する。

    python build_chart.py           # charts/nise-keisatsu.png を生成
    python build_chart.py --verify  # 生成物の画素サイズ・フォント解決を機械検証
    python build_chart.py --all     # 生成 → 検証

配色とフォントは slides-spec.json の meta.design（build_pptx.py / build_html.py と同値）に
合わせる。強調は1色（accent）のみで、他は無彩色。凡例は置かず各スライスへ直接ラベルを打つ。
"""

import argparse
import math
import sys
from pathlib import Path

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt                                  # noqa: E402
from matplotlib import font_manager                              # noqa: E402

BASE = Path(__file__).resolve().parent
OUT_PATH = BASE / "charts" / "nise-keisatsu.png"

# ---------------------------------------------------------------- 定数
INK = "#20242C"                 # 本文インク
MUTED = "#5C6472"               # 補助インク
ACCENT = "#B9791F"              # 強調スライス
NEUTRAL = "#D9DCE2"             # 無彩色スライス
BG = "#FFFFFF"

FONT = "Yu Gothic"
FONT_FILES = ("YuGothR.ttc", "YuGothB.ttc", "YuGothM.ttc")

DPI = 200
PX_W, PX_H = 2200, 950          # 投影に耐える実ピクセル
GAP_PX = 2                      # スライス間の白い隙間（px 相当）
RADIUS, WIDTH = 0.62, 0.186     # ドーナツの外径と帯幅（内径 0.434）
XLIM, YLIM = 1.30, 1.032        # 軸範囲（比 1.26 = パネルの縦横比）
AX_RECT = (0.02, 0.0, 0.46, 0.845)   # パネル軸（左パネル。右は x に +0.50）
HEAD_Y = 0.985                  # 見出しブロックの上端（figure 座標）

# 直接ラベルの置き場所。強調はドーナツの下（下半分は必ず強調スライス）、
# 無彩は上（無彩スライスが左上に来る）。どちらも輪の外に出して重なりを避ける。
LAB_GAP = 0.05                  # 輪からラベルまでの間隔
LAB_STEP = 0.175                # ラベルの行送り
EMPH_X, REST_X = 0.25, -0.55    # ラベルブロックの水平位置

FS_HEAD = 21                    # パネル見出し
FS_HEAD_SUB = 15                # 見出しの補足
FS_HERO = 36                    # 中央のヒーロー数字
FS_HERO_SUB = 14                # 中央の名称
FS_LABEL = 18                   # 直接ラベルの名称
FS_LABEL_SUB = 16               # 直接ラベルの数値

# パネル定義：見出し / 補足（任意） / 強調 / 無彩 / 中央ヒーロー。
# 強調・無彩は (名称, 数値行, 比率)。名称の "\n" は改行。
# 数値・年次は spec の chart スライド（title / footer / notes）にある値だけを使う。
PANELS = (
    {
        "head": "特殊詐欺の被害額の内訳",
        "sub": "2025年・総額 約1,423億円",
        "emph": ("ニセ警察詐欺", "約1,005億円・70.6%", 70.6),
        "rest": ("その他の特殊詐欺", "約418億円・29.4%", 29.4),
        "hero": ("70.6%", "ニセ警察詐欺"),
    },
    {
        "head": "ニセ警察の犯行電話の発信元",
        "emph": ("「＋」で始まる国際電話", "75.5%", 75.5),
        "rest": ("その他", "24.5%", 24.5),
        "hero": ("75.5%", "国際電話"),
    },
)


# ---------------------------------------------------------------- フォント
def setup_font():
    """Yu Gothic を確実に解決する。落ちたら豆腐になるので例外で止める。"""
    if FONT not in {f.name for f in font_manager.fontManager.ttflist}:
        for name in FONT_FILES:
            path = Path("C:/Windows/Fonts") / name
            if path.exists():
                font_manager.fontManager.addfont(str(path))
    if FONT not in {f.name for f in font_manager.fontManager.ttflist}:
        raise RuntimeError("フォント '%s' が見つからない（日本語が豆腐になる）" % FONT)
    plt.rcParams["font.family"] = FONT
    plt.rcParams["axes.unicode_minus"] = False


# ---------------------------------------------------------------- 描画
def draw_panel(ax, panel):
    """単一比率のドーナツ1つ。12時から時計回り、強調スライスが先。"""
    emph, rest = panel["emph"], panel["rest"]
    wedges, _ = ax.pie(
        [emph[2], rest[2]],
        colors=[ACCENT, NEUTRAL],
        startangle=90,
        counterclock=False,
        radius=RADIUS,
        wedgeprops={"width": WIDTH, "edgecolor": BG,
                    "linewidth": GAP_PX * 72.0 / DPI},   # lw は pt なので px から換算
    )

    # 各スライスへの直接ラベル（凡例は置かない）。テキストは常にインク色。
    # 強調は輪の下、無彩は輪の上に置く。名称（インク）→数値（補助）の順で外へ積む。
    draw_label(ax, EMPH_X, -(RADIUS + LAB_GAP), emph[0], emph[1], down=True)
    draw_label(ax, REST_X, RADIUS + LAB_GAP, rest[0], rest[1], down=False)

    hero, hero_sub = panel["hero"]
    ax.text(0, 0.075, hero, ha="center", va="center",
            fontsize=FS_HERO, color=INK, fontweight="bold")
    ax.text(0, -0.14, hero_sub, ha="center", va="center",
            fontsize=FS_HERO_SUB, color=MUTED)

    ax.set_aspect("equal")
    ax.set_xlim(-XLIM, XLIM)
    ax.set_ylim(-YLIM, YLIM)
    ax.axis("off")


def draw_label(ax, x, y, name, value, down):
    """(x, y) を起点に、輪から離れる向きへ名称と数値を積む直接ラベル。"""
    lines = name.split("\n") + [value]
    sizes = [FS_LABEL] * (len(lines) - 1) + [FS_LABEL_SUB]
    colors = [INK] * (len(lines) - 1) + [MUTED]
    if not down:                        # 上側は外（上）へ向かって逆順に積む
        lines, sizes, colors = lines[::-1], sizes[::-1], colors[::-1]
    for i, (line, size, color) in enumerate(zip(lines, sizes, colors)):
        ax.text(x, y + (-1 if down else 1) * i * LAB_STEP, line,
                ha="center", va="top" if down else "bottom",
                fontsize=size, color=color, fontweight="bold" if color == INK else "normal")


def build():
    setup_font()
    fig = plt.figure(figsize=(PX_W / DPI, PX_H / DPI), dpi=DPI, facecolor=BG)
    for i, panel in enumerate(PANELS):
        cx = 0.25 + 0.50 * i
        fig.text(cx, HEAD_Y, panel["head"], ha="center", va="top",
                 fontsize=FS_HEAD, color=INK, fontweight="bold")
        if panel.get("sub"):
            # 見出し1行分（pt）を figure 座標へ換算して下にずらす
            fig.text(cx, HEAD_Y - FS_HEAD * 1.45 * (DPI / 72.0) / PX_H, panel["sub"],
                     ha="center", va="top", fontsize=FS_HEAD_SUB, color=MUTED)
        x, y, w, h = AX_RECT
        ax = fig.add_axes([x + 0.50 * i, y, w, h])
        ax.set_facecolor(BG)
        draw_panel(ax, panel)

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    # bbox_inches は指定しない（キャンバスが再計算され実ピクセルがずれる）
    fig.savefig(str(OUT_PATH), dpi=DPI, facecolor=BG)
    plt.close(fig)
    return OUT_PATH


# ---------------------------------------------------------------- 検証
def verify():
    from PIL import Image

    ok = True
    exists = OUT_PATH.exists()
    print("[1] file : %s -> %s" % (OUT_PATH, "OK" if exists else "NG"))
    ok &= exists
    if not exists:
        print("RESULT: FAIL")
        return 1

    with Image.open(OUT_PATH) as im:
        size = im.size
    size_ok = size == (PX_W, PX_H)
    print("[2] pixels : %dx%d (expected %dx%d) -> %s"
          % (size[0], size[1], PX_W, PX_H, "OK" if size_ok else "NG"))
    ok &= size_ok

    setup_font()
    font_ok = plt.rcParams["font.family"] == [FONT]
    print("[3] font family : %s -> %s"
          % (plt.rcParams["font.family"], "OK" if font_ok else "NG"))
    ok &= font_ok

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

    path = build()
    print("built: %s (%dx%d px)" % (path, PX_W, PX_H))
    if args.all:
        return verify()
    return 0


if __name__ == "__main__":
    sys.exit(main())
