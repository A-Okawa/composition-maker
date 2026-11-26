"""EDS定量結果 整理ツール（Streamlit版）

EDSレポート（.docx / .txt / .csv）から元素・質量%・原子%だけを取り出し、
Word や PowerPoint に表として貼り付けられる形に整える。

起動:  streamlit run app.py
"""
import html
import io
import json
import math
import re
import unicodedata
import zipfile
import xml.etree.ElementTree as ET

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st

# ---------------------------------------------------------------- 原子量
MASS = {
    "H": 1.008, "He": 4.0026, "Li": 6.94, "Be": 9.0122, "B": 10.81, "C": 12.011, "N": 14.007,
    "O": 15.999, "F": 18.998, "Ne": 20.180, "Na": 22.990, "Mg": 24.305, "Al": 26.982, "Si": 28.085,
    "P": 30.974, "S": 32.06, "Cl": 35.45, "Ar": 39.95, "K": 39.098, "Ca": 40.078, "Sc": 44.956,
    "Ti": 47.867, "V": 50.942, "Cr": 51.996, "Mn": 54.938, "Fe": 55.845, "Co": 58.933, "Ni": 58.693,
    "Cu": 63.546, "Zn": 65.38, "Ga": 69.723, "Ge": 72.630, "As": 74.922, "Se": 78.971, "Br": 79.904,
    "Kr": 83.798, "Rb": 85.468, "Sr": 87.62, "Y": 88.906, "Zr": 91.224, "Nb": 92.906, "Mo": 95.95,
    "Tc": 98, "Ru": 101.07, "Rh": 102.91, "Pd": 106.42, "Ag": 107.87, "Cd": 112.41, "In": 114.82,
    "Sn": 118.71, "Sb": 121.76, "Te": 127.60, "I": 126.90, "Xe": 131.29, "Cs": 132.91, "Ba": 137.33,
    "La": 138.91, "Ce": 140.12, "Pr": 140.91, "Nd": 144.24, "Pm": 145, "Sm": 150.36, "Eu": 151.96,
    "Gd": 157.25, "Tb": 158.93, "Dy": 162.50, "Ho": 164.93, "Er": 167.26, "Tm": 168.93, "Yb": 173.05,
    "Lu": 174.97, "Hf": 178.49, "Ta": 180.95, "W": 183.84, "Re": 186.21, "Os": 190.23, "Ir": 192.22,
    "Pt": 195.08, "Au": 196.97, "Hg": 200.59, "Tl": 204.38, "Pb": 207.2, "Bi": 208.98, "Po": 209,
    "At": 210, "Rn": 222, "Fr": 223, "Ra": 226, "Ac": 227, "Th": 232.04, "Pa": 231.04, "U": 238.03,
    "Np": 237, "Pu": 244,
}
ZNUM = {el: i + 1 for i, el in enumerate(MASS)}


# ---------------------------------------------------------------- 解析
def norm(s):
    """半角カナなどを正規化して前後の空白を落とす。"""
    return unicodedata.normalize("NFKC", "" if s is None else str(s)).replace(" ", " ").strip()


def key(s):
    return re.sub(r"[\s_.]", "", norm(s).lower())


_NUM = re.compile(r"^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$", re.I)


def num(s):
    t = norm(s).replace(",", "")
    return float(t) if t and _NUM.match(t) else None


def is_el(s):
    return bool(re.fullmatch(r"[A-Z][a-z]?", norm(s)))


def is_sigma(k):
    return bool(re.search(r"(シグマ|sigma|σ|誤差|error)", k))


def is_wt(k):
    return bool(re.search(r"(質量|重量|wt|weight|mass)", k)) and "%" in k


def is_at(k):
    return ("原子" in k or bool(re.match(r"at(omic|om)?\W*%", k))) and "%" in k


def is_total(k):
    return bool(re.fullmatch(r"(トータル|total|合計|計)[::]?", k))


def text_to_grid(text):
    """テキストを行×セルに分ける（タブ区切り、カンマ区切り、空白区切りに対応）。"""
    out = []
    for line in norm(text.replace("\r", "")).split("\n"):
        if "\t" in line:
            out.append([norm(c) for c in line.split("\t")])
            continue
        line = norm(line)
        if not line:
            out.append([])
            continue
        if "," in line and not re.search(r"\s{2,}", line):
            out.append([norm(c) for c in line.split(",")])
            continue
        # 空白区切りの場合、語の途中に空白を含む用語を先に連結しておく
        line = re.sub(r"\b([KLMN])\s*(シリーズ|series|線|lines?)", r"\1_series", line, flags=re.I)
        line = re.sub(r"(質量|重量|原子|Wt|Weight|At|Atomic|Mass)\s*%\s*(シグマ|sigma|σ)", r"\1%_sigma", line, flags=re.I)
        line = re.sub(r"k\s*(レシオ|ratio)", "k_ratio", line, flags=re.I)
        line = re.sub(r"Line\s+Type", "LineType", line, flags=re.I)
        line = re.sub(r"Apparent\s+Concentration", "ApparentConcentration", line, flags=re.I)
        line = re.sub(r"(Standard|Factory)\s+(Label|Standard|Calibration\s+Date)",
                      lambda m: re.sub(r"\s+", "_", m.group(0)), line, flags=re.I)
        line = re.sub(r"Pure\s+Element", "Pure_Element", line, flags=re.I)
        line = re.sub(r"\s+\((v|V)\)", r"(\1)", line)
        out.append([norm(c) for c in line.split()])
    return out


def grid_to_spectra(grid):
    """行×セルから定量結果の表を探し、[{title, rows:[{el, wt, sig, at}]}] を返す。"""
    res, cur, cmap, title = [], None, None, None

    def close():
        nonlocal cur, cmap
        if cur and cur["rows"]:
            res.append(cur)
        cur, cmap = None, None

    for raw in grid:
        cells = [norm(c) for c in raw]
        filled = [c for c in cells if c]
        if not filled:
            continue
        hi = next((i for i, c in enumerate(cells) if re.fullmatch(r"(元素|element|elt\.?|el\.?)", c, re.I)), -1)
        if hi >= 0 and len(filled) > 1:
            close()
            cmap = {"el": hi, "wt": -1, "sig": -1, "at": -1}
            for i, c in enumerate(cells):
                k = key(c)
                if is_sigma(k):
                    if is_wt(k) and cmap["sig"] < 0:
                        cmap["sig"] = i
                elif is_wt(k):
                    if cmap["wt"] < 0:
                        cmap["wt"] = i
                elif is_at(k):
                    if cmap["at"] < 0:
                        cmap["at"] = i
            cur = {"title": title, "rows": []}
            title = None
            continue
        if cmap:
            el = cells[cmap["el"]] if cmap["el"] < len(cells) else ""
            if is_el(el):
                g = lambda i: num(cells[i]) if 0 <= i < len(cells) else None
                cur["rows"].append({"el": el, "wt": g(cmap["wt"]), "sig": g(cmap["sig"]), "at": g(cmap["at"])})
                continue
            if is_total(key(cells[0] or filled[0])):
                close()
                continue
            close()
        if len(filled) <= 4 and not all(num(c) is not None for c in filled) and not (is_el(filled[0]) and len(filled) > 1):
            title = " ".join(filled)
    close()

    if not res:  # 見出しなし:「元素 数値」の並びを質量%として読む
        rows = []
        for raw in grid:
            c = [x for x in (norm(v) for v in raw) if x]
            if len(c) >= 2 and is_el(c[0]):
                v = next((n for n in (num(x) for x in c[1:]) if n is not None), None)
                if v is not None:
                    rows.append({"el": c[0], "wt": v, "sig": None, "at": None})
        if len(rows) >= 2:
            res.append({"title": None, "rows": rows, "guess": True})
    return [s for s in res if any(r["wt"] is not None or r["at"] is not None for r in s["rows"])]


def _local(tag):
    return tag.rsplit("}", 1)[-1]


def docx_to_grid(data):
    """Word文書の本文を上から順に、段落は1セルの行、表は行×セルとして取り出す。"""
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    body = next((e for e in root.iter() if _local(e.tag) == "body"), None)
    grid = []

    def text_of(node):
        paras = []
        for p in node.iter():
            if _local(p.tag) == "p":
                paras.append("".join(t.text or "" for t in p.iter() if _local(t.tag) == "t"))
        return " ".join(x for x in paras if x)

    def own(node, name, stop):
        """行やセルがコンテンツコントロール(sdt)で包まれていても拾う。入れ子の表は除く。"""
        found = []
        for ch in node:
            ln = _local(ch.tag)
            if ln == name:
                found.append(ch)
            elif ln != stop:
                found.extend(own(ch, name, stop))
        return found

    def walk(node):
        for ch in node:
            ln = _local(ch.tag)
            if ln == "tbl":
                for tr in own(ch, "tr", "tbl"):
                    grid.append([text_of(tc) for tc in own(tr, "tc", "tr")])
                grid.append([])
            elif ln == "p":
                t = "".join(x.text or "" for x in ch.iter() if _local(x.tag) == "t")
                grid.append([t] if t else [])
            elif ln in ("sdt", "sdtContent"):
                walk(ch)

    if body is not None:
        walk(body)
    return grid


def decode_text(data):
    if data[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return data.decode("utf-16")
    for enc in ("utf-8-sig", "cp932"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            pass
    return data.decode("utf-8", errors="replace")


IMAGE_EXT = ("jpg", "jpeg", "png", "bmp", "tif", "tiff")


def load_image(name, data, engine=None):
    """EPMA の出力画像 → ([{name, src, rows}], 警告のリスト)。"""
    try:
        import epma_image
    except ImportError:
        raise ValueError("画像の読み取りには rapidocr-onnxruntime が必要です（pip install -r requirements.txt を実行してください）")
    try:
        rows, warns = epma_image.image_to_rows(data, MASS, engine)
    except ImportError:
        raise ValueError("画像の読み取りには rapidocr-onnxruntime が必要です（pip install -r requirements.txt を実行してください）")
    if not rows:
        raise ValueError(warns[0] if warns else "画像から定量表を読み取れませんでした")
    base = re.sub(r"\.[^.]+$", "", name)
    return [{"name": base, "src": name + "（画像から読み取り）", "rows": rows}], warns


def load_file(name, data):
    """1ファイル → [{name, src, rows}]。読めないときは ValueError。"""
    low = name.lower()
    base = re.sub(r"\.[^.]+$", "", name)
    if low.endswith(".docx"):
        try:
            grid = docx_to_grid(data)
        except (zipfile.BadZipFile, KeyError, ET.ParseError):
            raise ValueError("Word文書として読めませんでした")
    elif re.search(r"\.(doc|xlsx?|pdf|pptx?)$", low):
        raise ValueError("この形式には対応していません（.docx / .txt / .csv を使ってください）")
    else:
        grid = text_to_grid(decode_text(data))
    sp = grid_to_spectra(grid)
    if not sp:
        raise ValueError("定量結果の表（元素と質量%または原子%の列）が見つかりませんでした")
    out = []
    for i, s in enumerate(sp):
        nm = f"{base} – {s['title'] or i + 1}" if len(sp) > 1 else base
        src = name + (f"（{s['title']}）" if s["title"] else "")
        if s.get("guess"):
            src += "／見出しなしのため数値を質量%として読み込み"
        out.append({"name": nm, "src": src, "rows": s["rows"]})
    return out


# ---------------------------------------------------------------- 計算
def compute(rows_all, excluded):
    """除外・再規格化・原子%換算。{el: {wt, sig, at}} を返す。"""
    rows = [r for r in rows_all if r["el"] not in excluded]
    was_excluded = len(rows) != len(rows_all)
    known = all(r["el"] in MASS for r in rows)
    out = {}
    if any(r["wt"] is not None for r in rows):
        W = sum(r["wt"] or 0 for r in rows)
        f = 100 / W if (was_excluded and W > 0) else 1
        A = sum((r["wt"] or 0) / MASS[r["el"]] for r in rows if r["el"] in MASS)
        for r in rows:
            out[r["el"]] = {
                "wt": None if r["wt"] is None else r["wt"] * f,
                "sig": None if r["sig"] is None else r["sig"] * f,
                "at": (r["wt"] / MASS[r["el"]]) / A * 100 if (known and A > 0 and r["wt"] is not None) else None,
            }
    else:  # 質量%の列がない → 原子%から質量%を換算
        T = sum(r["at"] or 0 for r in rows)
        W = sum((r["at"] or 0) * MASS[r["el"]] for r in rows if r["el"] in MASS)
        for r in rows:
            out[r["el"]] = {
                "at": None if (r["at"] is None or T <= 0) else r["at"] / T * 100,
                "sig": None,
                "wt": r["at"] * MASS[r["el"]] / W * 100 if (known and W > 0 and r["at"] is not None) else None,
            }
    return out


LABELS = {
    False: {"el": "Element", "wt": "wt%", "sig": "σ (wt%)", "at": "at%", "total": "Total", "mean": "Mean", "sd": "SD"},
    True: {"el": "元素", "wt": "質量%", "sig": "σ (質量%)", "at": "原子%", "total": "合計", "mean": "平均", "sd": "標準偏差"},
}


def build_table(spectra, qs, excluded, dec=2, orient=False, total=True, stats=False, grid=False, jp=False):
    """spectra: [{name, rows}]。貼り付け用HTMLとタブ区切りテキストを返す。"""
    lab = LABELS[jp]
    act = [{"name": s["name"], "c": compute(s["rows"], excluded)} for s in spectra]
    els = sorted({el for a in act for el in a["c"]}, key=lambda e: ZNUM.get(e, 999))
    if not act or not qs or not els:
        return None

    def getter(a):
        return lambda el, q: (a["c"].get(el) or {}).get(q)

    groups = [{"name": a["name"], "qs": qs, "get": getter(a), "sum": True} for a in act]
    if stats and len(act) >= 2:
        sq = [q for q in qs if q != "sig"]

        def vals(el, q):  # あるスペクトルに無い元素は 0 として扱う
            vs = [(a["c"][el][q] if el in a["c"] else 0) for a in act]
            return [v for v in vs if v is not None]

        def mean(el, q):
            v = vals(el, q)
            return sum(v) / len(v) if v else None

        def sd(el, q):
            v = vals(el, q)
            if len(v) < 2:
                return None
            m = sum(v) / len(v)
            return math.sqrt(sum((x - m) ** 2 for x in v) / (len(v) - 1))

        if sq:
            groups.append({"name": lab["mean"], "qs": sq, "get": mean, "sum": True})
            groups.append({"name": lab["sd"], "qs": sq, "get": sd, "sum": False})

    nG, nQ = len(groups), len(qs)
    cols = []
    for g in groups:
        for q in g["qs"]:
            flat = lab[q] if nG == 1 else (g["name"] if nQ == 1 else f"{g['name']} {lab[q]}")
            tot = sum(g["get"](el, q) or 0 for el in els) if (g["sum"] and q != "sig") else None
            cols.append({"g": g, "q": q, "flat": flat, "total": tot})

    def fmt(v):
        return "–" if v is None or not math.isfinite(v) else f"{v:.{dec}f}"

    head, body = [], []
    if not orient:
        corner = lab["el"] + (f" ({lab[qs[0]]})" if (nG > 1 and nQ == 1) else "")
        if nG > 1 and nQ > 1:
            head.append([{"t": lab["el"], "rs": 2, "left": 1}] + [{"t": g["name"], "cs": len(g["qs"]), "under": 1} for g in groups])
            head.append([{"t": lab[c["q"]]} for c in cols])
        else:
            head.append([{"t": corner, "left": 1}] + [{"t": c["flat"]} for c in cols])
        for el in els:
            body.append([{"t": el, "left": 1}] + [{"t": fmt(c["g"]["get"](el, c["q"]))} for c in cols])
        if total:
            body.append([{"t": lab["total"], "left": 1, "total": 1}] +
                        [{"t": "" if c["total"] is None else fmt(c["total"]), "total": 1} for c in cols])
        tsv = [[corner] + [c["flat"] for c in cols]] + [[c["t"] for c in r] for r in body]
    else:
        corner = lab[qs[0]] if (nG > 1 and nQ == 1) else ""
        head.append([{"t": corner, "left": 1}] + [{"t": e} for e in els] + ([{"t": lab["total"]}] if total else []))
        for c in cols:
            body.append([{"t": c["flat"], "left": 1}] + [{"t": fmt(c["g"]["get"](el, c["q"]))} for el in els] +
                        ([{"t": "" if c["total"] is None else fmt(c["total"])}] if total else []))
        tsv = [[c["t"] for c in head[0]]] + [[c["t"] for c in r] for r in body]

    THICK, THIN, HAIR = "1.5pt solid #000", "0.75pt solid #000", "0.5pt solid #000"
    base = "padding:2pt 8pt;font-family:Arial,'Yu Gothic',sans-serif;font-size:10.5pt;color:#000;white-space:nowrap;"

    def cell(c, is_h, first, last_h, last_b):
        st_ = base + f"text-align:{'left' if c.get('left') else 'center'};font-weight:{'bold' if is_h else 'normal'};"
        if grid:
            st_ += f"border:{THIN};"
        else:
            if is_h and first:
                st_ += f"border-top:{THICK};"
            if is_h and (last_h or c.get("rs")):
                st_ += f"border-bottom:{THIN};"
            elif is_h and c.get("under"):
                st_ += f"border-bottom:{HAIR};"
            if c.get("total"):
                st_ += f"border-top:{THIN};"
            if last_b:
                st_ += f"border-bottom:{THICK};"
        tag = "th" if is_h else "td"
        span = (f' colspan="{c["cs"]}"' if c.get("cs", 1) > 1 else "") + (f' rowspan="{c["rs"]}"' if c.get("rs", 1) > 1 else "")
        return f'<{tag}{span} style="{st_}">{html.escape(c["t"])}</{tag}>'

    h = '<table style="border-collapse:collapse;border-spacing:0;"><thead>'
    for i, r in enumerate(head):
        h += "<tr>" + "".join(cell(c, True, i == 0, i == len(head) - 1, False) for c in r) + "</tr>"
    h += "</thead><tbody>"
    for i, r in enumerate(body):
        h += "<tr>" + "".join(cell(c, False, False, False, i == len(body) - 1) for c in r) + "</tr>"
    h += "</tbody></table>"
    return {"html": h, "tsv": "\n".join("\t".join(r) for r in tsv), "n_rows": len(head) + len(body)}


# ---------------------------------------------------------------- 棒グラフ
SERIES_COLORS = ["#257de0", "#f47230", "#00b894", "#f2a600", "#ee7fac", "#008a17", "#6a4ce2", "#ec4647"]
INK, INK2, GRIDC = "#191919", "#666666", "#e3e3e3"


CHART_DEFAULTS = {
    "font": "Arial", "fontsize": 10, "width_cm": 16.0, "height_cm": 9.0, "bar_width": 0.78,
    "frame": False, "edge": False, "edge_color": "#000000", "grid": True, "legend": True,
    "ymax": 0.0, "colors": [],
}
FONT_FALLBACK = ["Yu Gothic", "Meiryo", "Hiragino Sans", "Noto Sans CJK JP", "DejaVu Sans"]


def chart_series(spectra, q, excluded, stats=False, jp=False):
    """棒グラフの系列を作る。(元素の並び, [(名前, 値, 誤差)], 平均表示かどうか) を返す。
    スペクトルが2つ以上で stats=True、または9つ以上のときは平均±標準偏差の1系列にまとめる。"""
    lab = LABELS[jp]
    act = [{"name": s["name"], "c": compute(s["rows"], excluded)} for s in spectra]
    els = sorted({el for a in act for el in a["c"]}, key=lambda e: ZNUM.get(e, 999))
    if not act or not els:
        return [], [], False
    use_mean = len(act) >= 2 and (stats or len(act) > len(SERIES_COLORS))
    series = []
    if use_mean:
        means, sds = [], []
        for el in els:
            v = [(a["c"][el][q] if el in a["c"] else 0) for a in act]
            v = [x for x in v if x is not None]
            m = sum(v) / len(v) if v else 0
            means.append(m)
            sds.append(math.sqrt(sum((x - m) ** 2 for x in v) / (len(v) - 1)) if len(v) > 1 else 0)
        series.append((f"{lab['mean']} ± {lab['sd']} (n={len(act)})", means, sds))
    else:
        for a in act:
            vals = [((a["c"].get(el) or {}).get(q) or 0) for el in els]
            err = None
            if q == "wt" and any((a["c"].get(el) or {}).get("sig") for el in els):
                err = [((a["c"].get(el) or {}).get("sig") or 0) for el in els]
            series.append((a["name"], vals, err))
    if not any(any(v) for _, v, _ in series):
        return [], [], False
    return els, series, use_mean


def build_chart(spectra, q, excluded, stats=False, jp=False, opts=None, fmt="png", dpi=300):
    """元素ごとの棒グラフを画像(bytes)で返す。q は "at" か "wt"、fmt は "png" か "tiff"。"""
    o = dict(CHART_DEFAULTS, **(opts or {}))
    lab = LABELS[jp]
    els, series, use_mean = chart_series(spectra, q, excluded, stats, jp)
    if not series:
        return None
    fs = o["fontsize"]
    rc = {"font.family": [o["font"]] + FONT_FALLBACK, "font.size": fs, "axes.unicode_minus": False}
    with plt.rc_context(rc):
        n, k = len(els), len(series)
        fig, ax = plt.subplots(figsize=(o["width_cm"] / 2.54, o["height_cm"] / 2.54), dpi=dpi)
        group_w = o["bar_width"]
        bw = group_w / k
        for i, (name, vals, err) in enumerate(series):
            color = o["colors"][i] if i < len(o["colors"]) else SERIES_COLORS[i % len(SERIES_COLORS)]
            xs = [j - group_w / 2 + bw * (i + 0.5) for j in range(n)]
            ax.bar(xs, vals, width=bw * (1.0 if o["edge"] else 0.9), color=color, label=name, zorder=3,
                   edgecolor=o["edge_color"] if o["edge"] else "none", linewidth=0.6 if o["edge"] else 0,
                   yerr=err, error_kw={"ecolor": INK, "elinewidth": 0.8, "capsize": 2, "zorder": 4})
        ax.set_xticks(range(n))
        ax.set_xticklabels(els, color="#000000", fontsize=fs)
        ax.set_ylabel(lab[q], color="#000000", fontsize=fs)
        ax.set_ylim(bottom=0, top=o["ymax"] if o["ymax"] and o["ymax"] > 0 else None)
        ax.margins(x=0.02)
        ax.yaxis.grid(bool(o["grid"]), color=GRIDC, linewidth=0.8, zorder=0)
        ax.set_axisbelow(True)
        if o["frame"]:  # 四辺を黒い枠線で囲み、目盛りを外向きに付ける
            for side in ("top", "right", "left", "bottom"):
                ax.spines[side].set_visible(True)
                ax.spines[side].set_color("#000000")
                ax.spines[side].set_linewidth(0.8)
            ax.tick_params(axis="both", colors="#000000", length=3, width=0.8, direction="out", labelsize=fs - 1)
            ax.tick_params(axis="x", labelsize=fs)
        else:
            for side in ("top", "right", "left"):
                ax.spines[side].set_visible(False)
            ax.spines["bottom"].set_color(INK2)
            ax.tick_params(axis="y", colors=INK2, length=0, labelsize=fs - 1)
            ax.tick_params(axis="x", length=0)
        if o["legend"] and (k > 1 or use_mean):
            ax.legend(frameon=False, fontsize=fs - 1, labelcolor="#000000", loc="upper right" if n > 6 else "best")
        fig.tight_layout()
        buf = io.BytesIO()
        if fmt == "tiff":
            fig.savefig(buf, format="tiff", dpi=dpi, facecolor="white", pil_kwargs={"compression": "tiff_lzw"})
        else:
            fig.savefig(buf, format="png", dpi=dpi, facecolor="white")
        plt.close(fig)
    return buf.getvalue()


# ---------------------------------------------------------------- 画面
COPY_COMPONENT = """
<div style="font-family:'Yu Gothic UI','Hiragino Sans',Meiryo,sans-serif;">
  <div style="display:flex;gap:12px;align-items:center;flex-wrap:wrap;margin-bottom:12px;">
    <button id="copy" style="font:inherit;font-weight:600;cursor:pointer;background:#1c5cc0;color:#fff;border:1px solid #1c5cc0;border-radius:6px;padding:8px 16px;">表をコピー（Word・PowerPoint用）</button>
    <span id="status" role="status" style="font-size:13px;color:#1c5cc0;"></span>
  </div>
  <div style="overflow-x:auto;background:#fff;border:1px solid #d3dae4;border-radius:6px;padding:20px;">
    <div id="out">__TABLE__</div>
  </div>
</div>
<script>
const html = __HTML__, tsv = __TSV__;
const status = document.getElementById("status");
const say = t => { status.textContent = t; };
function selectCopy(){
  const node = document.querySelector("#out table"); if(!node) return false;
  const sel = window.getSelection(), r = document.createRange();
  r.selectNode(node); sel.removeAllRanges(); sel.addRange(r);
  let ok = false; try { ok = document.execCommand("copy"); } catch(e) {}
  if (ok) sel.removeAllRanges();
  return ok;
}
document.getElementById("copy").addEventListener("click", async () => {
  try {
    if (!(navigator.clipboard && window.ClipboardItem)) throw new Error("no api");
    await navigator.clipboard.write([new ClipboardItem({
      "text/html": new Blob(['<meta charset="utf-8">' + html], {type: "text/html"}),
      "text/plain": new Blob([tsv], {type: "text/plain"})})]);
    say("コピーしました。Word や PowerPoint に貼り付けてください。");
  } catch (e) {
    if (selectCopy()) say("コピーしました。Word や PowerPoint に貼り付けてください。");
    else say("自動でコピーできませんでした。表が選択された状態なので Ctrl+C（Mac は ⌘C）でコピーしてください。");
  }
});
</script>
"""


@st.cache_resource(show_spinner=False)
def _ocr_engine():
    from rapidocr_onnxruntime import RapidOCR
    return RapidOCR()


@st.cache_data(show_spinner="画像から定量表を読み取っています…")
def _load_image_cached(name, data):
    return load_image(name, data, _ocr_engine())


def main():
    st.set_page_config(page_title="composition-maker", layout="wide")
    st.title("composition-maker（EDS・EPMA 定量結果の整理）")
    st.caption("EDS のレポート（Word / txt）や EPMA の出力画像から元素・質量%・原子%だけを取り出し、Word や PowerPoint に貼り付けられる表と棒グラフにします。")

    left, right = st.columns([1, 2], gap="large")

    with left:
        st.subheader("1. 読み込み")
        files = st.file_uploader("EDS レポート、EPMA の出力画像（複数可）", type=["docx", "txt", "csv", "tsv", *IMAGE_EXT],
                                 accept_multiple_files=True)
        pasted = st.text_area("または、表やテキストをコピーして貼り付け", height=110,
                              placeholder="元素\tラインタイプ\t…\t質量%\t…\nO\tK シリーズ\t…\t22.12\t…")

        spectra, errors, n_img = [], [], 0
        for f in files or []:
            try:
                if f.name.lower().rsplit(".", 1)[-1] in IMAGE_EXT:
                    sp, warns = _load_image_cached(f.name, f.getvalue())
                    spectra += sp
                    n_img += 1
                    errors += [f"{f.name}：{w}" for w in warns]
                else:
                    spectra += load_file(f.name, f.getvalue())
            except ValueError as e:
                errors.append(f"{f.name}：{e}")
        if pasted.strip():
            sp = grid_to_spectra(text_to_grid(pasted))
            if not sp:
                errors.append("貼り付けた内容から、元素と質量%（または原子%）の列を見つけられませんでした。見出し行ごとコピーしてください。")
            for i, s in enumerate(sp):
                spectra.append({"name": s["title"] or f"貼り付け {i + 1}", "src": "貼り付け", "rows": s["rows"]})
        for e in errors:
            st.warning(e)
        if n_img:
            st.caption("画像は文字認識で読み取っています。合計と原子量による照合で誤りを検出しますが、"
                       "使う前に元の画像の数値と見比べてください。")

        if spectra:
            st.subheader("読み込んだスペクトル")
            df = pd.DataFrame({
                "含める": [True] * len(spectra),
                "表での名前": [s["name"] for s in spectra],
                "読み込み元": [f"{s['src']}／{len(s['rows'])}元素" for s in spectra],
            })
            sig = "|".join(s["src"] + s["name"] for s in spectra)  # 入力が変わったら編集内容をリセット
            edited = st.data_editor(df, hide_index=True, disabled=["読み込み元"],
                                    key=f"specs_{abs(hash(sig))}")
            chosen = [{"name": str(edited["表での名前"][i]), "rows": s["rows"]}
                      for i, s in enumerate(spectra) if bool(edited["含める"][i])]
        else:
            chosen = []

    all_els = sorted({r["el"] for s in spectra for r in s["rows"]}, key=lambda e: ZNUM.get(e, 999))

    with st.sidebar:
        st.header("2. 表に載せる内容")
        q_names = {"原子%": "at", "質量%": "wt", "σ（質量%）": "sig"}
        picked = st.multiselect("項目", list(q_names), default=["原子%"])
        qs = [q for q in ("wt", "sig", "at") if q in {q_names[p] for p in picked}]
        excluded = set(st.multiselect("除外する元素（残りを100%に再規格化）", all_els))
        dec = st.number_input("小数点以下の桁数", 0, 4, 2)
        st.markdown("**体裁**")
        orient = st.toggle("元素を横に並べる")
        total = st.toggle("合計を表示", value=True)
        stats = st.toggle("平均・標準偏差")
        grid = st.toggle("格子罫線")
        jp = st.toggle("見出しを日本語")
        st.markdown("**棒グラフ**")
        show_chart = st.toggle("棒グラフを表示", value=True)
        chart_qs = [q for q in ("at", "wt") if q in qs]
        chart_q = chart_qs[0] if chart_qs else None
        if show_chart and len(chart_qs) > 1:
            chart_q = st.radio("グラフの値", chart_qs, format_func=lambda q: {"at": "原子%", "wt": "質量%"}[q], horizontal=True)
        copts = dict(CHART_DEFAULTS)
        if show_chart:
            with st.expander("グラフの調整"):
                copts["font"] = st.selectbox("フォント", ["Arial", "Helvetica", "Times New Roman"])
                copts["fontsize"] = st.number_input("文字サイズ (pt)", 5, 24, 10)
                copts["width_cm"] = st.slider("幅 (cm)", 6.0, 30.0, 16.0, 0.5)
                copts["height_cm"] = st.slider("高さ (cm)", 4.0, 20.0, 9.0, 0.5)
                copts["bar_width"] = st.slider("棒の太さ", 0.2, 1.0, 0.78, 0.02)
                copts["ymax"] = st.number_input("縦軸の最大値（0 で自動）", 0.0, 100.0, 0.0, 5.0)
                copts["frame"] = st.toggle("枠線で囲む")
                copts["edge"] = st.toggle("棒に縁取りを付ける")
                if copts["edge"]:
                    copts["edge_color"] = st.color_picker("縁取りの色", "#000000")
                copts["grid"] = st.toggle("横の目盛線", value=True)
                copts["legend"] = st.toggle("凡例を表示", value=True)

    with right:
        st.subheader("3. 貼り付け用の表")
        if not spectra:
            st.info("レポートを読み込むと、ここに表が表示されます。")
            return
        if not chosen:
            st.info("表に含めるスペクトルにチェックを入れてください。")
            return
        if not qs:
            st.info("サイドバーの「項目」から質量%または原子%を選んでください。")
            return
        t = build_table(chosen, qs, excluded, int(dec), orient, total, stats, grid, jp)
        if t is None:
            st.info("すべての元素が除外されています。")
            return

        page = (COPY_COMPONENT.replace("__TABLE__", t["html"])
                .replace("__HTML__", json.dumps(t["html"]).replace("</", "<\\/"))
                .replace("__TSV__", json.dumps(t["tsv"]).replace("</", "<\\/")))
        height = 130 + 27 * t["n_rows"]
        if hasattr(st, "iframe"):
            st.iframe(page, height=height)
        else:  # 古いStreamlit
            import streamlit.components.v1 as components
            components.html(page, height=height, scrolling=True)

        notes = []
        from_at = any(not any(r["wt"] is not None for r in s["rows"]) for s in chosen)
        if "at" in qs and not from_at:
            notes.append("原子%は質量%を原子量で換算した値です。")
        if from_at:
            notes.append("質量%の列がないスペクトルは、原子%から質量%を換算しています。")
        if excluded:
            notes.append("・".join(sorted(excluded, key=lambda e: ZNUM.get(e, 999))) + " を除外し、残りを100%に再規格化しています"
                         + ("（σも同じ係数で換算）。" if "sig" in qs else "。"))
        unknown = sorted({r["el"] for s in chosen for r in s["rows"] if r["el"] not in MASS and r["el"] not in excluded})
        if unknown:
            notes.append("原子量が不明な記号（" + "・".join(unknown) + "）があるため換算できません。除外すると計算できます。")
        if stats and len(chosen) >= 2:
            sets = [{r["el"] for r in s["rows"]} - excluded for s in chosen]
            if any(x != set.union(*sets) for x in sets):
                notes.append("平均・標準偏差では、検出されていない元素を0として計算しています。")
        if stats and len(chosen) < 2:
            notes.append("平均・標準偏差は2つ以上のスペクトルで表示されます。")
        if notes:
            st.caption(" ".join(notes))

        if show_chart and chart_q:
            _, series, use_mean = chart_series(chosen, chart_q, excluded, stats, jp)
            if series:
                st.subheader("4. 棒グラフ")
                st.caption("棒の色")
                pick_cols = st.columns(min(len(series), 4))
                colors = []
                for i, (name, _v, _e) in enumerate(series):
                    with pick_cols[i % len(pick_cols)]:
                        colors.append(st.color_picker(name, SERIES_COLORS[i % len(SERIES_COLORS)], key=f"bar_color_{i}"))
                copts["colors"] = colors
                st.image(build_chart(chosen, chart_q, excluded, stats, jp, copts, "png", 200))
                st.caption("画像を右クリックして「画像をコピー」で Word や PowerPoint に貼り付けられます。"
                           + ("誤差棒は標準偏差です。" if use_mean else ("誤差棒は質量%のσです。" if chart_q == "wt" else "")))
                c1, c2, c3 = st.columns([1, 1, 2])
                fmt = c1.selectbox("保存形式", ["TIFF", "PNG"])
                dpi = c2.selectbox("解像度 (dpi)", [300, 600, 1200], index=1)
                with c3:
                    st.write("")
                    st.write("")
                    if st.button("保存用の画像を作る"):
                        data = build_chart(chosen, chart_q, excluded, stats, jp, copts, fmt.lower(), int(dpi))
                        st.session_state["chart_file"] = (data, fmt, int(dpi))
                saved = st.session_state.get("chart_file")
                if saved:
                    data, f_, d_ = saved
                    ext = "tif" if f_ == "TIFF" else "png"
                    st.download_button(f"{f_}（{d_} dpi、{len(data) / 1e6:.1f} MB）をダウンロード", data,
                                       file_name=f"composition_bar_{d_}dpi.{ext}",
                                       mime="image/tiff" if f_ == "TIFF" else "image/png")
                    st.caption("設定を変えたあとは、もう一度「保存用の画像を作る」を押してください。")

        with st.expander("テキストでコピー（Excel用・タブ区切り）"):
            st.code(t["tsv"], language=None)
            st.download_button("CSVをダウンロード", "﻿" + "\n".join(",".join(r.split("\t")) for r in t["tsv"].split("\n")),
                               file_name="eds_table.csv", mime="text/csv")


if __name__ == "__main__":
    main()
