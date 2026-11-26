"""EPMA（JEOL）の定性・簡易定量の出力画像から、定量表（Element / Mass% / Atom%）を読み取る。

文字認識の誤りを見逃さないように、次の3つを確かめて、合わないときは警告を返す。
  - 質量%の合計が表の Total と一致するか
  - 原子%の合計が 100 になるか
  - 各行の 質量% ÷ 原子% が、その元素の原子量に比例しているか
"""
import io
import re

import numpy as np
from PIL import Image

_NUMFIX = str.maketrans({"O": "0", "o": "0", "l": "1", "I": "1", "S": "5", ",": ".", "-": ".", " ": ""})

def _ocr_engine():
    from rapidocr_onnxruntime import RapidOCR
    return RapidOCR()

def _ocr(engine, img):
    """[{x, x2, y, h, t}] を返す（座標は渡した画像のピクセル）。"""
    res = engine(np.array(img))
    res = res[0] if isinstance(res, tuple) else res
    out = []
    for item in (res or []):
        if len(item) < 2 or not isinstance(item[0], (list, tuple, np.ndarray)):
            continue
        box, txt = item[0], item[1]
        xs = [p[0] for p in box]; ys = [p[1] for p in box]
        out.append({"x": min(xs), "x2": max(xs), "y": (min(ys) + max(ys)) / 2, "h": max(ys) - min(ys), "t": str(txt).strip()})
    return out


def _read_cell(engine, img):
    """1セルだけを文字検出なしで読む（高さ30px以下の画像は検出を省いて認識だけ行われる）。"""
    if img.width < 4 or img.height < 4:
        return ""
    k = 30 / img.height
    img = img.resize((max(8, int(img.width * k)), 30), Image.LANCZOS)
    res = engine(np.array(img))
    res = res[0] if isinstance(res, tuple) else res
    for item in (res or []):
        if len(item) >= 3:
            return str(item[1]).strip()
        if len(item) == 2 and isinstance(item[0], str):
            return item[0].strip()
    return ""


def _numtok(s):
    s2 = s.translate(_NUMFIX)
    return float(s2) if re.fullmatch(r"\d+\.\d+", s2) else None


def _lines(toks):
    if not toks:
        return []
    hmed = sorted(t["h"] for t in toks)[len(toks) // 2]
    toks = sorted(toks, key=lambda t: t["y"])
    lines, cur = [], [toks[0]]
    for t in toks[1:]:
        if abs(t["y"] - cur[-1]["y"]) <= 0.6 * hmed:
            cur.append(t)
        else:
            lines.append(sorted(cur, key=lambda t: t["x"])); cur = [t]
    lines.append(sorted(cur, key=lambda t: t["x"]))
    return lines


def image_to_rows(data, MASS, engine=None):
    """EPMA(JEOL)の定性・簡易定量の出力画像から、定量表を読み取る。
    戻り値: (rows, warnings)。rows は [{el, wt, sig, at}]。"""
    engine = engine or _ocr_engine()
    im = Image.open(io.BytesIO(data)).convert("RGB")
    W, H = im.size
    x0 = int(W * 0.78)

    # 1) 右側の欄全体から、表の見出し（Mass / Atom）と Total の行を探す
    s1 = 2 if W < 2400 else 1
    panel = im.crop((x0, 0, W, H)).resize((int((W - x0) * s1), int(H * s1)), Image.LANCZOS)
    lines = _lines(_ocr(engine, panel))
    head = next((l for l in lines if any(re.search(r"mass|atom", t["t"], re.I) for t in l)), None)
    if head is None:
        return [], ["画像の中に定量表（Element / Mass% / Atom%）を見つけられませんでした。"]
    y_head = head[0]["y"] / s1
    tot = next((l for l in lines if l[0]["y"] / s1 > y_head and any(re.search(r"tot\s*a", t["t"], re.I) for t in l)), None)
    y_tot = tot[0]["y"] / s1 if tot else H
    hh = max(6, head[0]["h"] / s1)

    # 2) 表の部分だけを拡大して読み直す
    s2 = 3 if W < 2400 else 2
    ty0, ty1 = int(y_head + hh * 0.6), int(min(H, y_tot - hh * 0.2))
    table = im.crop((x0, ty0, W, ty1))
    table = table.resize((table.width * s2, table.height * s2), Image.LANCZOS)
    toks = _ocr(engine, table)
    nums = [t for t in toks if _numtok(t["t"]) is not None]
    if len(nums) < 4:
        return [], ["定量表の数値を読み取れませんでした。"]
    # 列：数値は右揃えなので、右端の位置でまとめる
    edges = sorted(t["x2"] for t in nums)
    cols = [[edges[0]]]
    for e in edges[1:]:
        if e - cols[-1][-1] > table.width * 0.06:
            cols.append([e])
        else:
            cols[-1].append(e)
    cols = [c for c in cols if len(c) >= max(2, len(nums) // 12)]
    if len(cols) < 2:
        return [], ["定量表の列（Mass% と Atom%）を見分けられませんでした。"]
    col_x2 = [sum(c) / len(c) for c in cols[:2]]
    col_w = max(t["x2"] - t["x"] for t in nums) * 1.15
    # 行：等間隔に並ぶので、読めた行の間隔から抜けた行を補う
    ys = [l[0]["y"] for l in _lines(nums)]
    diffs = sorted(b - a for a, b in zip(ys, ys[1:]))
    d = diffs[len(diffs) // 2] if diffs else table.height
    small = [x for x in diffs if x < 1.5 * diffs[0]] if diffs else []
    if small:
        d = sum(small) / len(small)
    rows_y = [ys[0]]
    while rows_y[0] - d > d * 0.45:           # 先頭より上に読めなかった行がある場合
        rows_y.insert(0, rows_y[0] - d)
    for y in ys[1:]:
        n = max(1, round((y - rows_y[-1]) / d))
        step = (y - rows_y[-1]) / n
        for k in range(1, n + 1):
            rows_y.append(rows_y[-1] + step if k < n else y)

    def cell_value(y, ci):
        for t in nums:
            if abs(t["y"] - y) < d * 0.4 and abs(t["x2"] - col_x2[ci]) < table.width * 0.04:
                return _numtok(t["t"])
        # 読めなかったセルは、その部分だけを切り出して読み直す
        cx0 = x0 + (col_x2[ci] - col_w) / s2 - 2
        cx1 = x0 + col_x2[ci] / s2 + 3
        cy0 = ty0 + (y - d * 0.5) / s2
        cy1 = ty0 + (y + d * 0.5) / s2
        return _numtok(_read_cell(engine, im.crop((int(cx0), int(cy0), int(cx1), int(cy1)))))

    raw, lost = [], 0
    left_edge = col_x2[0] - col_w
    for y in rows_y:
        wt, at = cell_value(y, 0), cell_value(y, 1)
        if wt is None and at is None:
            continue
        if wt is None or at is None:
            lost += 1
            continue
        el = None
        for t in toks:
            if abs(t["y"] - y) < d * 0.4 and t["x2"] < left_edge + col_w * 0.3 and re.fullmatch(r"[A-Za-z]{1,2}", t["t"]):
                el = t["t"][0].upper() + t["t"][1:].lower()
        raw.append({"ocr_el": el, "wt": wt, "at": at})
    total_line = [dict(t, t=t["t"]) for t in (tot or [])]
    warns = []
    if lost:
        warns.append(f"数値を読み取れなかった行が {lost} 行あります。")
    if len(raw) < 2:
        return [], ["画像の中に定量表（Element / Mass% / Atom%）を見つけられませんでした。"]

    # 質量%÷原子% は原子量に比例する。この比から元素を決め、読み取りの誤りも検出する。
    ratios = [r["wt"] / r["at"] if r["at"] > 0 else None for r in raw]
    masses = sorted(MASS.items(), key=lambda kv: kv[1])
    def nearest(m):
        return min(masses, key=lambda kv: abs(kv[1] - m))
    def score(c):
        n = 0
        for r in ratios:
            if r is None: continue
            el, m = nearest(r * c)
            if abs(m - r * c) / m < 0.002: n += 1
        return n
    cands = []
    for r_, q in zip(raw, ratios):
        if q and r_["ocr_el"] in MASS:
            cands.append(MASS[r_["ocr_el"]] / q)
    for q in ratios:                     # 文字が読めなかった場合に備えて総当たりも試す
        if q:
            cands += [m / q for _, m in masses]
    best = max(cands, key=score) if cands else None
    rows = []
    for r_, q in zip(raw, ratios):
        el, ok = r_["ocr_el"], False
        if best and q:
            # 許容幅：表示桁（質量% 小数3桁、原子% 小数4桁）の丸めによる誤差を見込む
            tol = 0.0015 + 0.0005 / r_["wt"] + 0.00005 / r_["at"]
            m_est = q * best
            if el in MASS and abs(MASS[el] - m_est) / MASS[el] < tol:
                ok = True                      # 読めた元素記号が数値と矛盾しない
            else:
                e2, m = nearest(m_est)
                if abs(m - m_est) / m < tol and tol < 0.006:
                    el, ok = e2, True          # 数値から元素が一意に決まる
        if el is None or el not in MASS:
            warns.append(f"元素を決められない行があります（質量% {r_['wt']}、原子% {r_['at']}）。")
            el = el or "?"
        elif not ok:
            warns.append(f"{el} の行は、質量%（{r_['wt']}）と原子%（{r_['at']}）の比が原子量と合いません。数値の読み取り誤りの可能性があります。")
        rows.append({"el": el, "wt": r_["wt"], "sig": None, "at": r_["at"]})
    sw = sum(r["wt"] for r in rows); sa = sum(r["at"] for r in rows)
    tv = [_numtok(t["t"]) for t in (total_line or [])]; tv = [v for v in tv if v is not None]
    tw = tv[0] if tv else 100.0
    if abs(sw - tw) > 0.02:
        warns.append(f"質量%の合計が {sw:.3f} で、表の Total（{tw:g}）と合いません。読み落としか読み取り誤りがあります。")
    if abs(sa - 100) > 0.02:
        warns.append(f"原子%の合計が {sa:.4f} で、100 になりません。読み落としか読み取り誤りがあります。")
    dup = {r["el"] for r in rows if sum(1 for x in rows if x["el"] == r["el"]) > 1}
    if dup:
        warns.append("同じ元素が複数の行に割り当てられました：" + "・".join(sorted(dup)))
    return rows, warns
