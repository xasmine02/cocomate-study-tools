# -*- coding: utf-8 -*-

__version__ = "2.2.0"

import argparse
import hashlib
import html as html_mod
import io
import json
import math
import os
import re
import sys
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, datetime, time

try:
    import openpyxl
    from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula
except ImportError:
    for _stream in (sys.stdout, sys.stderr):
        try:
            _stream.reconfigure(encoding="utf-8")
        except Exception:
            pass
    _MSG = ("오류: 채점에 필요한 openpyxl 라이브러리가 설치되어 있지 않습니다.\n"
            "설치 방법: 명령 프롬프트에서 다음 명령을 실행하세요.\n"
            "    py -m pip install openpyxl\n"
            "  (위 명령이 안 되면: python -m pip install openpyxl)")
    print(_MSG)
    try:
        print(_MSG, file=sys.stderr)
    except Exception:
        pass
    sys.exit(3)

DEFAULT_POINTS = {
    "기본작업-1": 5,
    "기본작업-2": 10,
    "기본작업-3": 5,
    "계산작업": 40,
    "분석작업-1": 10,
    "분석작업-2": 10,
    "매크로작업": 10,
    "차트작업": 10,
}
PASS_LINE = 70
CALC_GROUP_TARGET = 5
MAX_SCAN_ROWS = 400
MAX_SCAN_COLS = 100

NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"c": NS_C, "a": NS_A}

FMT_KIND_LABEL = {
    "value": "셀 값",
    "number_format": "표시 형식",
    "font": "글꼴 서식",
    "fill": "채우기",
    "alignment": "맞춤(정렬)",
    "border": "테두리",
    "merge": "셀 병합",
    "rowheight": "행 높이",
    "colwidth": "열 너비",
    "names": "정의된 이름",
    "style": "셀 스타일",
}

class Unknown:

    __slots__ = ("why",)

    def __init__(self, why="원인을 특정하지 못했습니다"):
        self.why = str(why)

    def __bool__(self):
        return False

    def __len__(self):
        return 0

    def __iter__(self):
        return iter(())

    def __contains__(self, _k):
        return False

    def __getitem__(self, _k):
        raise KeyError(_k)

    def get(self, _k, default=None):
        return default

    def items(self):
        return ()

    def keys(self):
        return ()

    def values(self):
        return ()

    def setdefault(self, _k, default=None):
        return default

    def __eq__(self, other):
        return other is self

    def __ne__(self, other):
        return other is not self

    def __hash__(self):
        return id(self)

    def __repr__(self):
        return f"UNKNOWN({self.why})"

    def __str__(self):
        return "(판정 불가)"

def is_unknown(v):
    return isinstance(v, Unknown)

def unknown_why(*vals, default="원인을 특정하지 못했습니다"):
    for v in vals:
        if isinstance(v, Unknown):
            return v.why
    return default

def any_unknown(*vals):
    return any(isinstance(v, Unknown) for v in vals)

def norm_sheet_name(name):
    n = re.sub(r"\s+", "", str(name))
    for suf in ("-정답", "(정답)", "_정답", "정답"):
        if n.endswith(suf) and len(n) > len(suf):
            n = n[: -len(suf)]
            break
    return n

def formula_text(raw):
    if isinstance(raw, str) and raw.startswith("="):
        return raw
    if isinstance(raw, ArrayFormula):
        return raw.text if raw.text else "=__ARRAY__"
    if isinstance(raw, DataTableFormula):
        return "=__DATATABLE__"
    return None

_CMP_MIRROR = {">": "<", "<": ">", ">=": "<=", "<=": ">=", "=": "=", "<>": "<>"}
_CMP_TOKENS = (">=", "<=", "<>", ">", "<", "=")

def _split_top(s, sep=","):
    parts, cur, depth = [], "", 0
    for ch in s:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == sep and depth == 0:
            parts.append(cur)
            cur = ""
        else:
            cur += ch
    parts.append(cur)
    return parts

def _canon_cmp_one(s):
    hits = []
    depth = 0
    i = 0
    while i < len(s):
        ch = s[i]
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif depth == 0:
            for op in _CMP_TOKENS:
                if s.startswith(op, i):
                    hits.append((i, op))
                    i += len(op) - 1
                    break
        i += 1
    if len(hits) != 1:
        return s
    pos, op = hits[0]
    lhs, rhs = s[:pos], s[pos + len(op):]
    if not lhs or not rhs:
        return s
    lhs, op, rhs = min((lhs, op, rhs), (rhs, _CMP_MIRROR[op], lhs))
    return lhs + op + rhs

def _canon_cmp_seg(seg):
    out = []
    i, n = 0, len(seg)
    while i < n:
        if seg[i] == "(":
            depth, j = 0, i
            while j < n:
                if seg[j] == "(":
                    depth += 1
                elif seg[j] == ")":
                    depth -= 1
                    if depth == 0:
                        break
                j += 1
            inner = seg[i + 1:j] if j < n else seg[i + 1:]
            out.append("(" + ",".join(_canon_cmp_seg(a)
                                      for a in _split_top(inner)) + ")")
            i = j + 1 if j < n else n
            continue
        out.append(seg[i])
        i += 1
    return _canon_cmp_one("".join(out))

def canon_compare(f):
    if not f or not any(op in f for op in ("<", ">", "=")):
        return f
    lits = []

    def _mask(m):
        lits.append(m.group(0))
        return "\x01%d\x01" % (len(lits) - 1)

    try:
        masked = re.sub(r'"(?:[^"]|"")*"', _mask, f)
        done = _canon_cmp_seg(masked)
        return re.sub(r"\x01(\d+)\x01",
                      lambda m: lits[int(m.group(1))], done)
    except Exception:
        return f

def norm_formula(f, swap_cmp=True):
    if f is None:
        return None
    if not isinstance(f, str):
        f = str(f)
    f = f.lstrip("=")
    parts = re.split(r'("(?:[^"]|"")*")', f)
    out = []
    for i, p in enumerate(parts):
        if i % 2 == 1:
            out.append(p)
        else:
            p = p.upper()
            p = re.sub(r"\s+", "", p)
            p = p.replace("$", "").replace("'", "")
            p = p.replace("_XLFN.", "")
            p = p.replace("정답!", "!")
            p = re.sub(r"\[\d+\]", "", p)
            p = p.replace("RANK.EQ(", "RANK(")
            p = re.sub(r"\bFALSE\b", "0", p)
            p = re.sub(r"\bTRUE\b", "1", p)
            out.append(p)
    joined = "".join(out)
    return canon_compare(joined) if swap_cmp else joined

def cmp_order_swapped(a, b):
    try:
        return (norm_formula(a) == norm_formula(b)
                and norm_formula(a, swap_cmp=False)
                != norm_formula(b, swap_cmp=False))
    except Exception:
        return False

_REF_TOKEN = re.compile(
    r"(?<![A-Z0-9_.])"
    r"(?:(?P<sheet>[A-Z0-9_가-힣 .\-]+)!)?"
    r"(?P<a>[A-Z]{1,3}[0-9]{1,7})"
    r"(?::(?P<b>[A-Z]{1,3}[0-9]{1,7}))?"
    r"(?![A-Z0-9_(])")

REF_EXPAND_LIMIT = 4000

def formula_ref_counts(f, limit=REF_EXPAND_LIMIT):
    if not f:
        return None
    try:
        norm = norm_formula(f, swap_cmp=False)
        norm = re.sub(r'"(?:[^"]|"")*"', '""', norm)
        from openpyxl.utils import range_boundaries, get_column_letter
        out = {}

        def bump(k):
            out[k] = out.get(k, 0) + 1

        for m in _REF_TOKEN.finditer(norm):
            sheet = (m.group("sheet") or "").strip()
            a, b = m.group("a"), m.group("b")
            pre = (sheet + "!") if sheet else ""
            if b is None:
                bump(pre + a)
                continue
            try:
                c1, r1, c2, r2 = range_boundaries(f"{a}:{b}")
                if None in (c1, r1, c2, r2):
                    raise ValueError
                if (r2 - r1 + 1) * (c2 - c1 + 1) > limit:
                    bump(pre + f"{a}:{b}")
                    continue
                for rr in range(r1, r2 + 1):
                    for cc in range(c1, c2 + 1):
                        bump(pre + f"{get_column_letter(cc)}{rr}")
            except Exception:
                bump(pre + f"{a}:{b}")
        return out
    except Exception:
        return None

def same_ref_usage(a, b):
    ra, rb = formula_ref_counts(a), formula_ref_counts(b)
    if ra is None or rb is None:
        return False
    return ra == rb

def ref_usage_diff(expected_f, student_f, limit=6):
    ra, rb = formula_ref_counts(expected_f), formula_ref_counts(student_f)
    if ra is None or rb is None:
        return [], []
    more_a, more_b = [], []
    for k in sorted(set(ra) | set(rb)):
        da, db = ra.get(k, 0), rb.get(k, 0)
        if da > db:
            more_a.append(k if da - db == 1 else f"{k}({da - db}회 더)")
        elif db > da:
            more_b.append(k if db - da == 1 else f"{k}({db - da}회 더)")
    return more_a[:limit], more_b[:limit]

RECALC_ENABLED = True
RECALC_TIMEOUT = 180
_RECALC_UNSET = object()
_recalc_cache = {}

def soffice_exe():
    import shutil
    for name in ("soffice", "libreoffice", "soffice.exe"):
        p = shutil.which(name)
        if p:
            return p
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              "/usr/bin/soffice", "/usr/bin/libreoffice",
              "/usr/lib/libreoffice/program/soffice",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        try:
            if os.path.isfile(p):
                return p
        except Exception:
            pass
    return None

def recalc_workbook(path):
    if not RECALC_ENABLED:
        return None
    try:
        st = os.stat(path)
        ck = (os.path.abspath(path), int(st.st_mtime), st.st_size)
    except Exception:
        return None
    if ck in _recalc_cache:
        return _recalc_cache[ck]
    _recalc_cache[ck] = None
    exe = soffice_exe()
    if not exe:
        return None
    import shutil
    import subprocess
    import tempfile
    tmp = None
    try:
        tmp = tempfile.mkdtemp(prefix="grade_recalc_")
        work = os.path.join(tmp, "in" + os.path.splitext(path)[1].lower())
        shutil.copyfile(path, work)
        outdir = os.path.join(tmp, "out")
        os.makedirs(outdir, exist_ok=True)
        env = dict(os.environ)
        env.pop("PYTHONPATH", None)
        cp = subprocess.run(
            [exe, "--headless", "--norestore", "--invisible", "--nolockcheck",
             "-env:UserInstallation=file://" + os.path.join(tmp, "profile"),
             "--convert-to", "xlsx", "--outdir", outdir, work],
            capture_output=True, timeout=RECALC_TIMEOUT, env=env,
            cwd=tmp)
        del cp
        out = os.path.join(outdir, "in.xlsx")
        if not os.path.isfile(out):
            cands = [os.path.join(outdir, n) for n in os.listdir(outdir)
                     if n.lower().endswith(".xlsx")]
            if not cands:
                return None
            out = cands[0]
        wb = openpyxl.load_workbook(out, data_only=True)
        data = {}
        for ws in wb.worksheets:
            cells = {}
            for row in ws.iter_rows(
                    max_row=min(ws.max_row or 1, MAX_SCAN_ROWS),
                    max_col=min(ws.max_column or 1, MAX_SCAN_COLS)):
                for cell in row:
                    if cell.value is not None:
                        cells[(cell.row, cell.column)] = cell.value
            data[ws.title] = cells
        _recalc_cache[ck] = data
        return data
    except Exception:
        return None
    finally:
        if tmp:
            try:
                import shutil as _sh
                _sh.rmtree(tmp, ignore_errors=True)
            except Exception:
                pass

def _as_number(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip().replace(",", "")
        if s.endswith("%"):
            try:
                return float(s[:-1]) / 100.0
            except ValueError:
                return None
        try:
            return float(s)
        except ValueError:
            return None
    return None

def _as_dt(v):
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    if isinstance(v, str):
        m = re.match(
            r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})"
            r"(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$", v.strip())
        if m:
            try:
                return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)),
                                int(m.group(4) or 0), int(m.group(5) or 0),
                                int(m.group(6) or 0))
            except ValueError:
                return None
    return None

def value_eq(a, b):
    a_empty = a is None or (isinstance(a, str) and not a.strip())
    b_empty = b is None or (isinstance(b, str) and not b.strip())
    if a_empty or b_empty:
        return a_empty and b_empty
    if isinstance(a, bool) or isinstance(b, bool):
        return a is b or a == b
    na, nb = _as_number(a), _as_number(b)
    if na is not None and nb is not None:
        if math.isclose(na, nb, rel_tol=1e-6, abs_tol=1e-9):
            return True
        return round(na, 6) == round(nb, 6)
    da, db = _as_dt(a), _as_dt(b)
    if da is not None and db is not None:
        return da == db
    if isinstance(a, time) and isinstance(b, time):
        return a == b
    if isinstance(a, str) and isinstance(b, str):
        return a.strip() == b.strip()
    return a == b

_NF_DIGITS = re.compile(r"[0#?]")

def display_precision(nf):
    if not nf or not isinstance(nf, str):
        return None
    s = nf.split(";")[0]
    s = re.sub(r'"[^"]*"', "", s)
    s = re.sub(r"\[[^\]]*\]", "", s)
    s = re.sub(r"\\.", "", s)
    s = re.sub(r"_.", "", s)
    s = re.sub(r"\*.", "", s)
    if s.strip().lower() == "general" or "@" in s or not _NF_DIGITS.search(s):
        return None
    if re.search(r"[ymdhsaAeE]", s):
        return None
    scale = 100 if "%" in s else 1
    if "." in s:
        dec = len(_NF_DIGITS.findall(s.split(".", 1)[1]))
    else:
        dec = 0
    return scale, dec

def value_eq_display(a, b, nf):
    prec = display_precision(nf)
    if prec is None or isinstance(a, bool) or isinstance(b, bool):
        return False
    na, nb = _as_number(a), _as_number(b)
    if na is None or nb is None:
        return False
    scale, dec = prec
    from decimal import Decimal, ROUND_HALF_UP
    q = Decimal(1).scaleb(-dec)
    try:
        return (Decimal(str(na * scale)).quantize(q, rounding=ROUND_HALF_UP)
                == Decimal(str(nb * scale)).quantize(q, rounding=ROUND_HALF_UP))
    except Exception:
        return False

def fmt_value(v):
    if v is None:
        return "(비어 있음)"
    if isinstance(v, float):
        r = round(v, 6)
        if r == int(r):
            return str(int(r))
        return str(r)
    if isinstance(v, datetime):
        if v.hour == v.minute == v.second == 0:
            return v.strftime("%Y-%m-%d")
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, (ArrayFormula, DataTableFormula)):
        return "(배열/데이터표 수식)"
    s = str(v)
    if len(s) > 60:
        s = s[:57] + "..."
    return s

def disp_width(s):
    w = 0
    for ch in s:
        w += 2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1
    return w

def pad(s, width, align="left"):
    gap = width - disp_width(s)
    if gap <= 0:
        return s
    if align == "right":
        return " " * gap + s
    if align == "center":
        l = gap // 2
        return " " * l + s + " " * (gap - l)
    return s + " " * gap

class Book:

    def __init__(self, path, label):
        self.path = path
        self.label = label
        if not os.path.isfile(path):
            raise FileNotFoundError(f"{label} 파일을 찾을 수 없습니다: {path}")
        try:
            self.raw = openpyxl.load_workbook(path, data_only=False)
            self.cached = openpyxl.load_workbook(path, data_only=True)
        except Exception as e:
            raise RuntimeError(f"{label} 파일을 열 수 없습니다 ({path}): {e}")
        self.zf = zipfile.ZipFile(path)
        self.sheet_part = self._map_sheet_parts()
        self.norm_map = {}
        for s in self.raw.sheetnames:
            self.norm_map.setdefault(norm_sheet_name(s), s)
        self._recalc = _RECALC_UNSET

    def recalc_map(self):
        if self._recalc is _RECALC_UNSET:
            self._recalc = recalc_workbook(self.path)
        return self._recalc

    def recalc_available(self):
        return bool(self.recalc_map())

    def recalc_value(self, sheet_name, r, c):
        m = self.recalc_map()
        if not m:
            return "__MISSING__"
        sh = m.get(sheet_name)
        if sh is None:
            key = norm_sheet_name(sheet_name)
            for t, d in m.items():
                if norm_sheet_name(t) == key:
                    sh = d
                    break
        if sh is None:
            return "__MISSING__"
        return sh.get((r, c), "")

    def drawing_shapes(self, sheet_name):
        out = []
        targets = self.sheet_rel_targets(sheet_name)
        if is_unknown(targets):
            return targets
        for t in targets:
            if not re.match(r"xl/drawings/drawing\d+\.xml$", t):
                continue
            try:
                x = self.zf.read(t).decode("utf-8", "replace")
            except Exception as e:
                return Unknown(f"{self.label} 파일의 그림 파트 {t} 를 읽지 "
                               f"못했습니다({e})")
            for blob in re.finditer(
                    r"<xdr:(?:twoCell|oneCell|absolute)Anchor\b.*?"
                    r"</xdr:(?:twoCell|oneCell|absolute)Anchor>", x, re.S):
                b = blob.group(0)
                if "<xdr:sp" not in b:
                    continue
                a = re.search(r"<xdr:from>.*?<xdr:col>(\d+)</xdr:col>.*?"
                              r"<xdr:row>(\d+)</xdr:row>.*?</xdr:from>.*?"
                              r"<xdr:to>.*?<xdr:col>(\d+)</xdr:col>.*?"
                              r"<xdr:row>(\d+)</xdr:row>.*?</xdr:to>", b, re.S)
                rng = "?"
                if a:
                    c1, r1, c2, r2 = (int(v) for v in a.groups())
                    rng = _anchor_range(c1, r1, max(c1, c2 - 1),
                                        max(r1, r2 - 1))
                mac = re.search(r'<xdr:sp[^>]*\bmacro="([^"]*)"', b)
                nm = re.search(r'<xdr:cNvPr[^>]*\bname="([^"]*)"', b)
                geo = re.search(r'<a:prstGeom[^>]*\bprst="([^"]*)"', b)
                hidden = bool(re.search(r'<xdr:cNvPr[^>]*\bhidden="(?:1|true)"',
                                        b))
                txt = "".join(html_mod.unescape(t2) for t2 in
                              re.findall(r"<a:t>([^<]*)</a:t>", b)).strip()
                macro = (mac.group(1) if mac else "").split("!")[-1].strip()
                out.append({
                    "text": txt, "anchor": rng, "macro": macro,
                    "geom": geo.group(1) if geo else "",
                    "name": html_mod.unescape(nm.group(1)) if nm else "",
                    "hidden": hidden})
        return out

    def _map_sheet_parts(self):
        mapping = {}
        self.parts_error = None
        try:
            wbxml = self.zf.read("xl/workbook.xml").decode("utf-8", "replace")
            relsxml = self.zf.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
            rels = {}
            for tag in re.findall(r"<Relationship\b[^>]*>", relsxml):
                rid = re.search(r'Id="([^"]+)"', tag)
                tgt = re.search(r'Target="([^"]+)"', tag)
                if rid and tgt:
                    rels[rid.group(1)] = tgt.group(1)
            for m in re.finditer(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', wbxml):
                name = html_mod.unescape(m.group(1))
                target = rels.get(m.group(2), "")
                if not target:
                    continue
                if target.startswith("/"):
                    part = target.lstrip("/")
                else:
                    part = "xl/" + target
                part = part.replace("xl/xl/", "xl/")
                mapping[name] = part
        except Exception as e:
            self.parts_error = f"xl/workbook.xml 의 시트 목록을 읽지 못했습니다({e})"
            return mapping
        if not mapping and self.raw.sheetnames:
            self.parts_error = ("xl/workbook.xml 에서 시트 ↔ 파트 연결을 "
                                "하나도 찾지 못했습니다")
        return mapping

    def _part_miss(self, sheet_name, what):
        if getattr(self, "parts_error", None):
            return Unknown(f"{self.label} 파일의 {what}: {self.parts_error}")
        return None

    def sheet_xml(self, sheet_name):
        part = self.sheet_part.get(sheet_name)
        if not part:
            return self._part_miss(sheet_name, f"'{sheet_name}' 시트 XML")
        try:
            return self.zf.read(part).decode("utf-8", "replace")
        except KeyError:
            return Unknown(f"{self.label} 파일에 시트 파트 {part} 가 없습니다")
        except Exception as e:
            return Unknown(f"{self.label} 파일의 시트 파트 {part} 를 "
                           f"읽지 못했습니다({e})")

    def sheet_rels(self, sheet_name):
        part = self.sheet_part.get(sheet_name)
        if not part:
            return self._part_miss(sheet_name, f"'{sheet_name}' 시트 관계") or {}
        rels_part = re.sub(r"worksheets/(sheet\d+\.xml)$",
                           r"worksheets/_rels/\1.rels", part)
        try:
            x = self.zf.read(rels_part).decode("utf-8", "replace")
        except KeyError:
            return {}
        except Exception as e:
            return Unknown(f"{self.label} 파일의 {rels_part} 를 읽지 "
                           f"못했습니다({e})")
        out = {}
        for m in re.finditer(r'<Relationship\b[^>]*>', x):
            tag = m.group(0)
            rid = re.search(r'Id="([^"]+)"', tag)
            tgt = re.search(r'Target="([^"]+)"', tag)
            if not (rid and tgt):
                continue
            t = tgt.group(1)
            out[rid.group(1)] = t.lstrip("/") if t.startswith("/") else \
                os.path.normpath("xl/worksheets/" + t).replace("\\", "/")
        return out

    def sheet_rel_targets(self, sheet_name):
        part = self.sheet_part.get(sheet_name)
        if not part:
            return self._part_miss(sheet_name,
                                   f"'{sheet_name}' 시트 관계 목록") or []
        rels_part = re.sub(r"worksheets/(sheet\d+\.xml)$", r"worksheets/_rels/\1.rels", part)
        try:
            x = self.zf.read(rels_part).decode("utf-8", "replace")
        except KeyError:
            return []
        except Exception as e:
            return Unknown(f"{self.label} 파일의 {rels_part} 를 읽지 "
                           f"못했습니다({e})")
        out = []
        for m in re.finditer(r'Target="([^"]+)"', x):
            t = m.group(1)
            if t.startswith("/"):
                out.append(t.lstrip("/"))
            else:
                out.append(os.path.normpath("xl/worksheets/" + t).replace("\\", "/"))
        return out

    def has_vba(self):
        return any("vbaProject" in n for n in self.zf.namelist())

    def vba_modules(self):
        if getattr(self, "_vba_mods", None) is None:
            mods = []
            for n in self.zf.namelist():
                if n.endswith("vbaProject.bin"):
                    try:
                        mods = vba_module_sources(self.zf.read(n))
                    except Exception as e:
                        mods = Unknown(f"{self.label} 파일의 {n} 을 여는 중 "
                                       f"오류: {e}")
                    break
            self._vba_mods = mods
        return self._vba_mods

    def vba_diag(self):
        for n in self.zf.namelist():
            if n.endswith("vbaProject.bin"):
                try:
                    return vba_read_diag(self.zf.read(n))
                except Exception as e:
                    return f"vbaProject.bin을 여는 중 오류: {e}"
        return None

    def vba_units(self):
        if getattr(self, "_vba_units", None) is None:
            mods = self.vba_modules()
            if is_unknown(mods):
                self._vba_units = mods
                return self._vba_units
            units = []
            for _name, src in mods:
                units.extend(vba_macro_units(src))
            if not units and self.has_vba():
                why = self.vba_diag()
                if why:
                    units = Unknown(f"{self.label} 파일의 매크로 코드: {why}")
            self._vba_units = units
        return self._vba_units

    def form_controls(self, sheet_name):
        out = []
        targets = self.sheet_rel_targets(sheet_name)
        if is_unknown(targets):
            return targets
        for t in targets:
            if "vmlDrawing" not in t:
                continue
            try:
                x = self.zf.read(t).decode("utf-8", "replace")
            except Exception as e:
                return Unknown(f"{self.label} 파일의 VML 파트 {t} 를 읽지 "
                               f"못했습니다({e})")
            for sh in re.findall(r"<v:shape\b.*?</v:shape>", x, re.S):
                if "ObjectType=\"Button\"" not in sh and "o:button" not in sh:
                    continue
                mac = re.search(r"<x:FmlaMacro>\s*([^<]*)</x:FmlaMacro>", sh)
                anc = re.search(r"<x:Anchor>\s*([^<]*)</x:Anchor>", sh)
                txt = _vml_textbox_text(sh)
                rng = "?"
                if anc:
                    nums = [n.strip() for n in anc.group(1).split(",")]
                    if len(nums) >= 8:
                        rng = _anchor_range(nums[0], nums[2], nums[4], nums[6])
                name = (mac.group(1).strip() if mac else "")
                out.append({"text": txt, "anchor": rng,
                            "macro": name.split("!")[-1] if name else ""})
        rels = self.sheet_rels(sheet_name)
        xml = self.sheet_xml(sheet_name)
        if is_unknown(xml) or is_unknown(rels):
            return Unknown(unknown_why(xml, rels))
        x = xml or ""
        for blob in re.findall(r"<control\b.*?(?:</control>|/>)", x, re.S):
            a = re.search(r"<(?:\w+:)?from>.*?<(?:\w+:)?col>(\d+)</(?:\w+:)?col>"
                          r".*?<(?:\w+:)?row>(\d+)</(?:\w+:)?row>.*?"
                          r"</(?:\w+:)?from>.*?<(?:\w+:)?to>.*?"
                          r"<(?:\w+:)?col>(\d+)</(?:\w+:)?col>.*?"
                          r"<(?:\w+:)?row>(\d+)</(?:\w+:)?row>.*?"
                          r"</(?:\w+:)?to>", blob, re.S)
            rng = _anchor_range(*a.groups()) if a else "?"
            nm = re.search(r'<control\b[^>]*\bname="([^"]*)"', blob)
            label = html_mod.unescape(nm.group(1)) if nm else ""
            mpr = re.search(r'<controlPr\b[^>]*\bmacro="([^"]*)"', blob)
            rid = re.search(r'r:id="(rId\d+)"', blob)
            text, macro = label, ""
            if mpr:
                macro = html_mod.unescape(mpr.group(1)).split("!")[-1].strip()
            if any(o["anchor"] == rng for o in out if o["anchor"] != "?"):
                continue
            if macro and any(o.get("macro") == macro for o in out):
                continue
            if rid and rels.get(rid.group(1)):
                try:
                    cx = self.zf.read(rels[rid.group(1)]).decode(
                        "utf-8", "replace")
                    mt = re.search(r"<text>([^<]*)</text>", cx)
                    if mt:
                        text = html_mod.unescape(mt.group(1)).strip()
                    mm = re.search(r'fmlaMacro="([^"]*)"', cx)
                    if mm:
                        macro = html_mod.unescape(mm.group(1)).split("!")[-1]
                except Exception as e:
                    return Unknown(
                        f"{self.label} 파일의 컨트롤 속성 파트 "
                        f"{rels[rid.group(1)]} 를 읽지 못했습니다({e})")
            if rng == "?" and not macro:
                return Unknown(f"{self.label} 파일의 '{sheet_name}' 시트에 "
                               "앵커·연결 매크로를 읽지 못한 컨트롤이 "
                               "있습니다")
            out.append({"text": text, "anchor": rng, "macro": macro})
        return out

    def has_pivot_for_sheet(self, sheet_name):
        return any("pivotTable" in t for t in self.sheet_rel_targets(sheet_name))

    def has_scenarios(self, sheet_name):
        x = self.sheet_xml(sheet_name)
        return bool(x) and "<scenarios" in x

    def sheet_has_drawing(self, sheet_name):
        targets = self.sheet_rel_targets(sheet_name)
        if is_unknown(targets):
            return targets
        for t in targets:
            if "/drawings/" in t or "vmlDrawing" in t or "ctrlProp" in t:
                return True
        x = self.sheet_xml(sheet_name)
        if is_unknown(x):
            return x
        x = x or ""
        return "<controls" in x or "<legacyDrawing" in x

    def chart_parts_for_sheet(self, sheet_name):
        charts = []
        targets = self.sheet_rel_targets(sheet_name)
        if is_unknown(targets):
            return targets
        for t in targets:
            m = re.match(r"xl/drawings/(drawing\d+)\.xml$", t)
            if not m:
                continue
            drels = f"xl/drawings/_rels/{m.group(1)}.xml.rels"
            try:
                x = self.zf.read(drels).decode("utf-8", "replace")
            except KeyError:
                continue
            except Exception as e:
                return Unknown(f"{self.label} 파일의 {drels} 를 읽지 "
                               f"못했습니다({e})")
            for mm in re.finditer(r'Target="([^"]+charts/chart\d+\.xml)"', x):
                p = mm.group(1)
                p = p.lstrip("/") if p.startswith("/") else \
                    os.path.normpath("xl/drawings/" + p).replace("\\", "/")
                charts.append(p)
        if not charts:
            charts = sorted(
                n for n in self.zf.namelist() if re.match(r"xl/charts/chart\d+\.xml$", n)
            )
        return charts

_CFB_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_CFB_FREESECT = 0xFFFFFFFF
_CFB_ENDOFCHAIN = 0xFFFFFFFE

def cfb_streams(data):
    import struct
    if len(data) < 512 or data[:8] != _CFB_MAGIC:
        return Unknown("OLE(CFB) 형식이 아닙니다")
    try:
        ssz = 1 << struct.unpack_from("<H", data, 30)[0]
        msz = 1 << struct.unpack_from("<H", data, 32)[0]
        n_fat = struct.unpack_from("<I", data, 44)[0]
        dir_first = struct.unpack_from("<I", data, 48)[0]
        cutoff = struct.unpack_from("<I", data, 56)[0] or 4096
        mini_first = struct.unpack_from("<I", data, 60)[0]
        difat_first = struct.unpack_from("<I", data, 68)[0]
        n_difat = struct.unpack_from("<I", data, 72)[0]
        if ssz < 64 or msz < 8:
            return Unknown(f"CFB 섹터 크기가 비정상입니다(섹터 {ssz}, "
                           f"미니 섹터 {msz})")

        def sector(i):
            off = (i + 1) * ssz
            return data[off:off + ssz]

        difat = list(struct.unpack_from("<109I", data, 76))
        nxt, guard = difat_first, 0
        while nxt not in (_CFB_ENDOFCHAIN, _CFB_FREESECT) \
                and guard <= n_difat + 8:
            blk = sector(nxt)
            if len(blk) < ssz:
                break
            vals = struct.unpack_from("<%dI" % (ssz // 4), blk, 0)
            difat.extend(vals[:-1])
            nxt = vals[-1]
            guard += 1
        fat = []
        for s in difat[:n_fat or len(difat)]:
            if s in (_CFB_ENDOFCHAIN, _CFB_FREESECT):
                continue
            blk = sector(s)
            if len(blk) < ssz:
                break
            fat.extend(struct.unpack_from("<%dI" % (ssz // 4), blk, 0))

        def read_chain(start, size=None):
            out, cur, seen = [], start, 0
            while cur not in (_CFB_ENDOFCHAIN, _CFB_FREESECT) \
                    and cur < len(fat) and seen < len(fat) + 8:
                out.append(sector(cur))
                cur = fat[cur]
                seen += 1
            buf = b"".join(out)
            return buf if size is None else buf[:size]

        minifat = []
        if mini_first not in (_CFB_ENDOFCHAIN, _CFB_FREESECT):
            mb = read_chain(mini_first)
            minifat = list(struct.unpack_from("<%dI" % (len(mb) // 4), mb, 0))
        dir_bytes = read_chain(dir_first)
        entries = []
        for i in range(0, len(dir_bytes) - 127, 128):
            e = dir_bytes[i:i + 128]
            nlen = struct.unpack_from("<H", e, 64)[0]
            name = e[:max(0, min(64, nlen - 2))].decode("utf-16-le", "replace")
            entries.append((name, e[66], struct.unpack_from("<I", e, 116)[0],
                            struct.unpack_from("<Q", e, 120)[0]))
        mini_stream = b""
        for (_nm, etype, start, size) in entries:
            if etype == 5:
                mini_stream = read_chain(start,
                                         size if size < (1 << 31) else None)
                break

        def read_mini(start, size):
            out, cur = bytearray(), start
            while cur not in (_CFB_ENDOFCHAIN, _CFB_FREESECT) \
                    and len(out) < size:
                out += mini_stream[cur * msz:cur * msz + msz]
                if cur >= len(minifat):
                    break
                cur = minifat[cur]
            return bytes(out[:size])

        streams = {}
        for (name, etype, start, size) in entries:
            if etype != 2 or not size or size > (1 << 26):
                continue
            try:
                streams[name] = read_mini(start, size) if size < cutoff \
                    else read_chain(start, size)
            except Exception:
                continue
        if not streams:
            return Unknown("CFB 구조는 열렸지만 스트림을 하나도 읽지 "
                           "못했습니다")
        return streams
    except Exception as e:
        return Unknown(f"CFB 구조를 해석하지 못했습니다({type(e).__name__})")

def ovba_decompress(data, start):
    if start >= len(data) or data[start] != 0x01:
        return None
    out = bytearray()
    pos, n = start + 1, len(data)
    chunks = 0
    try:
        while pos + 1 < n:
            hdr = data[pos] | (data[pos + 1] << 8)
            pos += 2
            if (hdr >> 12) & 0x07 != 0b011:
                return bytes(out) if chunks else None
            chunks += 1
            end = min(pos + (hdr & 0x0FFF) + 1, n)
            chunk_start = len(out)
            if not (hdr >> 15) & 1:
                out += data[pos:end]
                pos = end
                continue
            p = pos
            while p < end:
                flags = data[p]
                p += 1
                for bit in range(8):
                    if p >= end:
                        break
                    if not (flags >> bit) & 1:
                        out.append(data[p])
                        p += 1
                        continue
                    if p + 1 >= end:
                        p = end
                        break
                    tok = data[p] | (data[p + 1] << 8)
                    p += 2
                    diff = len(out) - chunk_start
                    bc = max(4, (diff - 1).bit_length()) if diff else 4
                    lbc = 16 - bc
                    length = (tok & ((1 << lbc) - 1)) + 3
                    src = len(out) - (((tok >> lbc) & ((1 << bc) - 1)) + 1)
                    if src < 0:
                        return None
                    for _ in range(length):
                        out.append(out[src])
                        src += 1
            pos = end
            if len(out) > 4 * 1024 * 1024:
                break
        return bytes(out)
    except Exception:
        return None

def _vml_textbox_text(shape_xml):
    box = re.search(r"<v:textbox\b.*?</v:textbox>", shape_xml, re.S)
    inner = box.group(0) if box else shape_xml
    text = re.sub(r"<[^>]*>", " ", inner)
    return re.sub(r"\s+", " ", html_mod.unescape(text)).strip()

_ATTEMPT_STAMP_RE = re.compile(r"_(\d{8})_(\d{4})(?:\D|$)")

def attempt_report_id(student_path):
    m = _ATTEMPT_STAMP_RE.search(os.path.basename(str(student_path or "")))
    return f"{m.group(1)}{m.group(2)}00" if m else None

def vba_read_diag(bin_bytes):
    if not bin_bytes:
        return "vbaProject.bin이 비어 있습니다"
    if len(bin_bytes) < 512 or bin_bytes[:8] != _CFB_MAGIC:
        return "vbaProject.bin이 OLE(CFB) 형식이 아닙니다"
    streams = cfb_streams(bin_bytes)
    if not streams:
        return "vbaProject.bin의 OLE(CFB) 구조를 해석하지 못했습니다"
    cand = [n for n in streams
            if n not in ("dir", "PROJECT", "PROJECTwm", "_VBA_PROJECT")
            and not n.startswith("__SRP_")]
    if not cand:
        return ("vbaProject.bin 안에서 모듈 스트림을 찾지 못했습니다 "
                f"(읽은 스트림: {', '.join(sorted(streams))[:80]})")
    if not vba_module_sources(bin_bytes):
        return (f"모듈 스트림({', '.join(cand[:3])})의 MS-OVBA 압축을 "
                "풀지 못했습니다")
    return None

def vba_module_sources(bin_bytes):
    mods = []
    streams = cfb_streams(bin_bytes)
    if is_unknown(streams):
        return Unknown(f"vbaProject.bin: {streams.why}")
    cand = 0
    for name, body in streams.items():
        if name in ("dir", "PROJECT", "PROJECTwm", "_VBA_PROJECT") \
                or name.startswith("__SRP_"):
            continue
        cand += 1
        for i, b in enumerate(body):
            if b != 0x01:
                continue
            raw = ovba_decompress(body, i)
            if not raw or b"Attribute VB_Name" not in raw:
                continue
            txt = None
            for enc in ("cp949", "utf-8", "latin-1"):
                try:
                    txt = raw.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
            if txt:
                mods.append((name, txt))
            break
    if not mods and cand:
        return Unknown("vbaProject.bin 의 모듈 스트림에서 MS-OVBA 압축을 "
                       "풀지 못했습니다")
    return mods

VBA_ACTION_KINDS = (
    (r"\.Interior\b|\.ColorIndex\b(?=[^\n]*Interior)", "fill"),
    (r"\.Font\b", "font"),
    (r"\.NumberFormat(?:Local)?\b", "number_format"),
    (r"\.(?:Horizontal|Vertical)Alignment\b|\.WrapText\b|\.Orientation\b",
     "alignment"),
    (r"\.Borders?\b|\.BorderAround\b", "border"),
    (r"\.Merge\b|\.MergeCells\b", "merge"),
    (r"\.Formula(?:R1C1|Array)?\b|\.Value2?\b|\.AutoFill\b|\.FillDown\b"
     r"|\.AutoSum\b", "value"),
)
VBA_KIND_LABEL = {"fill": "채우기", "font": "글꼴", "number_format": "표시 형식",
                  "alignment": "맞춤", "border": "테두리", "merge": "병합",
                  "value": "값/수식"}

def vba_macro_units(source):
    units = []
    cur = None
    for line in (source or "").splitlines():
        s = line.strip()
        if s.startswith("'"):
            continue
        m = re.match(r"^(?:Public\s+|Private\s+)?(?:Sub|Function)\s+"
                     r"([^\s(]+)", s, re.I)
        if m:
            cur = {"name": m.group(1).strip(), "ranges": [], "kinds": set()}
            units.append(cur)
            continue
        if re.match(r"^End\s+(?:Sub|Function)\b", s, re.I):
            cur = None
            continue
        if cur is None:
            continue
        for m in re.finditer(r'Range\(\s*"([^"]+)"', s, re.I):
            for part in m.group(1).split(","):
                part = part.replace("$", "").strip().upper()
                if part:
                    cur["ranges"].append(part)
        for m in re.finditer(r"\[\s*([A-Za-z]{1,3}\d+(?::[A-Za-z]{1,3}\d+)?)"
                             r"\s*\]", s):
            cur["ranges"].append(m.group(1).replace("$", "").upper())
        for pat, kind in VBA_ACTION_KINDS:
            if re.search(pat, s, re.I):
                cur["kinds"].add(kind)
    return units

def vba_covers(units, coord, kind):
    from openpyxl.utils import range_boundaries
    from openpyxl.utils.cell import coordinate_to_tuple
    try:
        r, c = coordinate_to_tuple(str(coord).upper())
    except Exception:
        return None
    for u in units:
        if kind not in u["kinds"]:
            continue
        for rng in u["ranges"]:
            try:
                c1, r1, c2, r2 = range_boundaries(
                    rng if ":" in rng else f"{rng}:{rng}")
            except Exception:
                continue
            if None in (c1, r1, c2, r2):
                continue
            if r1 <= r <= r2 and c1 <= c <= c2:
                return u["name"]
    return None

def _anchor_range(col1, row1, col2, row2):
    from openpyxl.utils import get_column_letter
    try:
        a = f"{get_column_letter(int(col1) + 1)}{int(row1) + 1}"
        b = f"{get_column_letter(int(col2) + 1)}{int(row2) + 1}"
        return a if a == b else f"{a}:{b}"
    except Exception:
        return "?"

_THEME_ORDER = ("lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3",
                "accent4", "accent5", "accent6", "hlink", "folHlink")
_DEFAULT_THEME = ("FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31",
                  "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72")
_theme_cache = {}

def _theme_palette(wb):
    key = id(wb)
    if key in _theme_cache:
        return _theme_cache[key]
    palette = list(_DEFAULT_THEME)
    try:
        xml = getattr(wb, "loaded_theme", None)
        if xml:
            if isinstance(xml, bytes):
                xml = xml.decode("utf-8", "replace")
            found = {}
            for name in ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2",
                         "accent3", "accent4", "accent5", "accent6",
                         "hlink", "folHlink"):
                m = re.search(
                    rf"<a:{name}>.*?(?:srgbClr\s+val=\"([0-9A-Fa-f]{{6}})\""
                    rf"|sysClr[^>]*lastClr=\"([0-9A-Fa-f]{{6}})\")",
                    xml, re.DOTALL)
                if m:
                    found[name] = (m.group(1) or m.group(2)).upper()
            for i, nm in enumerate(_THEME_ORDER):
                if nm in found:
                    palette[i] = found[nm]
    except Exception:
        pass
    _theme_cache[key] = palette
    return palette

def _apply_tint(rgb_hex, tint):
    import colorsys
    try:
        r = int(rgb_hex[0:2], 16) / 255.0
        g = int(rgb_hex[2:4], 16) / 255.0
        b = int(rgb_hex[4:6], 16) / 255.0
        h, l, s = colorsys.rgb_to_hls(r, g, b)
        if tint < 0:
            l = l * (1.0 + tint)
        elif tint > 0:
            l = l * (1.0 - tint) + tint
        r2, g2, b2 = colorsys.hls_to_rgb(h, max(0.0, min(1.0, l)), s)
        return "%02X%02X%02X" % (round(r2 * 255), round(g2 * 255),
                                 round(b2 * 255))
    except Exception:
        return rgb_hex

def _quant_hex(rgb_hex):
    try:
        vals = [min(255, (int(rgb_hex[i:i + 2], 16) + 2) // 4 * 4)
                for i in (0, 2, 4)]
        return "%02X%02X%02X" % tuple(vals)
    except Exception:
        return rgb_hex

def _color_key(color, wb=None):
    if color is None:
        return None
    try:
        if color.type == "rgb":
            argb = color.rgb
            if argb in (None, "00000000"):
                return None
            return ("rgb", _quant_hex(str(argb)[-6:].upper()))
        if color.type == "theme":
            idx = color.theme
            tint = float(color.tint or 0.0)
            if wb is not None and isinstance(idx, int) and \
                    0 <= idx < len(_THEME_ORDER):
                base = _theme_palette(wb)[idx]
                return ("rgb", _quant_hex(_apply_tint(base, tint)))
            return ("theme", idx, round(tint, 2))
        if color.type == "indexed":
            if color.indexed in (64, 65):
                return None
            try:
                from openpyxl.styles.colors import COLOR_INDEX
                argb = COLOR_INDEX[color.indexed]
                return ("rgb", _quant_hex(str(argb)[-6:].upper()))
            except Exception:
                return ("idx", color.indexed)
    except Exception:
        pass
    return None

_COLOR_NAMES_KO = (
    ("FFFFFF", "흰색"), ("000000", "검정"), ("FF0000", "빨강"),
    ("FFFF00", "노랑"), ("FFC000", "황금색"), ("BF8F00", "황금색(어둡게)"),
    ("ED7D31", "주황"), ("00B050", "초록"), ("70AD47", "연한 초록(강조6)"),
    ("A9D08E", "연한 초록"), ("4472C4", "파랑(강조1)"), ("0070C0", "파랑"),
    ("5B9BD5", "하늘색(강조5)"), ("A5A5A5", "회색(강조3)"),
    ("D9D9D9", "연한 회색"), ("808080", "회색"), ("7030A0", "보라"),
    ("FFE699", "연한 황금색"), ("DDEBF7", "연한 파랑"),
    ("FCE4D6", "연한 주황"),
)

def color_name_ko(rgb_hex):
    try:
        r = int(rgb_hex[0:2], 16)
        g = int(rgb_hex[2:4], 16)
        b = int(rgb_hex[4:6], 16)
        best = None
        for hx, nm in _COLOR_NAMES_KO:
            d = (abs(r - int(hx[0:2], 16)) + abs(g - int(hx[2:4], 16))
                 + abs(b - int(hx[4:6], 16)))
            if best is None or d < best[0]:
                best = (d, nm)
        if best and best[0] == 0:
            return f"{best[1]}(#{rgb_hex})"
        if best and best[0] <= 90:
            return f"{best[1]}(#{rgb_hex} 계열)"
        return f"#{rgb_hex}"
    except Exception:
        return f"#{rgb_hex}"

FONT_ALIASES = {
    "malgun gothic": "맑은 고딕", "gulim": "굴림", "gulimche": "굴림체",
    "dotum": "돋움", "dotumche": "돋움체", "batang": "바탕",
    "batangche": "바탕체", "gungsuh": "궁서", "gungsuhche": "궁서체",
    "nanumgothic": "나눔고딕", "nanum gothic": "나눔고딕",
}

def _norm_font_name(name):
    if not name:
        return name
    return FONT_ALIASES.get(str(name).strip().lower(), str(name).strip())

_NF_SIGNIFICANT = set('0123456789#?.,%eEgG@ymdhsSbBnN*_/:[]"\\;aA')

_NF_GENERAL_KO = "G/표준"

def _nf_split_quoted(s):
    return re.split(r'("(?:[^"]|"")*")', str(s))

def _nf_general_alias(s):
    parts = _nf_split_quoted(s)
    return "".join(p if i % 2 else p.replace(_NF_GENERAL_KO, "General")
                   for i, p in enumerate(parts))

def nf_display(nf):
    if is_unknown(nf):
        return "확인 불가"
    s = str(nf) if nf not in (None, "") else "General"
    parts = _nf_split_quoted(s)
    return "".join(p if i % 2 else re.sub(r"General", _NF_GENERAL_KO, p,
                                          flags=re.IGNORECASE)
                   for i, p in enumerate(parts))

def _nf_general_str(v):
    f = float(v)
    if f == int(f) and abs(f) < 1e15:
        return str(int(f))
    return ("%.10f" % f).rstrip("0").rstrip(".")

def _nf_sections(s):
    secs, cur, in_q = [], "", False
    for ch in s:
        if ch == '"':
            in_q = not in_q
        if ch == ";" and not in_q:
            secs.append(cur)
            cur = ""
        else:
            cur += ch
    secs.append(cur)
    return secs

def _nf_parse_section(sec):
    pre, post = [], []
    cur = pre
    core = None
    i, n = 0, len(sec)
    while i < n:
        ch = sec[i]
        if ch == '"':
            j = sec.find('"', i + 1)
            if j == -1:
                return None
            cur.append(sec[i + 1:j])
            i = j + 1
            continue
        if ch == "\\" and i + 1 < n:
            cur.append(sec[i + 1])
            i += 2
            continue
        if ch == "_" and i + 1 < n:
            cur.append(" ")
            i += 2
            continue
        if ch == "[":
            j = sec.find("]", i)
            if j == -1:
                return None
            i = j + 1
            continue
        if sec[i:i + 7].lower() == "general":
            if core is not None:
                return None
            core, cur = "General", post
            i += 7
            continue
        if ch in "0#?.,%":
            if core is None:
                core, cur = "", post
            elif core == "General" or post:
                return None
            core += ch
            i += 1
            continue
        if ch in "@*/" or ch in "eE" or ch.lower() in "ymdhsb":
            return None
        cur.append(ch)
        i += 1
    if core is None:
        return None
    return "".join(pre), core, "".join(post)

def _nf_render_core(v, core):
    if core == "General":
        return _nf_general_str(v)
    body = core
    scale = 1.0
    pct_pre = pct_post = ""
    if "%" in body:
        scale *= 100.0 ** body.count("%")
        pct_pre = "%" * (len(body) - len(body.lstrip("%")))
        pct_post = "%" * (len(body) - len(body.rstrip("%")))
        body = body.replace("%", "")
    m = re.search(r"(,+)$", body)
    if m:
        scale /= 1000.0 ** len(m.group(1))
        body = body[:m.start()]
    thousands = "," in body
    body = body.replace(",", "")
    int_pat, _, dec_pat = body.partition(".")
    dec = len(re.findall(r"[0#?]", dec_pat))
    from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
    try:
        d = Decimal(str(v * scale)).quantize(Decimal(1).scaleb(-dec),
                                             rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError, OverflowError):
        return None
    neg = d < 0
    txt = f"{-d if neg else d:f}"
    itxt, _, dtxt = txt.partition(".")
    min_int = int_pat.count("0")
    if len(itxt) < min_int:
        itxt = itxt.rjust(min_int, "0")
    if min_int == 0 and itxt == "0":
        itxt = ""
    if thousands and itxt:
        groups, t = [], itxt
        while len(t) > 3:
            groups.insert(0, t[-3:])
            t = t[:-3]
        groups.insert(0, t)
        itxt = ",".join(groups)
    out = pct_pre + itxt + ("." + dtxt if dec else "") + pct_post
    return ("-" if neg else "") + out

def nf_render(value, nf):
    if value is None or isinstance(value, bool) \
            or isinstance(value, (datetime, date, time)):
        return None
    v = _as_number(value)
    if v is None:
        return None
    secs = _nf_sections(_nf_general_alias(str(nf or "General").strip()))
    used_neg = False
    if v < 0 and len(secs) >= 2 and secs[1].strip():
        sec, used_neg = secs[1], True
    elif v == 0 and len(secs) >= 3 and secs[2].strip():
        sec = secs[2]
    else:
        sec = secs[0]
    parsed = _nf_parse_section(sec)
    if parsed is None:
        return None
    pre, core, post = parsed
    body = _nf_render_core(abs(v) if used_neg else v, core)
    if body is None:
        return None
    return pre + body + post

def nf_same_display(value, nf_a, nf_b):
    if _norm_number_format(nf_a) == _norm_number_format(nf_b):
        return True
    ra = nf_render(value, nf_a)
    return ra is not None and ra == nf_render(value, nf_b)

def _norm_number_format(nf):
    s = (nf or "General").strip()
    s = _nf_general_alias(s)
    if s in ("", "General"):
        return "General"
    s = re.sub(r"\[\$([^\]-]*)-[0-9A-Za-z\-]+\]", r"\1", s)

    def _norm_section(sec):
        out = []
        i = 0
        while i < len(sec):
            ch = sec[i]
            if ch == "\\" and i + 1 < len(sec):
                lit = sec[i + 1]
                out.append(f'"{lit}"' if lit in _NF_SIGNIFICANT else lit)
                i += 2
                continue
            if ch == '"':
                j = sec.find('"', i + 1)
                if j == -1:
                    out.append(ch)
                    i += 1
                    continue
                for lit in sec[i + 1:j]:
                    out.append(f'"{lit}"' if lit in _NF_SIGNIFICANT else lit)
                i = j + 1
                continue
            out.append(ch)
            i += 1
        return "".join(out)

    sections = []
    cur = ""
    in_q = False
    for ch in s:
        if ch == '"':
            in_q = not in_q
        if ch == ";" and not in_q:
            sections.append(cur)
            cur = ""
        else:
            cur += ch
    sections.append(cur)
    sections = [_norm_section(x) for x in sections]
    while len(sections) > 1 and sections[-1].strip() in ("@", ""):
        sections.pop()
    if len(sections) == 2 and sections[1] == "-" + sections[0]:
        sections.pop()
    return ";".join(sections)

def _norm_h_align(h, value):
    if h in (None, "general"):
        return None
    default = None
    if isinstance(value, bool):
        default = "center"
    elif isinstance(value, str) and value != "" \
            and not value.startswith("="):
        default = "left"
    elif isinstance(value, (int, float, datetime, date)):
        default = "right"
    return None if h == default else h

def fmt_signature(cell, kind):
    try:
        wb = None
        try:
            wb = cell.parent.parent
        except Exception:
            pass
        if kind == "number_format":
            return _norm_number_format(cell.number_format)
        if kind == "font":
            f = cell.font
            u = f.underline
            u = None if u in (None, "none") else u
            col = _color_key(f.color, wb)
            if col == ("rgb", "000000"):
                col = None
            return (_norm_font_name(f.name), float(f.size or 11),
                    bool(f.bold), bool(f.italic), u, bool(f.strike), col)
        if kind == "fill":
            fl = cell.fill
            pat = fl.patternType if fl else None
            if pat in (None, "none"):
                return None
            return (pat, _color_key(fl.fgColor, wb))
        if kind == "alignment":
            al = cell.alignment
            v = al.vertical
            v = None if v in (None, "bottom") else v
            return (_norm_h_align(al.horizontal, cell.value), v,
                    bool(al.wrap_text), int(al.text_rotation or 0))
        if kind == "border":
            b = cell.border
            return tuple(bool(s is not None and s.style) for s in
                         (b.left, b.right, b.top, b.bottom, b.diagonal))
    except Exception as e:
        return Unknown(f"셀 서식({FMT_KIND_LABEL.get(kind, kind)})을 읽지 "
                       f"못했습니다({type(e).__name__})")
    return None

CELL_FMT_KINDS = ("number_format", "font", "fill", "alignment")

_BORDER_WEIGHT = {"hair": 1, "thin": 2, "dotted": 2, "dashed": 2,
                  "dashDot": 2, "dashDotDot": 2, "mediumDashed": 3,
                  "mediumDashDot": 3, "mediumDashDotDot": 3,
                  "slantDashDot": 3, "medium": 3, "thick": 4, "double": 5}

BORDER_STYLE_KO = {"thin": "실선", "medium": "굵은 실선", "double": "이중 실선",
                   "dotted": "점선", "dashed": "파선", "hair": "가는 실선",
                   "thick": "매우 굵은 실선", "mediumDashed": "굵은 파선",
                   "dashDot": "일점쇄선", "dashDotDot": "이점쇄선"}

def _combine_edge(s1, s2):
    if s1 and s2:
        return s1 if _BORDER_WEIGHT.get(s1, 2) >= _BORDER_WEIGHT.get(s2, 2) \
            else s2
    return s1 or s2

def sheet_edge_map(ws, max_r, max_c):
    _members, anchors = merge_maps(ws)
    failed = []

    def side(r, c, name):
        if r < 1 or c < 1:
            return None
        try:
            s = getattr(ws.cell(r, c).border, name)
            return s.style if s is not None else None
        except Exception as e:
            failed.append(f"테두리를 읽지 못했습니다({type(e).__name__})")
            return None

    edges = {}
    if is_unknown(anchors):
        return Unknown(unknown_why(anchors))
    for r in range(1, max_r + 2):
        for c in range(1, max_c + 1):
            v = _combine_edge(side(r, c, "top"), side(r - 1, c, "bottom"))
            if v:
                edges[("h", r, c)] = v
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 2):
            v = _combine_edge(side(r, c, "left"), side(r, c - 1, "right"))
            if v:
                edges[("v", r, c)] = v
    for (min_c, min_r, mx_c, mx_r) in anchors.values():
        for r in range(min_r + 1, mx_r + 1):
            for c in range(min_c, mx_c + 1):
                edges.pop(("h", r, c), None)
        for r in range(min_r, mx_r + 1):
            for c in range(min_c + 1, mx_c + 1):
                edges.pop(("v", r, c), None)
    if failed:
        return Unknown(failed[0])
    return edges

def _edge_owner(ws, key):
    if ws is None:
        return None
    try:
        ek, r, c = key
        if ek == "h":
            if r > 1:
                s = ws.cell(r - 1, c).border.bottom
                if s is not None and s.style:
                    return (r - 1, c)
            s = ws.cell(r, c).border.top
            if s is not None and s.style:
                return (r, c)
        else:
            if c > 1:
                s = ws.cell(r, c - 1).border.right
                if s is not None and s.style:
                    return (r, c - 1)
            s = ws.cell(r, c).border.left
            if s is not None and s.style:
                return (r, c)
    except Exception:
        return None
    return None

def merge_maps(ws):
    members = {}
    anchors = {}
    try:
        for rng in ws.merged_cells.ranges:
            min_c, min_r, max_c, max_r = rng.bounds
            anchors[(min_r, min_c)] = (min_c, min_r, max_c, max_r)
            for r in range(min_r, max_r + 1):
                for c in range(min_c, max_c + 1):
                    if (r, c) != (min_r, min_c):
                        members[(r, c)] = (min_r, min_c)
    except Exception as e:
        u = Unknown(f"병합 정보를 읽지 못했습니다({type(e).__name__})")
        return u, u
    return members, anchors

def fmt_signature_m(ws, r, c, kind, members, anchors):
    if any_unknown(members, anchors):
        return Unknown(unknown_why(members, anchors))
    if ws is None:
        return Unknown("시트를 찾지 못해 서식을 읽지 못했습니다")
    if (r, c) in members:
        r, c = members[(r, c)]
    if kind == "border" and (r, c) in anchors:
        min_c, min_r, max_c, max_r = anchors[(r, c)]

        def _has(side, cells):
            for rr, cc in cells:
                try:
                    s = getattr(ws.cell(rr, cc).border, side)
                    if s is not None and s.style:
                        return True
                except Exception:
                    pass
            return False

        return (
            _has("left", [(rr, min_c) for rr in range(min_r, max_r + 1)]),
            _has("right", [(rr, max_c) for rr in range(min_r, max_r + 1)]),
            _has("top", [(min_r, cc) for cc in range(min_c, max_c + 1)]),
            _has("bottom", [(max_r, cc) for cc in range(min_c, max_c + 1)]),
            False,
        )
    return fmt_signature(ws.cell(r, c), kind)

def cell_raw_m(ws, r, c, members, attr):
    if ws is None:
        return None
    if is_unknown(members):
        return Unknown(unknown_why(members))
    if (r, c) in members:
        r, c = members[(r, c)]
    try:
        return getattr(ws.cell(r, c), attr)
    except Exception as e:
        return Unknown(f"{attr} 속성을 읽지 못했습니다({type(e).__name__})")

def nf_display_equal(sha, shs, r, c, a_members, s_members):
    if shs is None:
        return False
    nf_a = cell_raw_m(sha, r, c, a_members, "number_format")
    nf_s = cell_raw_m(shs, r, c, s_members, "number_format")
    if nf_a is None or nf_s is None or any_unknown(nf_a, nf_s):
        return False
    val = cell_raw_m(sha, r, c, a_members, "value")
    if val is None:
        val = cell_raw_m(shs, r, c, s_members, "value")
    return nf_same_display(val, nf_a, nf_s)

H_ALIGN_KO = {"center": "가운데", "centerContinuous": "선택 영역의 가운데로",
              "distributed": "균등 분할", "fill": "채우기", "left": "왼쪽",
              "right": "오른쪽", "justify": "양쪽 맞춤", "general": "일반",
              None: "일반"}
V_ALIGN_KO = {"center": "가운데", "top": "위쪽", "bottom": "아래쪽",
              "justify": "양쪽", "distributed": "균등 분할", None: "기본"}
UNDERLINE_KO = {"double": "이중 실선 밑줄", "single": "밑줄",
                "doubleAccounting": "이중 실선 밑줄(회계형)",
                "singleAccounting": "밑줄(회계형)"}

def describe_font(sig):
    if not sig or sig == "?" or is_unknown(sig):
        return "확인 불가"
    name, size, bold, italic, u, strike, color = sig
    head = f"{name or '기본 글꼴'} {size:g}pt"
    parts = []
    if bold and italic:
        parts.append("굵은 기울임꼴")
    elif bold:
        parts.append("굵게")
    elif italic:
        parts.append("기울임꼴")
    if u:
        parts.append(UNDERLINE_KO.get(u, "밑줄"))
    if strike:
        parts.append("취소선")
    if color:
        if isinstance(color, tuple) and color[0] == "rgb":
            parts.append(f"글꼴 색 {color_name_ko(color[1])}")
        else:
            parts.append("글꼴 색 지정")
    return head + (" " + ", ".join(parts) if parts else "")

def describe_border(sig):
    if not sig or sig == "?" or is_unknown(sig):
        return "확인 불가"
    l, r, t, b = sig[:4]
    if all((l, r, t, b)):
        return "테두리(사방)"
    if not any((l, r, t, b)):
        return "테두리 없음"
    sides = [nm for nm, v in (("왼쪽", l), ("오른쪽", r),
                              ("위쪽", t), ("아래쪽", b)) if v]
    return "테두리(" + "·".join(sides) + ")"

def describe_fill(sig):
    if is_unknown(sig):
        return "확인 불가"
    if sig is None:
        return "채우기 없음"
    if sig == "?":
        return "확인 불가"
    pat, color = sig
    if isinstance(color, tuple) and color and color[0] == "rgb":
        return f"채우기 {color_name_ko(str(color[1])[-6:])}"
    if isinstance(color, tuple) and color and color[0] == "theme":
        return "채우기 있음(테마 색)"
    return "채우기 있음"

def props_for_kind(kind, sig_a, sig_s, raw_a=None, raw_s=None):
    out = []
    if kind == "alignment":
        sa = sig_a if isinstance(sig_a, tuple) else (None, None, False, 0)
        ss = sig_s if isinstance(sig_s, tuple) else (None, None, False, 0)
        if sa[0] != ss[0]:
            out.append({"name": "가로 맞춤",
                        "expected": H_ALIGN_KO.get(sa[0], sa[0]),
                        "got": H_ALIGN_KO.get(ss[0], ss[0])})
        if sa[1] != ss[1]:
            out.append({"name": "세로 맞춤",
                        "expected": V_ALIGN_KO.get(sa[1], sa[1]),
                        "got": V_ALIGN_KO.get(ss[1], ss[1])})
        if sa[2] != ss[2]:
            out.append({"name": "텍스트 줄 바꿈",
                        "expected": "적용" if sa[2] else "해제",
                        "got": "적용" if ss[2] else "해제"})
        if len(sa) > 3 and len(ss) > 3 and sa[3] != ss[3]:
            out.append({"name": "텍스트 회전",
                        "expected": f"{sa[3]}도", "got": f"{ss[3]}도"})
    elif kind == "font":
        out.append({"name": "글꼴", "expected": describe_font(sig_a),
                    "got": describe_font(sig_s)})
    elif kind == "fill":
        out.append({"name": "채우기", "expected": describe_fill(sig_a),
                    "got": describe_fill(sig_s)})
    elif kind == "border":
        out.append({"name": "테두리", "expected": describe_border(sig_a),
                    "got": describe_border(sig_s)})
    elif kind == "number_format":
        exp_raw = raw_a if raw_a not in (None, "") else sig_a
        got_raw = raw_s if raw_s not in (None, "") else sig_s
        exp = nf_display(exp_raw)
        mean = explain_number_format(exp_raw)
        out.append({"name": "표시 형식",
                    "expected": exp + (f" — {mean}" if mean else ""),
                    "got": nf_display(got_raw) if got_raw is not None
                    else "확인 불가"})
    elif kind == "style":
        out.append({"name": "셀 스타일", "expected": str(sig_a),
                    "got": str(sig_s)})
    return out

def hints_for_kind(kind, sig_a, raw_a=None):
    if kind == "alignment" and isinstance(sig_a, tuple):
        h = sig_a[0]
        if h == "distributed":
            return ["Ctrl+1 → 맞춤 탭 → 가로 '균등 분할'"]
        if h == "centerContinuous":
            return ["Ctrl+1 → 맞춤 탭 → 가로 '선택 영역의 가운데로' "
                    "(병합이 아닙니다)"]
        if h == "fill":
            return ["Ctrl+1 → 맞춤 탭 → 가로 '채우기'"]
        if sig_a[2]:
            return ["Ctrl+1 → 맞춤 탭 → '텍스트 줄 바꿈' 체크"]
        return ["홈 탭 맞춤 그룹(가로/세로 가운데)"]
    if kind == "font" and isinstance(sig_a, tuple):
        if sig_a[4] in ("double", "doubleAccounting"):
            return ["Ctrl+1 → 글꼴 탭 → 밑줄 '이중 실선' (회계용 아님 주의)",
                    "홈 탭 글꼴 그룹에서 글꼴 이름·크기·굵게/기울임 지정"]
        return ["홈 탭 글꼴 그룹에서 글꼴 이름·크기·굵게/기울임 지정"]
    if kind == "merge":
        return ["범위를 선택하고 홈 탭 → '병합하고 가운데 맞춤'"]
    if kind == "number_format":
        code = nf_display(raw_a if raw_a not in (None, "") else sig_a)
        return [f"Ctrl+1 → 표시 형식 → 사용자 지정: {code}"]
    if kind == "border":
        return ["홈 탭 → 테두리 → 모든 테두리/바깥쪽 테두리"]
    if kind == "fill":
        return ["홈 탭 → 채우기 색(페인트 통)"]
    if kind == "style":
        return ["홈 탭 → 셀 스타일에서 지정(쉼표/통화 등)"]
    if kind == "value":
        return ["문제 지시 값을 정확히 입력했는지 확인"]
    return []

def compress_coords(coords):
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.cell import coordinate_to_tuple
    pts = set()
    for co in coords:
        try:
            if ":" in str(co):
                pts.add(str(co))
                continue
            pts.add(coordinate_to_tuple(str(co)))
        except Exception:
            pts.add(str(co))
    ranges = sorted(p for p in pts if isinstance(p, str))
    cells = sorted(p for p in pts if isinstance(p, tuple))
    out = []
    used = set()
    bycol = {}
    for (r, c) in cells:
        bycol.setdefault(c, []).append(r)
    for c in sorted(bycol):
        rows = sorted(bycol[c])
        run = [rows[0]]
        for r in rows[1:] + [None]:
            if r is not None and r == run[-1] + 1:
                run.append(r)
                continue
            if len(run) >= 2:
                L = get_column_letter(c)
                out.append(f"{L}{run[0]}:{L}{run[-1]}")
                used.update((rr, c) for rr in run)
            if r is not None:
                run = [r]
    remaining = [p for p in cells if p not in used]
    byrow = {}
    for (r, c) in remaining:
        byrow.setdefault(r, []).append(c)
    for r in sorted(byrow):
        cols = sorted(byrow[r])
        run = [cols[0]]
        for c in cols[1:] + [None]:
            if c is not None and c == run[-1] + 1:
                run.append(c)
                continue
            if len(run) >= 2:
                out.append(f"{get_column_letter(run[0])}{r}:"
                           f"{get_column_letter(run[-1])}{r}")
                used.update((r, cc) for cc in run)
            elif (r, run[0]) not in used:
                out.append(f"{get_column_letter(run[0])}{r}")
                used.add((r, run[0]))
            if c is not None:
                run = [c]
    return sorted(set(out + ranges))

def _cluster_units(units, target=None):
    if not units:
        return []
    clusters = [[u] for u in units]

    def near(c1, c2):
        for a in c1:
            for b in c2:
                dr = abs(a["pos"][0] - b["pos"][0])
                if dr <= 1 and dr + abs(a["pos"][1] - b["pos"][1]) <= 2:
                    return True
        return False

    merged = True
    while merged:
        merged = False
        for i in range(len(clusters)):
            for j in range(i + 1, len(clusters)):
                if near(clusters[i], clusters[j]):
                    clusters[i].extend(clusters[j])
                    del clusters[j]
                    merged = True
                    break
            if merged:
                break
    for cl in clusters:
        cl.sort(key=lambda u: (u["pos"][0], u["pos"][1]))
    clusters.sort(key=lambda cl: (cl[0]["pos"][0], cl[0]["pos"][1]))
    return clusters

def norm_ref(ref):
    if ref is None:
        return None
    s = str(ref).upper().replace("$", "").replace("'", "")
    s = re.sub(r"\s+", "", s)
    s = s.replace("정답!", "!").replace("정답".upper() + "!", "!")
    return s

def norm_range_only(ref):
    if ref is None:
        return None
    s = str(ref)
    if "!" in s:
        s = s.rsplit("!", 1)[1]
    return s.upper().replace("$", "").replace("'", "").strip()

_CF_REF_RE = re.compile(
    r"(?<![A-Za-z0-9_.\$])(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(])")

CF_TYPE_KO = {
    "expression": "수식을 사용하여 서식을 지정할 셀 결정",
    "cellIs": "셀 값 기준", "containsText": "특정 텍스트 포함",
    "notContainsText": "특정 텍스트 미포함", "beginsWith": "시작 문자",
    "endsWith": "끝 문자", "duplicateValues": "중복 값",
    "uniqueValues": "고유 값", "top10": "상위/하위", "aboveAverage": "평균 초과/미만",
    "colorScale": "색조", "dataBar": "데이터 막대", "iconSet": "아이콘 집합",
    "containsBlanks": "빈 셀", "notContainsBlanks": "빈 셀 아님",
    "containsErrors": "오류", "timePeriod": "발생 날짜",
}

CF_OP_KO = {
    "lessThan": ("<", "미만"), "lessThanOrEqual": ("<=", "이하"),
    "equal": ("=", "같음"), "notEqual": ("<>", "같지 않음"),
    "greaterThanOrEqual": (">=", "이상"), "greaterThan": (">", "초과"),
    "between": ("between", "사이"), "notBetween": ("notBetween", "사이 아님"),
}

def _cf_op_text(op):
    sym, ko = CF_OP_KO.get(op or "", (op or "?", ""))
    return f"{sym} ({ko})" if ko else str(sym)

def _cf_cond_text(rule):
    fs = rule.get("formulas") or []
    if rule.get("type") == "cellIs":
        op = rule.get("operator")
        if op in ("between", "notBetween") and len(fs) >= 2:
            return f"셀 값 {fs[0]} ~ {fs[1]} {CF_OP_KO[op][1]}"
        sym = CF_OP_KO.get(op or "", (op or "?", ""))[0]
        return f"셀 값 {sym} {fs[0] if fs else '?'}"
    if fs:
        return "=" + fs[0]
    return f"({CF_TYPE_KO.get(rule.get('type'), rule.get('type'))})"

CF_ROW_PRINCIPLE = (
    "'행 전체'에 서식을 적용하려면 ① 적용 범위는 표 전체(예: A4:H18)를 선택한 "
    "뒤 규칙을 만들고 ② 수식은 열만 고정한 혼합참조(예: =LEFT($A4,4)=\"2019\")"
    "로 씁니다. 그래야 B~H열의 셀도 각 행의 A열 값을 검사해 행 전체가 같은 "
    "색이 됩니다. A열만 선택하거나(범위 오답) 상대참조 A4를 쓰면(참조 오답) "
    "실제 시험에서는 오답입니다.")

def cf_row_principle(rule, ws_a=None):
    try:
        from openpyxl.utils import column_index_from_string, \
            get_column_letter, range_boundaries
        sq = (rule.get("sqref") or "").split()[0]
        formula = (rule.get("formulas") or [""])[0]
        m = re.search(r"\$([A-Z]{1,3})(?!\$)(\d+)", formula.upper())
        if not (sq and formula and m):
            return CF_ROW_PRINCIPLE
        col, row = m.group(1), m.group(2)
        c1, r1, c2, _r2 = range_boundaries(sq if ":" in sq else f"{sq}:{sq}")
        first, last = get_column_letter(c1), get_column_letter(c2)
        head = ""
        if ws_a is not None:
            try:
                hv = ws_a.cell(r1 - 1, column_index_from_string(col)).value
                if isinstance(hv, str) and hv.strip():
                    head = f"({hv.strip()})"
            except Exception:
                head = ""
        txt = (
            f"'행 전체'에 서식을 적용하려면 ① 적용 범위를 표 전체({sq})로 선택한 "
            f"뒤 규칙을 만들고 ② 수식은 기준 열만 고정한 혼합참조(${col}{row})로 "
            f"씁니다. 그래야 {first}~{last}열의 모든 셀이 같은 행의 "
            f"{col}열{head}을 검사해 행 전체가 같은 서식이 됩니다. "
            f"$를 빼고 {col}{row}으로 쓰면 {first}열 셀은 {first}{row}을, "
            f"그 옆 열 셀은 자기 열의 셀을 검사하게 되어 행 전체가 칠해지지 "
            f"않고, 반대로 ${col}${row}처럼 행까지 고정하면 모든 행이 "
            f"{row}행만 검사합니다.")
        if "AVERAGE(" in formula.upper():
            rng = re.search(r"AVERAGE\(([^)]*)\)", formula.upper())
            inner = rng.group(1) if rng else ""
            txt += (f" 평균 기준은 AVERAGE({inner})처럼 범위를 행·열 모두 "
                    "고정($)해야 모든 행이 같은 평균과 비교합니다.")
        return txt
    except Exception:
        return CF_ROW_PRINCIPLE

def shift_formula(formula, dr, dc):
    if not formula or (dr == 0 and dc == 0):
        return formula
    from openpyxl.utils import column_index_from_string, get_column_letter
    parts = re.split(r'("(?:[^"]|"")*")', str(formula))
    out = []
    for i, p in enumerate(parts):
        if i % 2 == 1:
            out.append(p)
            continue

        def _rep(m):
            cabs, col, rabs, row = m.groups()
            try:
                ci = column_index_from_string(col.upper())
                ri = int(row)
            except Exception:
                return m.group(0)
            if not cabs:
                ci += dc
            if not rabs:
                ri += dr
            if ci < 1 or ri < 1 or ci > 16384 or ri > 1048576:
                return "#REF!"
            return f"{cabs}{get_column_letter(ci)}{rabs}{ri}"
        out.append(_CF_REF_RE.sub(_rep, p))
    return "".join(out)

def _dxf_sig(dxf, wb=None):
    sig = {}
    if dxf is None:
        return sig
    try:
        f = dxf.font
        if f is not None:
            col = _color_key(f.color, wb)
            if col:
                sig["글꼴 색"] = col
            if f.b:
                sig["굵게"] = True
            if f.i:
                sig["기울임꼴"] = True
            u = f.u
            if u and u != "none":
                sig["밑줄"] = str(u)
            if f.strike:
                sig["취소선"] = True
    except Exception as e:
        return Unknown(f"조건부 서식의 글꼴을 읽지 못했습니다"
                       f"({type(e).__name__})")
    try:
        fl = dxf.fill
        if fl is not None:
            col = _color_key(getattr(fl, "bgColor", None), wb) or \
                _color_key(getattr(fl, "fgColor", None), wb)
            if col:
                sig["채우기 색"] = col
    except Exception as e:
        return Unknown(f"조건부 서식의 채우기를 읽지 못했습니다"
                       f"({type(e).__name__})")
    return sig

def _dxf_desc(sig):
    if not sig:
        return "(서식 없음)"
    parts = []
    for k, v in sig.items():
        if isinstance(v, tuple) and v and v[0] == "rgb":
            parts.append(f"{k} {color_name_ko(v[1])}")
        elif isinstance(v, tuple):
            parts.append(f"{k} 테마색{v[1:]}")
        elif v is True:
            parts.append(k)
        else:
            parts.append(f"{k} {v}")
    return ", ".join(parts)

def _dxf_compatible(exp, got):
    for k, v in (exp or {}).items():
        if (got or {}).get(k) != v:
            return False
    return True

def cf_rules_detail(ws, wb=None):
    rules = []
    try:
        for cf in ws.conditional_formatting:
            sq = str(cf.sqref)
            sq_disp = sq.replace("$", "").upper().strip()
            for r in cf.rules:
                formulas = []
                try:
                    formulas = [str(f) for f in (r.formula or [])]
                except Exception:
                    pass
                key = (norm_ref(sq), r.type, getattr(r, "operator", None),
                       tuple(sorted(norm_formula("=" + f) for f in formulas)))
                dxf_sig = _dxf_sig(getattr(r, "dxf", None), wb)
                if is_unknown(dxf_sig):
                    return Unknown(unknown_why(dxf_sig))
                rules.append({
                    "sqref": sq_disp, "ranges": sq_disp.split(),
                    "type": r.type, "operator": getattr(r, "operator", None),
                    "formulas": formulas,
                    "dxf": dxf_sig,
                    "key": key})
    except Exception as e:
        return Unknown("조건부 서식 규칙을 읽지 못했습니다"
                       f"({type(e).__name__})")
    return rules

def _cf_rule_cells(rule, limit=20000):
    from openpyxl.utils import range_boundaries
    cells = []
    anchor = None
    for rng in rule["ranges"]:
        try:
            c1, r1, c2, r2 = range_boundaries(rng if ":" in rng else f"{rng}:{rng}")
        except Exception:
            continue
        if None in (c1, r1, c2, r2):
            return anchor, cells, True
        if anchor is None:
            anchor = (r1, c1)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cells.append((r, c))
                if len(cells) > limit:
                    return anchor, cells, True
    return anchor, cells, False

def _cf_effective(rule, r, c, anchor):
    ar, ac = anchor
    return tuple(sorted(norm_formula("=" + shift_formula(f, r - ar, c - ac))
                        for f in rule["formulas"]))

def cf_effective_map(rules):
    m = {}
    overflow = False
    for idx, rule in enumerate(rules):
        anchor, cells, over = _cf_rule_cells(rule)
        if over or anchor is None:
            overflow = True
            continue
        for (r, c) in cells:
            m.setdefault((r, c), []).append(
                (rule["type"], rule["operator"] or None,
                 _cf_effective(rule, r, c, anchor), rule["dxf"], idx))
    return m, overflow

def cf_compare_rule(rule, rules_s, map_s):
    out = {"ok": False, "missing": [], "type_bad": [], "operator_bad": [],
           "formula_bad": [], "dxf_bad": [], "fallback": False}
    anchor, cells, over = _cf_rule_cells(rule)
    if over or anchor is None:
        out["fallback"] = True
        out["ok"] = any(s["key"] == rule["key"] for s in rules_s)
        return out
    for (r, c) in cells:
        eff = _cf_effective(rule, r, c, anchor)
        entries = map_s.get((r, c)) or []
        if not entries:
            out["missing"].append((r, c))
            continue
        same_kind = [e for e in entries if e[0] == rule["type"]]
        if not same_kind:
            out["type_bad"].append((r, c))
            continue
        same_type = [e for e in same_kind
                     if (e[1] or None) == (rule["operator"] or None)]
        if not same_type:
            same_f_other_op = [e for e in same_kind if e[2] == eff]
            if same_f_other_op:
                out["operator_bad"].append((r, c))
                for e in same_f_other_op:
                    if e[1] not in out.setdefault("operators", []):
                        out["operators"].append(e[1])
            else:
                out["formula_bad"].append((r, c))
            continue
        same_f = [e for e in same_type if e[2] == eff]
        if not same_f:
            out["formula_bad"].append((r, c))
            continue
        if not any(_dxf_compatible(rule["dxf"], e[3]) for e in same_f):
            out["dxf_bad"].append((r, c))
    out["ok"] = not (out["missing"] or out["type_bad"] or out["operator_bad"]
                     or out["formula_bad"] or out["dxf_bad"])
    return out

def _cf_rule_text(rule):
    return f"{rule['sqref']}: {_cf_cond_text(rule)} · {_dxf_desc(rule['dxf'])}"

def cf_rule_set(ws):
    rules = set()
    try:
        for cf in ws.conditional_formatting:
            sqref = norm_ref(str(cf.sqref))
            for r in cf.rules:
                formulas = ()
                try:
                    if getattr(r, "formula", None):
                        formulas = tuple(sorted(norm_formula("=" + f) for f in r.formula))
                except Exception:
                    pass
                rules.add((sqref, r.type, getattr(r, "operator", None), formulas))
    except Exception:
        pass
    return rules

def scan_bounds(*sheets):
    mr = max([s.max_row or 1 for s in sheets] + [1])
    mc = max([s.max_column or 1 for s in sheets] + [1])
    return min(mr, MAX_SCAN_ROWS), min(mc, MAX_SCAN_COLS)

def cell_differs(raw_p, raw_a):
    fp, fa = formula_text(raw_p), formula_text(raw_a)
    if fp is not None or fa is not None:
        if (fp is None) != (fa is None):
            return True
        return norm_formula(fp) != norm_formula(fa)
    return not value_eq(raw_p, raw_a)

def value_diff_cells(shp, sha):
    mr, mc = scan_bounds(shp, sha)
    out = []
    for r in range(1, mr + 1):
        for c in range(1, mc + 1):
            if cell_differs(shp.cell(r, c).value, sha.cell(r, c).value):
                out.append((r, c, sha.cell(r, c).coordinate))
    return out

def _blank_answer_cell(sha, r, c):
    try:
        v = sha.cell(r, c).value
    except Exception:
        return False
    return v is None or (isinstance(v, str) and v.strip() == "")

BLANK_SKIP_KINDS = ("font", "number_format")

def _blank_align_sig(sig):
    if isinstance(sig, tuple) and sig and sig[0] == "centerContinuous":
        return ("centerContinuous",)
    return None

def format_diff_items(book_p, book_a, psheet, asheet):
    shp, sha = book_p.raw[psheet], book_a.raw[asheet]
    mr, mc = scan_bounds(shp, sha)
    items = {}
    unknown = []
    p_members, p_anchors = merge_maps(shp)
    a_members, a_anchors = merge_maps(sha)
    if any_unknown(p_members, a_members):
        unknown.append((psheet, "merge", hold_reason(p_members, a_members)))
    for r in range(1, mr + 1):
        for c in range(1, mc + 1):
            if (r, c) in a_members or (r, c) in p_members:
                continue
            blank_a = _blank_answer_cell(sha, r, c)
            for kind in CELL_FMT_KINDS:
                if blank_a and kind in BLANK_SKIP_KINDS:
                    continue
                sp = fmt_signature_m(shp, r, c, kind, p_members, p_anchors)
                sa = fmt_signature_m(sha, r, c, kind, a_members, a_anchors)
                if blank_a and kind == "alignment":
                    sp, sa = _blank_align_sig(sp), _blank_align_sig(sa)
                if any_unknown(sp, sa):
                    unknown.append((sha.cell(r, c).coordinate, kind,
                                    hold_reason(sp, sa)))
                    continue
                if sp != sa:
                    items.setdefault(kind, []).append(
                        sha.cell(r, c).coordinate)
    from openpyxl.utils import get_column_letter
    p_edges = sheet_edge_map(shp, mr, mc)
    a_edges = sheet_edge_map(sha, mr, mc)
    if any_unknown(p_edges, a_edges):
        unknown.append((asheet, "border", hold_reason(p_edges, a_edges)))
        p_edges = a_edges = {}
    border_edges = {}
    for key in set(p_edges) | set(a_edges):
        av, pv = a_edges.get(key), p_edges.get(key)
        if av != pv:
            ek, r, c = key
            owner = _edge_owner(sha, key) or _edge_owner(shp, key)
            if owner:
                rr, cc = owner
            else:
                rr, cc = min(r, mr), min(c, mc)
            coord = f"{get_column_letter(cc)}{rr}"
            border_edges.setdefault(coord, []).append((key, av))
    if border_edges:
        items["border"] = sorted(border_edges)
        items["border_edges"] = border_edges
    merges_p = {str(x).upper() for x in shp.merged_cells.ranges}
    merges_a = {str(x).upper() for x in sha.merged_cells.ranges}
    if merges_a - merges_p:
        items["merge"] = sorted(merges_a - merges_p)
    if merges_p - merges_a:
        items["merge_del"] = sorted(merges_p - merges_a)
    rows = set(shp.row_dimensions) | set(sha.row_dimensions)
    rh = []
    for r in sorted(rows):
        hp = shp.row_dimensions[r].height if r in shp.row_dimensions else None
        ha = sha.row_dimensions[r].height if r in sha.row_dimensions else None
        if _size_differs(hp, ha):
            rh.append((r, ha))
    if rh:
        items["rowheight"] = rh
    cols = set(shp.column_dimensions) | set(sha.column_dimensions)
    cw = []
    for cl in sorted(cols):
        wp = shp.column_dimensions[cl].width if cl in shp.column_dimensions else None
        wa = sha.column_dimensions[cl].width if cl in sha.column_dimensions else None
        if not col_width_same(wp, wa):
            cw.append((cl, wa))
    if cw:
        items["colwidth"] = cw
    names = defined_names_for_sheet(book_a, asheet)
    names_p = defined_names_for_sheet(book_p, psheet)
    if any_unknown(names, names_p):
        unknown.append((asheet, "names", hold_reason(names, names_p)))
    else:
        new_names = {k: v for k, v in names.items() if names_p.get(k) != v}
        if new_names:
            items["names"] = new_names
    if unknown:
        items["_unknown"] = unknown
    return items

def _size_differs(a, b):
    if a is None and b is None:
        return False
    if a is None or b is None:
        return True
    return abs(a - b) > 0.5

def _size_matches(a, b):
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    return abs(a - b) <= 0.5

COL_WIDTH_MDW = (8, 7)
COL_WIDTH_PAD = 5
COL_WIDTH_TOL = 0.5
ROW_HEIGHT_TOL = 0.75

def col_width_stored(chars, mdw=8):
    return math.trunc((float(chars) * mdw + COL_WIDTH_PAD) / mdw * 256) / 256

def col_width_candidates(w):
    if w is None:
        return []
    w = float(w)
    return [round(w, 2)] + [round(w - float(COL_WIDTH_PAD) / m, 2)
                            for m in COL_WIDTH_MDW]

def _col_width_rank(c):
    for i, q in enumerate((1.0, 0.5, 0.25, 0.1)):
        if abs(c / q - round(c / q)) < 1e-6:
            return i
    return 4

def col_width_chars(w, want=None):
    cands = col_width_candidates(w)
    if not cands:
        return None
    if want is None:
        return min(cands, key=_col_width_rank)
    return min(cands, key=lambda c: (_col_width_rank(c),
                                     abs(c - float(want))))

def col_width_is(stored, want_chars, tol=COL_WIDTH_TOL):
    if stored is None or want_chars is None:
        return stored is None and want_chars is None
    return any(abs(c - float(want_chars)) <= tol
               for c in col_width_candidates(stored))

def col_width_same(a, b, tol=COL_WIDTH_TOL):
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    ca, cb = col_width_candidates(a), col_width_candidates(b)
    return any(abs(x - y) <= tol for x in ca for y in cb)

def col_width_text(w, want=None):
    if w is None:
        return "지정 안 함(기본 너비)"
    ch = col_width_chars(w, want)
    if ch is None:
        return f"{float(w):g}"
    if abs(ch - float(w)) < 5e-3:
        return f"{ch:g}자"
    return f"{ch:g}자(저장값 {float(w):g})"

def row_height_text(h):
    if h is None:
        return "지정 안 함(기본 높이)"
    return f"{float(h):g}"

def _book_defined_name(book, name):
    try:
        for nm, dn in book.raw.defined_names.items():
            if str(nm) == str(name):
                return norm_ref(dn.value)
    except Exception:
        pass
    return None

def _ref_areas(ref):
    if not ref:
        return []
    out = []
    for part in str(ref).split(","):
        part = part.strip()
        if not part:
            continue
        out.append(part.rsplit("!", 1)[-1])
    return out

_RESERVED_NAME_RE = re.compile(r"^_xl(?:nm|eta|udf|fn)?[._]", re.I)

def _is_reserved_name(name):
    return bool(_RESERVED_NAME_RE.match(str(name or "")))

def defined_names_for_sheet(book, sheet_name):
    out = {}
    key = norm_sheet_name(sheet_name)
    try:
        for name, dn in book.raw.defined_names.items():
            if _is_reserved_name(name):
                continue
            val = norm_ref(dn.value)
            if val and key.upper() in norm_sheet_name(val.split("!")[0]).upper():
                out[str(name)] = val
    except Exception as e:
        return Unknown(f"{book.label} 파일의 통합 문서 이름 정의를 읽지 "
                       f"못했습니다({type(e).__name__})")
    try:
        ws = book.raw[sheet_name]
        for name, dn in getattr(ws, "defined_names", {}).items():
            if _is_reserved_name(name):
                continue
            val = norm_ref(dn.value)
            if val:
                out.setdefault(str(name), val)
    except KeyError:
        pass
    except Exception as e:
        return Unknown(f"{book.label} 파일의 '{sheet_name}' 시트 이름 정의를 "
                       f"읽지 못했습니다({type(e).__name__})")
    return out

class CellJudge:

    def __init__(self, book_a, book_s, asheet, ssheet, key_cells,
                 display_tolerant=False):
        self.sha_raw = book_a.raw[asheet]
        self.sha_val = book_a.cached[asheet]
        self.shs_raw = book_s.raw[ssheet] if ssheet else None
        self.shs_val = book_s.cached[ssheet] if ssheet else None
        self.key_cells = key_cells or {}
        self.formula_only_used = False
        self.book_a, self.book_s = book_a, book_s
        self.asheet, self.ssheet = asheet, ssheet
        self.recalc_cells = []
        self.recalc_failed = False
        self.recalc_vals_s = {}
        self.recalc_vals_a = {}
        self.ref_mismatch = []
        self.display_tolerant = display_tolerant
        self.display_matched = []

    def expected(self, r, c):
        coord = self.sha_raw.cell(r, c).coordinate
        raw = self.sha_raw.cell(r, c).value
        f = formula_text(raw)
        ov = self.key_cells.get(coord)
        if ov is not None:
            if ov.get("formula"):
                f = ov["formula"]
            if ov.get("value") is not None:
                return ov["value"], f
        if f is not None:
            cached = self.sha_val.cell(r, c).value
            if cached is not None:
                return cached, f
            return "__MISSING__", f
        return raw, None

    def student(self, r, c):
        if self.shs_raw is None:
            return "__MISSING__", None
        raw = self.shs_raw.cell(r, c).value
        f = formula_text(raw)
        if f is not None:
            cached = self.shs_val.cell(r, c).value
            if cached is not None:
                return cached, f
            return "__MISSING__", f
        return raw, None

    def _recalc(self, r, c, ef, sf):
        got_v = exp_v = "__MISSING__"
        if self.ssheet and sf:
            got_v = self.book_s.recalc_value(self.ssheet, r, c)
        if got_v == "__MISSING__":
            self.recalc_failed = True
            return "__MISSING__", "__MISSING__"
        self.recalc_vals_s[(r, c)] = got_v
        if ef:
            exp_v = self.book_a.recalc_value(self.asheet, r, c)
            if exp_v != "__MISSING__":
                self.recalc_vals_a[(r, c)] = exp_v
        return exp_v, got_v

    def judge(self, r, c):
        exp, ef = self.expected(r, c)
        got, sf = self.student(r, c)
        if self.shs_raw is None:
            return False, exp, None, ef
        used_recalc = False
        if got == "__MISSING__" and sf:
            exp_v, got_v = self._recalc(r, c, ef, sf)
            if got_v != "__MISSING__":
                got = got_v
                used_recalc = True
                if exp == "__MISSING__" and exp_v != "__MISSING__":
                    exp = exp_v
                coord = self.sha_raw.cell(r, c).coordinate
                if coord not in self.recalc_cells:
                    self.recalc_cells.append(coord)
        if exp != "__MISSING__" and got != "__MISSING__":
            ok = value_eq(exp, got)
            if not ok and ef and sf and norm_formula(ef) == norm_formula(sf):
                ok = True
            if ok and used_recalc and ef and sf \
                    and norm_formula(ef) != norm_formula(sf) \
                    and not same_ref_usage(ef, sf):
                ok = False
                coord = self.sha_raw.cell(r, c).coordinate
                if not any(m[0] == coord for m in self.ref_mismatch):
                    self.ref_mismatch.append((coord, ef, sf))
            if not ok and self.display_tolerant:
                try:
                    nf = self.sha_raw.cell(r, c).number_format
                except Exception:
                    nf = None
                if value_eq_display(exp, got, nf):
                    ok = True
                    coord = self.sha_raw.cell(r, c).coordinate
                    if coord not in self.display_matched:
                        self.display_matched.append(coord)
            return ok, exp, got, ef
        if ef is not None or sf is not None:
            self.formula_only_used = True
            ok = ef is not None and sf is not None and \
                norm_formula(ef) == norm_formula(sf)
            return ok, exp if exp != "__MISSING__" else None, sf or got, ef
        e = None if exp == "__MISSING__" else exp
        g = None if got == "__MISSING__" else got
        return value_eq(e, g), e, g, ef

def expand_ref(ref):
    from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries
    ref = str(ref).replace("$", "").replace(" ", "").upper()
    if ":" in ref:
        mn_c, mn_r, mx_c, mx_r = range_boundaries(ref)
        return [(r, c) for r in range(mn_r, mx_r + 1)
                for c in range(mn_c, mx_c + 1)]
    r, c = coordinate_to_tuple(ref)
    return [(r, c)]

def _unwrap_cell_override(v):
    if isinstance(v, dict) and ("value" in v or "formula" in v):
        return {"value": v.get("value"), "formula": v.get("formula")}
    return {"value": v, "formula": None}

def normalize_key(key):
    out = {"points": {}, "cells": {}, "groups": {}}
    for k, v in (key.get("points") or {}).items():
        out["points"][norm_sheet_name(k)] = v
    for k, v in (key.get("cells") or {}).items():
        if "!" in str(k):
            sheet, coord = str(k).rsplit("!", 1)
            out["cells"].setdefault(norm_sheet_name(sheet), {})[
                coord.replace("$", "").strip().upper()] = _unwrap_cell_override(v)
        elif isinstance(v, dict):
            tgt = out["cells"].setdefault(norm_sheet_name(k), {})
            for coord, cv in v.items():
                tgt[str(coord).replace("$", "").strip().upper()] = \
                    _unwrap_cell_override(cv)
    for k, v in (key.get("groups") or {}).items():
        out["groups"][norm_sheet_name(k)] = v
    out["format_checks"] = {}
    for k, v in (key.get("format_checks") or {}).items():
        if isinstance(v, list):
            out["format_checks"][norm_sheet_name(k)] = v
    out["exclude"] = {}
    for k, v in (key.get("exclude") or {}).items():
        items = []
        for it in (v if isinstance(v, list) else [v]):
            if isinstance(it, str):
                it = {"range": it}
            if not isinstance(it, dict) or not it.get("range"):
                continue
            kinds = it.get("kinds")
            if isinstance(kinds, str):
                kinds = [kinds]
            items.append({"range": str(it["range"]).strip(),
                          "kinds": [str(x) for x in kinds] if kinds else None,
                          "why": it.get("why")})
        if items:
            out["exclude"][norm_sheet_name(k)] = items
    out["chart"] = {}
    for k, v in (key.get("chart") or {}).items():
        if isinstance(v, dict):
            out["chart"][norm_sheet_name(k)] = v
    return out

EXCLUDE_FMT_KINDS = ("number_format", "font", "fill", "alignment", "border",
                     "merge", "rowheight", "colwidth")
EXCLUDE_KIND_KO = {"number_format": "표시 형식", "font": "글꼴", "fill": "채우기",
                   "alignment": "맞춤", "border": "테두리", "merge": "병합",
                   "rowheight": "행 높이", "colwidth": "열 너비", "value": "값"}

def apply_key_exclusions(vdiffs, fdiffs, items, label="세트 기대값에 따라 채점 제외"):
    from openpyxl.utils import get_column_letter
    notes = []
    for it in items or []:
        try:
            cells = set(expand_ref(it["range"]))
        except Exception as e:
            notes.append(f"채점 제외 범위 '{it.get('range')}'를 해석하지 "
                         f"못해 적용하지 못했습니다({type(e).__name__}) — "
                         "이 범위는 그대로 채점됩니다. 기대값 파일을 "
                         "확인하세요.")
            continue
        kinds = it.get("kinds") or list(EXCLUDE_FMT_KINDS)
        coords = {f"{get_column_letter(c)}{r}" for (r, c) in cells}
        removed = []
        for kind in kinds:
            if kind == "value":
                before = len(vdiffs)
                vdiffs = [d for d in vdiffs if (d[0], d[1]) not in cells]
                if len(vdiffs) != before:
                    removed.append(kind)
            elif kind in CELL_FMT_KINDS and fdiffs.get(kind):
                keep = [co for co in fdiffs[kind] if co not in coords]
                if len(keep) != len(fdiffs[kind]):
                    removed.append(kind)
                fdiffs[kind] = keep
            elif kind == "border" and fdiffs.get("border"):
                keep = [co for co in fdiffs["border"] if co not in coords]
                if len(keep) != len(fdiffs["border"]):
                    removed.append(kind)
                fdiffs["border"] = keep
                fdiffs["border_edges"] = {
                    co: e for co, e in (fdiffs.get("border_edges") or {}).items()
                    if co not in coords}
            elif kind == "merge":
                for mk in ("merge", "merge_del"):
                    if not fdiffs.get(mk):
                        continue
                    keep = []
                    for rng in fdiffs[mk]:
                        try:
                            anchor = expand_ref(rng)[0]
                        except Exception:
                            anchor = None
                        if anchor in cells:
                            removed.append(kind)
                        else:
                            keep.append(rng)
                    fdiffs[mk] = keep
            elif kind == "rowheight" and fdiffs.get("rowheight"):
                rows = {r for (r, _c) in cells}
                keep = [x for x in fdiffs["rowheight"] if x[0] not in rows]
                if len(keep) != len(fdiffs["rowheight"]):
                    removed.append(kind)
                fdiffs["rowheight"] = keep
            elif kind == "colwidth" and fdiffs.get("colwidth"):
                cols = {get_column_letter(c) for (_r, c) in cells}
                keep = [x for x in fdiffs["colwidth"] if x[0] not in cols]
                if len(keep) != len(fdiffs["colwidth"]):
                    removed.append(kind)
                fdiffs["colwidth"] = keep
        if removed:
            kinds_ko = ", ".join(dict.fromkeys(
                EXCLUDE_KIND_KO.get(k, k) for k in removed))
            why = f" — {it['why']}" if it.get("why") else ""
            notes.append(f"{label}: {it['range']} {kinds_ko}{why}")
    return vdiffs, fdiffs, notes

AUTHOR_MEMO_MAX_LEN = 40
_MEMO_RANGE_RE = re.compile(r"\$?([A-Z]{1,3})\$?([0-9]{1,7})"
                            r":\$?([A-Z]{1,3})\$?([0-9]{1,7})")

def _memo_range_on_sheet(sha, text):
    from openpyxl.utils import column_index_from_string
    try:
        mr, mc = sha.max_row, sha.max_column
    except Exception:
        return None
    for m in _MEMO_RANGE_RE.finditer(text.upper()):
        try:
            c1 = column_index_from_string(m.group(1))
            r1 = int(m.group(2))
            c2 = column_index_from_string(m.group(3))
            r2 = int(m.group(4))
        except Exception:
            continue
        if 1 <= r1 <= mr and 1 <= r2 <= mr and 1 <= c1 <= mc and 1 <= c2 <= mc:
            return m.group(0)
    return None

def detect_author_memo_cells(shp, sha, vdiffs):
    out = []
    for (r, c, coord) in list(vdiffs):
        try:
            pv = shp.cell(r, c).value
            av = sha.cell(r, c).value
        except Exception:
            continue
        if pv is not None and str(pv).strip() != "":
            continue
        if not isinstance(av, str):
            continue
        text = av.strip()
        if not text or len(text) > AUTHOR_MEMO_MAX_LEN or text.startswith("="):
            continue
        rng = _memo_range_on_sheet(sha, text)
        if not rng:
            continue
        out.append({"range": coord,
                    "kinds": ["value", "number_format", "font", "fill",
                              "alignment", "border", "merge"],
                    "why": f"정답 파일에만 있는 짧은 메모 '{text}' — 수식에서나 "
                           f"쓰는 범위 참조({rng})를 담고 있어 제작자가 남긴 "
                           "작업 메모로 봅니다 (문제지에 입력 지시가 "
                           "없습니다)"})
    return out

def cluster_cells(cells, target):
    if not cells:
        return []
    clusters = [[cell] for cell in cells]

    def dist(c1, c2):
        return min(
            max(abs(a[0] - b[0]), abs(a[1] - b[1])) for a in c1 for b in c2
        )

    merged = True
    while merged:
        merged = False
        for i in range(len(clusters)):
            for j in range(i + 1, len(clusters)):
                if dist(clusters[i], clusters[j]) <= 1:
                    clusters[i].extend(clusters[j])
                    del clusters[j]
                    merged = True
                    break
            if merged:
                break
    while len(clusters) > target:
        best = None
        for i in range(len(clusters)):
            for j in range(i + 1, len(clusters)):
                d = dist(clusters[i], clusters[j])
                if best is None or d < best[0]:
                    best = (d, i, j)
        _, i, j = best
        clusters[i].extend(clusters[j])
        del clusters[j]
    for cl in clusters:
        cl.sort(key=lambda x: (x[0], x[1]))
    clusters.sort(key=lambda cl: (cl[0][0], cl[0][1]))
    return clusters

def _dlbls_active(dl):
    if dl is None:
        return False
    d = dl.find("c:delete", NS)
    if d is not None and d.get("val") in ("1", "true"):
        return False
    shows = [ch for ch in dl if ch.tag.split("}")[1].startswith("show")]
    if shows and all(ch.get("val") in ("0", "false") for ch in shows):
        return False
    return True

def _cell_value_from_ref(book, ref):
    try:
        s = str(ref).strip().replace("$", "")
        if "!" not in s:
            return None
        sheet, coord = s.rsplit("!", 1)
        sheet = sheet.strip("'")
        real = book.norm_map.get(norm_sheet_name(sheet))
        if real is None:
            return None
        v = book.cached[real][coord].value
        if v is None:
            v = book.raw[real][coord].value
        return str(v).strip() if v is not None else None
    except Exception:
        return None

def chart_features(book, sheet_name):
    parts = book.chart_parts_for_sheet(sheet_name)
    if is_unknown(parts):
        return Unknown(unknown_why(parts))
    if not parts:
        return None
    try:
        root = ET.fromstring(book.zf.read(parts[0]))
    except Exception as e:
        return Unknown(f"{book.label} 파일의 차트 파트 {parts[0]} 를 해석하지 "
                       f"못했습니다({type(e).__name__})")
    plot = root.find(".//c:plotArea", NS)
    if plot is None:
        return Unknown(f"{book.label} 파일의 차트 {parts[0]} 에서 그림 "
                       "영역(plotArea)을 찾지 못했습니다")
    types = []
    series = []
    for child in plot:
        tag = child.tag.split("}")[1]
        if not tag.endswith("Chart"):
            continue
        types.append(tag)
        for ser in child.findall("c:ser", NS):
            ref = None
            for path in (".//c:val/c:numRef/c:f", ".//c:yVal/c:numRef/c:f",
                         ".//c:val//c:f"):
                e = ser.find(path, NS)
                if e is not None:
                    ref = e.text
                    break
            has_dlbls = _dlbls_active(ser.find("c:dLbls", NS))
            series.append((norm_range_only(ref), has_dlbls))
    title = None
    t = root.find("c:chart/c:title", NS)
    autodel = root.find("c:chart/c:autoTitleDeleted", NS)
    if t is not None and (autodel is None or autodel.get("val") in ("0", "false", None)):
        texts = [e.text or "" for e in t.findall(".//a:t", NS)]
        txt = "".join(texts).strip()
        if not txt:
            cache = [e.text or ""
                     for e in t.findall(".//c:strCache//c:v", NS)]
            txt = "".join(cache).strip()
        if not txt:
            f = t.find(".//c:strRef/c:f", NS)
            if f is not None and f.text:
                txt = _cell_value_from_ref(book, f.text) \
                    or norm_range_only(f.text)
        title = ("있음", txt)
    legend = root.find("c:chart/c:legend/c:legendPos", NS)
    legend_pos = legend.get("val") if legend is not None else (
        "r" if root.find("c:chart/c:legend", NS) is not None else None)
    ax_values = {}
    for child in plot:
        tag = child.tag.split("}")[1]
        if not tag.endswith("Chart"):
            continue
        ids = [e.get("val") for e in child.findall("c:axId", NS)]
        vals = []
        for ser in child.findall("c:ser", NS):
            got = []
            for path in ("c:val//c:numCache//c:pt/c:v",
                         "c:yVal//c:numCache//c:pt/c:v"):
                for v in ser.findall(path, NS):
                    try:
                        got.append(float(v.text))
                    except (TypeError, ValueError):
                        pass
            if not got:
                for path in ("c:val//c:f", "c:yVal//c:f"):
                    e = ser.find(path, NS)
                    if e is not None and e.text:
                        got = _range_values_from_ref(book, e.text)
                        break
            vals.extend(got)
        for i in ids:
            ax_values.setdefault(i, []).extend(vals)
    axes = []
    axes_detail = []
    for ax in root.findall(".//c:valAx", NS):
        mx = ax.find("c:scaling/c:max", NS)
        mn = ax.find("c:scaling/c:min", NS)
        mju = ax.find("c:majorUnit", NS)
        tup = (
            float(mx.get("val")) if mx is not None else None,
            float(mn.get("val")) if mn is not None else None,
            float(mju.get("val")) if mju is not None else None,
        )
        axes.append(tup)
        axid = ax.find("c:axId", NS)
        pos = ax.find("c:axPos", NS)
        crosses = ax.find("c:crosses", NS)
        deleted = ax.find("c:delete", NS)
        axes_detail.append({
            "max": tup[0], "min": tup[1], "unit": tup[2],
            "id": axid.get("val") if axid is not None else None,
            "pos": pos.get("val") if pos is not None else None,
            "crosses": crosses.get("val") if crosses is not None else None,
            "deleted": deleted is not None and deleted.get("val") in ("1", "true"),
            "values": ax_values.get(axid.get("val") if axid is not None else None,
                                    []),
        })
    return {
        "types": sorted(types),
        "title": title,
        "series": series,
        "legend": legend_pos,
        "axes": axes,
        "axes_detail": axes_detail,
    }

def _range_values_from_ref(book, ref):
    try:
        s = str(ref).strip().replace("$", "")
        if "!" not in s:
            return []
        sheet, rng = s.rsplit("!", 1)
        sheet = sheet.strip("'")
        real = book.norm_map.get(norm_sheet_name(sheet))
        if real is None:
            return []
        out = []
        for (r, c) in expand_ref(rng):
            v = book.cached[real].cell(r, c).value
            if v is None:
                v = book.raw[real].cell(r, c).value
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                out.append(float(v))
        return out
    except Exception:
        return []

AXIS_ROLE_KO = {"primary": "기본 축", "secondary": "보조 축"}
AXIS_ATTR_KO = {"max": "최대", "min": "최소", "unit": "주 단위"}

def _axes_by_role(feat):
    out = {}
    if not feat:
        return out
    for ax in feat.get("axes_detail") or []:
        role = "secondary" if (ax.get("crosses") == "max"
                               or ax.get("pos") in ("r", "t")) else "primary"
        if role in out:
            role = "secondary" if "secondary" not in out else \
                f"axis{len(out) + 1}"
        out[role] = ax
    return out

def _axis_role_name(v):
    s = str(v or "primary").strip().lower()
    if s in ("secondary", "보조", "보조축", "보조 축", "2", "1"):
        return "secondary"
    return "primary"

def _num_txt(v):
    if v is None:
        return "자동"
    try:
        f = float(v)
        return f"{int(f):,}" if f.is_integer() else f"{f:g}"
    except (TypeError, ValueError):
        return str(v)

def _auto_min_zero(values):
    vals = [v for v in values if isinstance(v, (int, float))]
    if not vals:
        return False
    lo, hi = min(vals), max(vals)
    if lo < 0 or hi <= 0:
        return False
    return (hi - lo) > hi / 6.0

def _axis_attr_ok(attr, want, have, ax_s, lenient_unit=False):
    if have is not None and want is not None:
        try:
            return abs(float(have) - float(want)) <= 1e-9 * max(1.0, abs(float(want)))
        except (TypeError, ValueError):
            return False
    if have is None:
        if attr == "unit" and lenient_unit:
            return True
        if attr == "min" and want == 0 and \
                _auto_min_zero((ax_s or {}).get("values") or []):
            return True
    return False

class SheetResult:
    def __init__(self, name, alloc):
        self.name = name
        self.alloc = alloc
        self.earned = 0
        self.details = []
        self.notes = []
        self.missing = False
        self.wrong = []
        self.diff_total = 0
        self.diff_matched = 0
        self.holds = []
        self.held = 0.0

def add_hold(res, label, why, points, what=None):
    pts = round(max(0.0, float(points or 0)), 2)
    res.holds.append({
        "sheet": res.name, "label": label, "why": str(why),
        "points": pts, "what": what,
    })
    res.held = round(res.held + pts, 2)
    return pts

def hold_items(results):
    return [h for r in results for h in getattr(r, "holds", [])]

def hold_total(results):
    return round(sum(getattr(r, "held", 0) for r in results), 2)

def hold_reason(*vals):
    seen = []
    for v in vals:
        if is_unknown(v) and v.why not in seen:
            seen.append(v.why)
    return " / ".join(seen) or "원인을 특정하지 못했습니다"

def _fmt_expected(exp, ef):
    if exp is None and ef:
        return "(수식 기준 채점)"
    return fmt_value(exp)

STATE_BLANK = "(빈 셀 — 값·수식 없음)"
STATE_NOCACHE = "(수식만 있음 · 계산값 없음)"
STATE_RECALC = "(수식만 있음 · 재계산값 {v})"

EXCEL_ERRORS = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!",
                "#N/A", "#SPILL!", "#CALC!", "#GETTING_DATA"}
EXCEL_ERROR_WHY = {
    "#NAME?": "Excel이 모르는 이름이 있을 때 납니다(함수 이름 철자, 따옴표 "
              "빠진 문자, 정의되지 않은 이름을 확인하세요).",
    "#N/A": "찾는 값을 찾지 못했을 때 납니다(찾을 값의 형태가 문자인지 "
            "숫자인지, 찾을 범위와 일치 옵션을 확인하세요).",
    "#VALUE!": "인수의 형태가 맞지 않을 때 납니다(숫자 자리에 문자가 "
               "들어갔는지 확인하세요).",
    "#REF!": "참조하는 셀이 사라졌을 때 납니다(행·열을 지우지 않았는지 "
             "확인하세요).",
    "#DIV/0!": "0으로 나눌 때 납니다(나누는 값이 0이거나 빈 셀입니다).",
    "#NUM!": "계산할 수 없는 숫자일 때 납니다(인수의 범위를 확인하세요).",
    "#NULL!": "범위 연산자를 잘못 썼을 때 납니다(쉼표와 공백을 "
              "확인하세요).",
}

def excel_error(v):
    if isinstance(v, str):
        s = v.strip().upper()
        if s in EXCEL_ERRORS:
            return s
    return None

def excel_error_note(err):
    if not err:
        return None
    why = EXCEL_ERROR_WHY.get(err, "수식이 계산되지 못했습니다.")
    return f"내 답은 값이 아니라 오류 {err} 입니다 — {why}"
NOCACHE_GUIDE = ("수식은 있으나 계산된 값이 파일에 저장되지 않아 수식 문자열로만 "
                 "비교했습니다 — Excel에서 열어 저장(Ctrl+S)한 뒤 다시 채점하면 "
                 "값 기준으로 채점됩니다.")
BLANK_GUIDE = ("채점한 파일에서 이 셀들은 값도 수식도 없는 빈 셀입니다 — 풀이를 "
               "저장(Ctrl+S)한 파일이 맞는지, 다른 사본을 채점한 것은 아닌지 "
               "확인하세요.")
RECALC_GUIDE = ("수식은 있으나 계산된 값이 파일에 저장돼 있지 않아, 이 채점기가 "
                "LibreOffice로 통합 문서를 다시 계산해 결과값으로 비교했습니다.")

def compare_mode_note(judge):
    if getattr(judge, "recalc_cells", None):
        cc = compress_coords(judge.recalc_cells)
        txt = ("계산값이 없어 재계산해 비교했습니다 — LibreOffice로 다시 계산한 "
               "결과값으로 비교한 셀: " + ", ".join(cc[:6])
               + (" 외" if len(cc) > 6 else "") + ".")
        if judge.recalc_failed:
            txt += " (일부 셀은 재계산하지 못해 수식 문자열로 비교했습니다.)"
        return txt
    if getattr(judge, "recalc_failed", False):
        return ("계산값이 없고 재계산도 하지 못해 수식 문자열로 비교했습니다 — "
                "LibreOffice가 설치돼 있지 않거나 변환에 실패했습니다. Excel에서 "
                "파일을 열어 저장(Ctrl+S)한 뒤 다시 채점하면 값 기준으로 "
                "채점됩니다.")
    if getattr(judge, "formula_only_used", False):
        return "계산값이 없어 수식 문자열로 비교했습니다."
    return "저장된 계산값으로 비교했습니다."

def ref_mismatch_note(judge, limit=4):
    items = getattr(judge, "ref_mismatch", None)
    if not items:
        return None
    lines = []
    for coord, ef, sf in items[:limit]:
        more_a, more_b = ref_usage_diff(ef, sf)
        if more_a or more_b:
            lines.append(
                f"{coord}: 정답은 {', '.join(more_a) or '(없음)'}을(를) "
                f"참조하는데 내 답은 {', '.join(more_b) or '(없음)'}을(를) "
                "참조합니다")
        else:
            lines.append(f"{coord}: 참조가 다릅니다")
    tail = f" 외 {len(items) - limit}개" if len(items) > limit else ""
    return ("다시 계산한 결과값은 정답과 같지만, 수식이 참조하는 셀이 정답과 "
            "달라 오답으로 두었습니다 — 이 자료에서 우연히 결과가 겹친 것일 뿐 "
            "참조가 바뀌면 결과가 달라집니다 (" + "; ".join(lines) + tail + "). "
            "실제 시험도 수식의 참조를 보고 채점합니다.")

def student_cell_state(judge, r, c):
    if judge.shs_raw is None:
        return "nosheet", "(시트 없음)", None
    try:
        got, sf = judge.student(r, c)
    except Exception:
        return "value", "?", None
    if got == "__MISSING__":
        rv = judge.recalc_vals_s.get((r, c), "__MISSING__")
        if rv != "__MISSING__":
            erv = excel_error(rv)
            return ("recalc",
                    STATE_RECALC.format(
                        v=(f"{erv} — 오류" if erv else fmt_value(rv))), sf)
        return "nocache", STATE_NOCACHE, sf
    if got is None and not sf:
        return "blank", STATE_BLANK, None
    err = excel_error(got)
    if err:
        return "error", f"{err} (오류 값)", sf
    return "value", fmt_value(got), sf

def make_cell_entries(judge, wrong, limit=3):
    out = []
    for (r, c, coord) in wrong[:limit]:
        ok, exp, got, ef = judge.judge(r, c)
        state, got_txt, sf = student_cell_state(judge, r, c)
        out.append({"coord": coord, "expected": _fmt_expected(exp, ef),
                    "got": got_txt, "formula": ef,
                    "got_formula": sf, "got_state": state})
    return out

def wrong_error_note(judge, wrong):
    for (r, c, _co) in wrong:
        try:
            got, _sf = judge.student(r, c)
        except Exception:
            continue
        err = excel_error(got)
        if err is None:
            rv = judge.recalc_vals_s.get((r, c), None)
            err = excel_error(rv)
        if err:
            return excel_error_note(err)
    return None

def cell_state_guide(judge, wrong):
    states = [student_cell_state(judge, r, c)[0] for (r, c, _co) in wrong]
    if states and all(s == "blank" for s in states):
        return BLANK_GUIDE
    if any(s == "nocache" for s in states):
        return NOCACHE_GUIDE
    if any(s == "recalc" for s in states):
        return RECALC_GUIDE
    return None

def add_card(res, label, lost, kind, cells=None, formula=None,
             note=None, hint=None, more=0, props=None):
    res.wrong.append({
        "sheet": res.name, "label": label, "lost": round(float(lost), 1),
        "kind": kind, "cells": cells or [], "formula": formula,
        "note": note, "hint": hint, "more": more, "props": props or [],
        "category": None, "explain": [], "point": None,
    })

def add_wrong_cells(res, judge, wrong, limit=30):
    for (r, c, coord) in wrong[:limit]:
        ok, exp, got, ef = judge.judge(r, c)
        _state, got_txt, sf = student_cell_state(judge, r, c)
        line = (f"{coord}: 기대값 {_fmt_expected(exp, ef)} / "
                f"학생값 {got_txt}")
        if sf:
            line += f" / 학생 수식 {sf}"
        if ef:
            line += f" / 기대 수식 {ef}"
        res.details.append(line)
    if len(wrong) > limit:
        res.details.append(f"... 외 {len(wrong) - limit}개 셀")

def grade_basic1(res, ctx):
    diffs = ctx["vdiffs"]
    if not diffs:
        res.earned = res.alloc
        res.notes.append("문제/정답 간 차이가 없어 자동 만점 처리")
        return
    judge = ctx["judge"]
    wrong = [(r, c, co) for (r, c, co) in diffs if not judge.judge(r, c)[0]]
    if judge.display_matched:
        res.notes.append(
            "표시 형식 기준으로 같은 값이라 정답 처리: "
            + ", ".join(judge.display_matched[:8])
            + (" 외" if len(judge.display_matched) > 8 else "")
            + " (자료 입력은 문제지에 보이는 대로 입력하면 됩니다 — 예: "
              "57.14%)")
    if wrong:
        res.earned = 0
        res.details.append(f"자료 입력 {len(diffs)}개 셀 중 {len(wrong)}개 불일치 "
                           "(부분점수 없음, 0점)")
        add_wrong_cells(res, judge, wrong)
        guide = cell_state_guide(judge, wrong)
        add_card(res, "자료 입력", res.alloc, "cell",
                 cells=make_cell_entries(judge, wrong, 5),
                 more=max(0, len(wrong) - 5), note=guide,
                 hint="기본작업-1은 부분점수가 없습니다. 입력을 마친 뒤 "
                      "문제지와 셀 단위로 1:1 대조하는 습관을 들이세요. "
                      "숫자는 문제지에 보이는 그대로(예: 57.14%) 입력하면 "
                      "표시 형식 기준으로 같은 값은 정답입니다.")
    else:
        res.earned = res.alloc

def _edge_desc(key, coord):
    from openpyxl.utils.cell import coordinate_to_tuple
    try:
        ek, r, c = key
        tr, tc = coordinate_to_tuple(coord)
        if ek == "h":
            return f"{coord} {'아래쪽' if r > tr else '위쪽'}"
        return f"{coord} {'오른쪽' if c > tc else '왼쪽'}"
    except Exception:
        return str(coord)

def _border_props(fail_units, s_edges):
    styles = set()
    misses = []
    for u in fail_units:
        for key, style in u.get("edges") or []:
            if style:
                styles.add(style)
            sv = s_edges.get(key) if s_edges is not None else None
            if (sv or None) != (style or None):
                if style and not sv:
                    misses.append(f"{_edge_desc(key, u['coord'])} 선 없음")
                elif style and sv:
                    misses.append(f"{_edge_desc(key, u['coord'])} 선 종류 다름"
                                  f"({BORDER_STYLE_KO.get(sv, sv)})")
                else:
                    misses.append(f"{_edge_desc(key, u['coord'])} 불필요한 선")
    if len(styles) == 1:
        exp = f"모든 지시 선({BORDER_STYLE_KO.get(next(iter(styles)), '실선')})"
    elif styles:
        exp = "테두리 적용(" + "·".join(
            BORDER_STYLE_KO.get(s, s) for s in sorted(styles)) + ")"
    else:
        exp = "선 제거"
    got = ", ".join(misses[:3]) + (f" 외 {len(misses) - 3}곳"
                                   if len(misses) > 3 else "")
    return exp, (got or "확인 불가")

def _key_row_heights(ctx, sheet_name):
    out = {}
    fcs = (ctx["key"].get("format_checks") or {}).get(
        norm_sheet_name(sheet_name)) or []
    for ent in fcs:
        chk = (ent or {}).get("check") or {}
        rh = chk.get("row_height", chk.get("row_heights"))
        if rh is None:
            continue
        if isinstance(rh, dict):
            for k, v in rh.items():
                try:
                    out[int(k)] = float(v)
                except (TypeError, ValueError):
                    pass
            continue
        rows = []
        rng = str(ent.get("range") or "").strip()
        m = re.match(r"^(\d+)(?:\s*[:~-]\s*(\d+))?\s*행?$", rng)
        if m:
            rows = range(int(m.group(1)), int(m.group(2) or m.group(1)) + 1)
        else:
            try:
                from openpyxl.utils.cell import range_boundaries
                _c1, r1, _c2, r2 = range_boundaries(rng)
                rows = range(r1, r2 + 1)
            except Exception:
                pass
        for r in rows:
            try:
                out[int(r)] = float(rh)
            except (TypeError, ValueError):
                pass
    return out

def _key_col_widths(ctx, sheet_name):
    out = {}
    fcs = (ctx["key"].get("format_checks") or {}).get(
        norm_sheet_name(sheet_name)) or []
    for ent in fcs:
        chk = (ent or {}).get("check") or {}
        cw = chk.get("col_width", chk.get("col_widths"))
        if cw is None:
            continue
        if isinstance(cw, dict):
            for k, v in cw.items():
                try:
                    out[str(k).strip().upper().rstrip("열")] = float(v)
                except (TypeError, ValueError):
                    pass
            continue
        cols = []
        rng = str(ent.get("range") or "").strip().upper()
        m = re.match(r"^([A-Z]{1,3})(?:\s*[:~-]\s*([A-Z]{1,3}))?\s*열?$", rng)
        if m:
            from openpyxl.utils import column_index_from_string, \
                get_column_letter
            c1 = column_index_from_string(m.group(1))
            c2 = column_index_from_string(m.group(2) or m.group(1))
            cols = [get_column_letter(i) for i in range(c1, c2 + 1)]
        for cl in cols:
            try:
                out[cl] = float(cw)
            except (TypeError, ValueError):
                pass
    return out

def grade_basic2(res, ctx):
    from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries
    book_a, book_s = ctx["book_a"], ctx["book_s"]
    asheet, ssheet = ctx["asheet"], ctx["ssheet"]
    judge = ctx["judge"]
    fd = ctx["fdiffs"]
    sha = book_a.raw[asheet]
    shs = book_s.raw[ssheet] if ssheet else None
    a_members, a_anchors = merge_maps(sha)
    s_members, s_anchors = merge_maps(shs) if shs is not None else ({}, {})
    if any_unknown(a_members, s_members):
        add_hold(res, "기본작업-2 서식",
                 hold_reason(a_members, s_members)
                 + " — 병합 기준을 못 잡아 서식을 대조하지 못했습니다.",
                 res.alloc, what=res.name)
        res.earned = 0
        res.details.append("서식을 읽지 못해 판정을 보류했습니다 (확인 필요)")
        return

    units = []
    for kind in CELL_FMT_KINDS:
        for coord in fd.get(kind, []):
            r, c = coordinate_to_tuple(coord)
            units.append({"pos": (r, c), "kind": kind, "coord": coord})
    for coord in fd.get("border", []):
        r, c = coordinate_to_tuple(coord)
        units.append({"pos": (r, c), "kind": "border", "coord": coord,
                      "edges": (fd.get("border_edges") or {}).get(coord, [])})
    for rng in fd.get("merge", []):
        try:
            min_c, min_r, _mc, _mr = range_boundaries(rng)
        except Exception:
            continue
        units.append({"pos": (min_r, min_c), "kind": "merge",
                      "range": rng, "coord": rng})
    for rng in fd.get("merge_del", []):
        try:
            min_c, min_r, _mc, _mr = range_boundaries(rng)
        except Exception:
            continue
        units.append({"pos": (min_r, min_c), "kind": "merge_del",
                      "range": rng, "coord": rng})
    for (r, c, coord) in ctx["vdiffs"]:
        units.append({"pos": (r, c), "kind": "value", "coord": coord,
                      "rc": (r, c)})

    standalone = []
    for (where, kind, why) in (ctx.get("fmt_unknown") or []):
        standalone.append(("_unknown", (where, kind, why)))
    if fd.get("names"):
        standalone.append(("names", fd["names"]))
    key_rh = _key_row_heights(ctx, res.name)
    if key_rh:
        standalone.append(("key_rowheight", key_rh))
    key_cw = _key_col_widths(ctx, res.name)
    if key_cw:
        standalone.append(("key_colwidth", key_cw))

    if not key_rh and fd.get("rowheight") and shs is not None:
        noted = []
        for (r, ha) in fd["rowheight"]:
            hs = shs.row_dimensions[r].height \
                if r in shs.row_dimensions else None
            if ha is not None and (hs is None or abs(hs - ha) > 1.0):
                noted.append(str(r))
        if noted:
            res.notes.append(
                f"행 높이가 정답 파일과 다릅니다({', '.join(noted[:6])}행) — "
                "글꼴 크기 등에 따른 자동 조정 차이일 수 있어 감점하지 "
                "않습니다. 문제지에 행 높이 지시가 있는 경우만 확인하세요.")
    if not key_cw and fd.get("colwidth") and shs is not None:
        noted = []
        for (cl, wa) in fd["colwidth"]:
            ws_ = shs.column_dimensions[cl].width \
                if cl in shs.column_dimensions else None
            if wa is not None and not col_width_same(ws_, wa, tol=1.0):
                noted.append(str(cl))
        if noted:
            res.notes.append(
                f"열 너비가 정답 파일과 다릅니다({', '.join(noted[:6])}열) — "
                "감점하지 않습니다. 문제지에 열 너비 지시가 있는 경우만 "
                "확인하세요.")

    if not units and not standalone:
        res.earned = res.alloc
        res.notes.append("채점할 서식 차이가 없어 만점 처리")
        return

    clusters = _cluster_units(units) if units else []
    n = len(clusters) + len(standalone)
    per = res.alloc / n
    passed = 0
    mr, mc = scan_bounds(sha, shs) if shs is not None else scan_bounds(sha)
    s_edges = sheet_edge_map(shs, mr, mc) if shs is not None else {}
    edges_blind = hold_reason(s_edges) if is_unknown(s_edges) else None
    if edges_blind:
        s_edges = {}
    s_merges = {str(x).upper() for x in shs.merged_cells.ranges} \
        if shs is not None else set()

    nf_same = []
    held_clusters = 0
    for cl in clusters:
        fails = []
        blind_units = []
        for u in cl:
            k = u["kind"]
            if k == "merge":
                if u["range"].upper() not in s_merges:
                    fails.append(u)
            elif k == "merge_del":
                if u["range"].upper() in s_merges:
                    fails.append(u)
            elif k == "value":
                r, c = u["rc"]
                if not judge.judge(r, c)[0]:
                    fails.append(u)
            elif k == "border":
                if edges_blind:
                    u["blind"] = edges_blind
                    blind_units.append(u)
                    continue
                bad = any((s_edges.get(key) or None) != (style or None)
                          for key, style in (u.get("edges") or []))
                if shs is None or bad:
                    fails.append(u)
            else:
                r, c = u["pos"]
                sig_a = fmt_signature_m(sha, r, c, k, a_members, a_anchors)
                sig_s = fmt_signature_m(shs, r, c, k, s_members, s_anchors) \
                    if shs is not None else None
                if any_unknown(sig_a, sig_s):
                    u["blind"] = hold_reason(sig_a, sig_s)
                    blind_units.append(u)
                    continue
                cmp_a, cmp_s = sig_a, sig_s
                if k == "alignment" and _blank_answer_cell(sha, r, c):
                    cmp_a, cmp_s = _blank_align_sig(sig_a), _blank_align_sig(sig_s)
                if k == "number_format":
                    u["raw_a"] = cell_raw_m(sha, r, c, a_members,
                                            "number_format")
                    u["raw_s"] = cell_raw_m(shs, r, c, s_members,
                                            "number_format")
                    if cmp_a != cmp_s and nf_display_equal(
                            sha, shs, r, c, a_members, s_members):
                        nf_same.append(u["coord"])
                        continue
                if cmp_a != cmp_s:
                    u["sig_a"], u["sig_s"] = sig_a, sig_s
                    fails.append(u)
        if blind_units and not fails:
            held_clusters += 1
            locs_b = compress_coords([u["coord"] for u in blind_units])
            add_hold(res, "서식 지시 (" + ", ".join(locs_b[:3]) + ")",
                     hold_reason(*[Unknown(u["blind"]) for u in blind_units]),
                     per, what=", ".join(locs_b[:3]))
            continue
        if not fails:
            passed += 1
            continue
        loc = compress_coords([u["coord"] for u in cl])

        def _kind_loc(k):
            locs = compress_coords([x["coord"] for x in fails
                                    if x["kind"] == k])
            return ", ".join(locs[:2]) + (" 외" if len(locs) > 2 else "")

        props = []
        hints = []
        seen_kinds = []
        for u in fails:
            k = u["kind"]
            if k in seen_kinds and k != "value":
                continue
            seen_kinds.append(k)
            if k == "merge":
                rngs = [x["range"] for x in fails if x["kind"] == "merge"]
                props.append({"name": "병합",
                              "expected": "병합하고 가운데 맞춤 ("
                                          + ", ".join(rngs[:3]) + ")",
                              "got": "병합 없음"})
                hints.extend(hints_for_kind("merge", None))
            elif k == "merge_del":
                rngs = [x["range"] for x in fails
                        if x["kind"] == "merge_del"]
                props.append({"name": "병합 해제",
                              "expected": "병합 해제 ("
                                          + ", ".join(rngs[:3]) + ")",
                              "got": "병합이 남아 있음"})
                hints.append("병합된 범위를 선택하고 홈 탭 → '병합하고 "
                             "가운데 맞춤'을 다시 눌러 해제합니다.")
            elif k == "value":
                if sum(1 for p in props if p["name"].startswith("셀 값")) < 2:
                    r, c = u["rc"]
                    _ok, exp, got, _ef = judge.judge(r, c)
                    props.append({"name": f"셀 값({u['coord']})",
                                  "expected": fmt_value(exp),
                                  "got": fmt_value(got)})
                    hints.extend(hints_for_kind("value", None))
            elif k == "border":
                b_fails = [x for x in fails if x["kind"] == "border"]
                exp, got = _border_props(b_fails, s_edges)
                props.append({"name": f"테두리({_kind_loc(k)})",
                              "expected": exp, "got": got})
                hints.extend(hints_for_kind("border", None))
            else:
                where = _kind_loc(k)
                for p in props_for_kind(k, u.get("sig_a"), u.get("sig_s"),
                                        u.get("raw_a"), u.get("raw_s")):
                    p["name"] = f"{p['name']}({where})"
                    props.append(p)
                hints.extend(hints_for_kind(k, u.get("sig_a"), u.get("raw_a")))
        uniq_hints = []
        for h in hints:
            if h not in uniq_hints:
                uniq_hints.append(h)
        label = "서식 지시 (" + ", ".join(loc[:2]) \
            + (" 외" if len(loc) > 2 else "") + ")"
        add_card(res, label, per, "format",
                 cells=[{"coord": t} for t in loc[:6]],
                 props=props, hint=" / ".join(uniq_hints[:3]) or None,
                 more=max(0, len(loc) - 6))
        res.details.append(
            "[서식] " + ", ".join(loc[:4]) + " 불일치: "
            + ", ".join(dict.fromkeys(p["name"] for p in props)))

    if nf_same:
        locs = ", ".join(compress_coords(nf_same)[:4])
        res.notes.append(
            f"표시 형식 코드가 정답 파일과 다르지만({locs}) 그 값의 화면 표시가 "
            "같아 정답 처리했습니다. 실제 시험도 표시 예(예: 460 → 460개)와 "
            "같게 보이면 정답으로 봅니다.")

    for kind, payload in standalone:
        if kind == "_unknown":
            where, ukind, why = payload
            add_hold(res, f"서식 지시 추출 ({where} "
                          f"{FMT_KIND_LABEL.get(ukind, ukind)})",
                     why + " — 문제·정답 파일에서 이 지시를 읽어 내지 못해 "
                           "채점 대상에서 뺐습니다.",
                     per, what=where)
            held_clusters += 1
            continue
        if kind == "names":
            stu_names = defined_names_for_sheet(book_s, ssheet) \
                if ssheet else {}
            if is_unknown(stu_names) or is_unknown(payload):
                add_hold(res, "서식 - 정의된 이름",
                         hold_reason(stu_names, payload)
                         + " — 이름 정의를 대조하지 못해 이 지시는 점수에서 "
                           "뺐습니다.",
                         per, what="정의된 이름")
                held_clusters += 1
                continue
            missing = [k for k, v in payload.items()
                       if stu_names.get(k) != v]
            if not missing:
                passed += 1
            else:
                props = []
                notes_nm = []
                for nm in missing[:5]:
                    exp_ref = str(payload[nm])
                    got_ref = stu_names.get(nm) or _book_defined_name(book_s, nm)
                    got_txt = f"'{nm}' = {got_ref}" if got_ref else \
                        f"'{nm}' 이름 없음"
                    props.append({"name": "정의된 이름",
                                  "expected": f"'{nm}' = {exp_ref}",
                                  "got": got_txt})
                    ea, ga = _ref_areas(exp_ref), _ref_areas(got_ref)
                    if got_ref and len(ea) > 1 and len(ga) < len(ea):
                        notes_nm.append(
                            f"'{nm}'의 정답 참조는 영역 {len(ea)}개"
                            f"({', '.join(ea)})를 쉼표로 이은 것인데 내 답은 "
                            f"{len(ga)}개({', '.join(ga) or '없음'})뿐입니다 — "
                            "Ctrl을 누른 채 영역을 모두 선택한 뒤 이름을 "
                            "정의해야 합니다.")
                    elif got_ref and ea != ga:
                        notes_nm.append(
                            f"'{nm}'의 참조 범위가 다릅니다: 정답 "
                            f"{', '.join(ea)} / 내 답 {', '.join(ga)}.")
                res.details.append("[정의된 이름] 불일치 (이름 "
                                   + ", ".join(missing[:5]) + "): "
                                   + "; ".join(p["got"] for p in props))
                add_card(res, "서식 - 정의된 이름", per, "format",
                         cells=[{"coord": nm} for nm in missing[:5]],
                         props=props,
                         hint="범위를 선택한 뒤 이름 상자(수식 입력줄 왼쪽)에 "
                              "이름을 입력하고 Enter를 누릅니다. 떨어진 영역 "
                              "두 개를 한 이름으로 정의하려면 첫 영역을 드래그한 "
                              "뒤 Ctrl을 누른 채 둘째 영역을 드래그하고 이름 "
                              "상자에 입력합니다 (수식 탭 → 이름 정의 → 참조 "
                              "대상에 =시트!$A$4:$A$10,시트!$C$4:$C$10 도 가능).")
                if notes_nm:
                    res.wrong[-1]["diff_notes"] = notes_nm
        elif kind == "key_rowheight":
            bad = []
            for r, want in sorted(payload.items()):
                hs = shs.row_dimensions[r].height \
                    if shs is not None and r in shs.row_dimensions else None
                if hs is None or abs(hs - want) > ROW_HEIGHT_TOL:
                    bad.append((r, want, hs))
            if not bad:
                passed += 1
            else:
                rows_txt = ", ".join(
                    f"{r}행: 정답 {w:g} / 내 답 {row_height_text(h)}"
                    for r, w, h in bad)
                res.details.append(f"[행 높이] 불일치 ({rows_txt})")
                add_card(res, "서식 - 행 높이", per, "format",
                         cells=[{"coord": f"{r}행"} for r, _w, _h in bad],
                         props=[{"name": f"{r}행 높이",
                                 "expected": f"{w:g}",
                                 "got": row_height_text(h)}
                                for r, w, h in bad[:3]],
                         hint="행 머리글 우클릭 → [행 높이]에 숫자를 그대로 "
                              "입력합니다. 행 높이는 포인트(pt) 단위라 입력한 "
                              "숫자가 그대로 저장됩니다 — 경계선을 끌어서 "
                              "맞추면 값이 어긋나니 반드시 숫자를 입력하세요.")
        elif kind == "key_colwidth":
            bad = []
            for cl, want in sorted(payload.items()):
                cw_s = shs.column_dimensions[cl].width \
                    if shs is not None and cl in shs.column_dimensions \
                    else None
                if not col_width_is(cw_s, want):
                    bad.append((cl, want, cw_s))
            if not bad:
                passed += 1
            else:
                cols_txt = ", ".join(
                    f"{cl}열: 정답 {w:g}자 / 내 답 {col_width_text(s, w)}"
                    for cl, w, s in bad)
                res.details.append(f"[열 너비] 불일치 ({cols_txt})")
                add_card(res, "서식 - 열 너비", per, "format",
                         cells=[{"coord": f"{cl}열"} for cl, _w, _s in bad],
                         props=[{"name": f"{cl}열 너비",
                                 "expected": f"{w:g}자",
                                 "got": col_width_text(s, w)}
                                for cl, w, s in bad[:3]],
                         note="열 너비는 Excel이 대화상자에 입력한 문자 수에 "
                              "좌우 여백 5픽셀을 더해 저장합니다 (맑은 고딕 "
                              "11pt 기준 12 입력 → 12.625 저장). 이 채점기는 "
                              "저장값을 다시 문자 수로 되돌려 비교하므로, "
                              "대화상자에 적힌 숫자만 정답과 같으면 됩니다.",
                         hint="열 머리글 우클릭 → [열 너비]에 숫자를 그대로 "
                              "입력합니다. 경계선을 끌어서 맞추면 소수점 값이 "
                              "들어가 어긋나니 반드시 숫자를 입력하세요.")

    res.earned = int(round(per * passed))
    if held_clusters:
        res.earned = min(res.earned, int(round(per * (n - held_clusters))))
    if passed < n - held_clusters:
        res.details.insert(0, f"서식 지시 {n}개 중 {passed}개 통과 "
                              f"(항목당 {per:.1f}점"
                              + (f", 확인 필요 {held_clusters}개 제외"
                                 if held_clusters else "") + ")")

def _is_cond_syntax(v):
    return isinstance(v, str) and bool(
        re.match(r"^\s*(<>|<=|>=|<|>|=)", v) or "*" in v or "?" in v)

def _norm_field(v):
    return re.sub(r"\s+", "", str(v)).lower() if v is not None else None

def _norm_cond(v):
    if v is None:
        return None
    if isinstance(v, str):
        s = v.strip()
        if s.startswith("="):
            return ("f", norm_formula(s))
        return ("s", re.sub(r"\s+", "", s).lower())
    if isinstance(v, bool):
        return ("b", v)
    if isinstance(v, (int, float)):
        return ("n", round(float(v), 6))
    return ("s", str(v))

def _block_grid(sh, r1, r2, c1, c2):
    if sh is None:
        return []
    return [[sh.cell(r, c).value for c in range(c1, c2 + 1)]
            for r in range(r1, r2 + 1)]

def _cond_rows(grid):
    from collections import Counter
    header = [_norm_field(v) for v in grid[0]]
    rows = []
    for row in grid[1:]:
        pairs = set()
        for i, v in enumerate(row):
            if v is None:
                continue
            cond = _norm_cond(v)
            if isinstance(v, str) and v.strip().startswith("="):
                pairs.add(("#수식#", cond))
            elif header[i]:
                pairs.add((header[i], cond))
        if pairs:
            rows.append(frozenset(pairs))
    return Counter(rows), header

def _excel_cond_match_norm(cond, value):
    if cond is None:
        return True
    kind, cv = cond
    if kind == "f":
        return None
    if kind == "n":
        vn = _as_number(value)
        return vn is not None and abs(vn - cv) < 1e-9
    if kind == "b":
        return isinstance(value, bool) and value == cv
    s = str(cv)
    vtxt = re.sub(r"\s+", "", str(value)).lower() if value is not None else ""
    m = re.match(r"^(<>|<=|>=|<|>|=)(.*)$", s)
    if m:
        op, rhs = m.groups()
        rn, vn = _as_number(rhs), _as_number(value)
        if rn is not None and vn is not None:
            return {"<>": vn != rn, "<=": vn <= rn, ">=": vn >= rn,
                    "<": vn < rn, ">": vn > rn,
                    "=": abs(vn - rn) < 1e-9}[op]
        if op == "=":
            return vtxt == rhs
        if op == "<>":
            return vtxt != rhs
        return {"<=": vtxt <= rhs, ">=": vtxt >= rhs,
                "<": vtxt < rhs, ">": vtxt > rhs}[op]
    if "*" in s or "?" in s:
        pat = "^" + re.escape(s).replace(r"\*", ".*").replace(r"\?", ".") + "$"
        return re.match(pat, vtxt) is not None
    return vtxt.startswith(s)

def _locate_source_table(shp, fields):
    mr, mc = scan_bounds(shp)
    for r in range(1, min(mr, 30) + 1):
        hdr = {}
        for c in range(1, mc + 1):
            v = shp.cell(r, c).value
            if isinstance(v, str) and v.strip():
                hdr.setdefault(_norm_field(v), c)
        if fields and fields <= set(hdr):
            rows = []
            rr = r + 1
            while rr <= mr:
                vals = {f: shp.cell(rr, col).value for f, col in hdr.items()}
                if all(v is None for v in vals.values()):
                    break
                rows.append(vals)
                rr += 1
            return hdr, rows
    return None, None

def filter_result_equivalent(shp, ga, gs):
    try:
        ca, _ha = _cond_rows(ga)
        cs, _hs = _cond_rows(gs)
        if not ca or not cs:
            return None
        pairs = [p for rows in list(ca) + list(cs) for p in rows]
        fields = set()
        for f, cond in pairs:
            if f == "#수식#" or (isinstance(cond, tuple) and cond[0] == "f"):
                return None
            fields.add(f)
        hdr, rows = _locate_source_table(shp, fields)
        if hdr is None or not rows:
            return None

        def match_set(counter):
            out = set()
            for i, vals in enumerate(rows):
                for rowset in counter:
                    ok = True
                    for f, cond in rowset:
                        m = _excel_cond_match_norm(cond, vals.get(f))
                        if m is None:
                            raise ValueError("판정 불가 조건")
                        if not m:
                            ok = False
                            break
                    if ok:
                        out.add(i)
                        break
            return out

        return match_set(ca) == match_set(cs)
    except Exception:
        return None

COND_EQUIV_NOTE = ("필터 결과는 정답과 같지만, 조건식이 문제 지시와 다릅니다 "
                   "(예: '~이 아니면서'는 <>값). 실제 시험은 조건 셀에 입력한 "
                   "내용도 채점할 수 있어 오답 처리했습니다 — '~이 아닌' "
                   "조건은 항상 <> 로 쓰세요.")

def compare_condition_block(ga, gs):
    ca, _ha = _cond_rows(ga)
    cs, _hs = _cond_rows(gs) if gs else (None, None)
    if cs is not None and ca == cs:
        return True, []
    msgs = []
    flat_s = {}
    if cs:
        for rowset in cs:
            for (f, cond) in rowset:
                flat_s.setdefault(f, []).append(cond)

    def _fmt(cond):
        if cond is None:
            return "없음"
        return str(cond[1])

    missing_rows = ca - cs if cs is not None else ca
    for rowset in missing_rows:
        for (f, cond) in sorted(rowset, key=str):
            got = flat_s.get(f)
            if got and cond not in got:
                msgs.append(f"'{f}' 조건이 {_fmt(cond)} 이어야 하는데 "
                            f"{_fmt(got[0])} 입니다")
            elif not got:
                msgs.append(f"'{f}' 조건({_fmt(cond)})이 없습니다 — 필드명 "
                            "오타 여부도 확인하세요")
    if not msgs:
        msgs.append("조건의 AND(같은 행)/OR(다른 행) 배치가 정답과 다릅니다")
    return False, msgs[:4]

def compare_result_block(ga, gs):
    from collections import Counter
    ha = [_norm_field(v) for v in ga[0]]
    hs = [_norm_field(v) for v in gs[0]] if gs else []
    ha_set = sorted(h for h in ha if h)
    hs_set = sorted(h for h in hs if h)
    if ha_set != hs_set:
        return False, ["추출된 필드(머리글)가 정답과 다릅니다: 정답 "
                       + ", ".join(ha_set) + " / 내 답 "
                       + (", ".join(hs_set) or "없음")], False
    idx_s = {}
    for i, h in enumerate(hs):
        if h and h not in idx_s:
            idx_s[h] = i
    cols_a = [i for i, h in enumerate(ha) if h]

    def _rows(grid, mapper):
        out = []
        for row in grid[1:]:
            t = tuple(mapper(row, i) for i in cols_a)
            if any(x is not None for x in t):
                out.append(t)
        return out

    rows_a = _rows(ga, lambda row, i: _norm_cond(row[i]))
    rows_s = _rows(gs, lambda row, i: _norm_cond(row[idx_s[ha[i]]]))
    if Counter(rows_a) == Counter(rows_s):
        return True, [], rows_a != rows_s
    only_a = Counter(rows_a) - Counter(rows_s)
    only_s = Counter(rows_s) - Counter(rows_a)
    msgs = []
    if only_a:
        msgs.append(f"정답에 있는 결과 행 {sum(only_a.values())}개가 "
                    "누락되었습니다")
    if only_s:
        msgs.append(f"정답에 없는 행 {sum(only_s.values())}개가 "
                    "포함되었습니다 (조건 재확인)")
    return False, msgs or ["결과 행이 정답과 다릅니다"], False

def _cf_boundary_note(rule, book_s, ssheet):
    try:
        if rule.get("type") != "cellIs" or not rule.get("formulas") or not ssheet:
            return ""
        thr = float(str(rule["formulas"][0]).rstrip("%")) / (
            100.0 if str(rule["formulas"][0]).endswith("%") else 1.0)
        ws = book_s.cached[ssheet]
        hits = []
        for rng in rule.get("ranges") or []:
            for (r, c) in expand_ref(rng):
                v = ws.cell(r, c).value
                if isinstance(v, (int, float)) and not isinstance(v, bool) \
                        and abs(float(v) - thr) < 1e-9:
                    hits.append(ws.cell(r, c).coordinate)
        if hits:
            return (f" 이 표에는 경계값 {thr:g}과 같은 셀({', '.join(hits[:4])})이 "
                    "있어 화면 결과도 달라집니다.")
        return (f" 이 표에는 경계값 {thr:g}과 정확히 같은 셀이 없어 화면 결과는 "
                "같아 보이지만, 실제 시험은 규칙 자체를 채점하므로 오답입니다.")
    except Exception:
        return ""

def grade_basic3(res, ctx):
    book_p, book_a, book_s = ctx["book_p"], ctx["book_a"], ctx["book_s"]
    psheet, asheet, ssheet = ctx["psheet"], ctx["asheet"], ctx["ssheet"]
    rules_a = cf_rules_detail(book_a.raw[asheet], book_a.raw)
    rules_p = cf_rules_detail(book_p.raw[psheet], book_p.raw)
    if any_unknown(rules_a, rules_p):
        add_hold(res, "기본작업-3 조건부 서식",
                 hold_reason(rules_a, rules_p)
                 + " — 정답/문제 파일의 규칙을 읽지 못해 이 문항은 점수에서 "
                   "뺐습니다.",
                 res.alloc, what=res.name)
        res.earned = 0
        res.details.append("조건부 서식 규칙을 읽지 못해 판정을 보류했습니다 "
                           "(확인 필요)")
        return
    p_keys = {r["key"] for r in rules_p}
    target = [r for r in rules_a if r["key"] not in p_keys]
    if target:
        raw_s = cf_rules_detail(book_s.raw[ssheet], book_s.raw) if ssheet \
            else []
        if is_unknown(raw_s):
            add_hold(res, "기본작업-3 조건부 서식",
                     hold_reason(raw_s)
                     + " — 내 파일의 규칙을 읽지 못해 이 문항은 점수에서 "
                       "뺐습니다.",
                     res.alloc, what=res.name)
            res.earned = 0
            res.details.append("내 파일의 조건부 서식 규칙을 읽지 못해 판정을 "
                               "보류했습니다 (확인 필요)")
            return
        rules_s = [r for r in raw_s if r["key"] not in p_keys]
        map_s, _over = cf_effective_map(rules_s)
        matched = 0
        from openpyxl.utils import get_column_letter

        def _rng_txt(cells):
            rs = [r for (r, _c) in cells]
            cs = [c for (_r, c) in cells]
            if cells and len(set(cells)) == \
                    (max(rs) - min(rs) + 1) * (max(cs) - min(cs) + 1):
                a = f"{get_column_letter(min(cs))}{min(rs)}"
                b = f"{get_column_letter(max(cs))}{max(rs)}"
                return a if a == b else f"{a}:{b}"
            co = compress_coords([f"{get_column_letter(c)}{r}" for (r, c) in cells])
            return ", ".join(co[:4]) + (" 외" if len(co) > 4 else "")

        for rule in target:
            v = cf_compare_rule(rule, rules_s, map_s)
            if v["ok"]:
                matched += 1
                continue
            sq, typ = rule["sqref"], rule["type"]
            typ_ko = CF_TYPE_KO.get(typ, typ)
            exp_f = ("=" + rule["formulas"][0]) if rule["formulas"] else None
            stu_f = next((("=" + s["formulas"][0]) for s in rules_s
                          if s["formulas"]), None)
            exp_c = _cf_cond_text(rule)
            same_thr = [s for s in rules_s if s["key"][3] == rule["key"][3]]
            stu_c = " / ".join(_cf_cond_text(s) for s in (same_thr or rules_s)) \
                or "(규칙 없음)"
            is_cellis = typ == "cellIs"
            diffs = []
            if not rules_s:
                diffs.append("내 시트에 조건부 서식 규칙이 없습니다 (문제 파일에 "
                             "원래 있던 규칙 제외). 규칙을 만든 뒤 저장했는지 "
                             "확인하세요.")
            if v["fallback"]:
                diffs.append("규칙 범위가 너무 커서 범위·유형·수식 문자열로만 "
                             "비교했습니다.")
            if v["missing"] and rules_s:
                diffs.append(
                    f"적용 범위 불일치: 정답은 {sq} 전체에 규칙이 있는데 내 답은 "
                    f"{_rng_txt(v['missing'])}에 규칙이 없습니다 — '행 전체' "
                    "서식은 표 전체 범위를 선택한 뒤 규칙을 만들어야 합니다.")
            if v["type_bad"]:
                got_types = sorted({CF_TYPE_KO.get(s["type"], s["type"])
                                    for s in rules_s})
                diffs.append(f"규칙 유형 불일치: 정답 '{typ_ko}' / 내 답 "
                             f"'{', '.join(got_types)}'")
            op_got = None
            if v["operator_bad"]:
                ops = [o for o in (v.get("operators") or []) if o]
                op_got = ", ".join(_cf_op_text(o) for o in ops) or "(없음)"
                diffs.append(
                    f"비교 연산자 불일치: 정답 {_cf_op_text(rule['operator'])} / "
                    f"내 답 {op_got} — '이상/이하'는 그 값을 포함하므로 >= / <=, "
                    f"'초과/미만'은 포함하지 않으므로 > / < 입니다."
                    + _cf_boundary_note(rule, book_s, ssheet))
            if v["formula_bad"]:
                anchoring = stu_f is not None and exp_f is not None and \
                    norm_formula(stu_f) == norm_formula(exp_f)
                if anchoring:
                    diffs.append(
                        f"수식의 참조 고정($) 방식이 다릅니다: 내 답 {stu_f} → "
                        f"정답 {exp_f}. 행마다 기준 열을 검사하려면 열만 고정한 "
                        "혼합참조($A4)여야 합니다 — A4(상대)는 B열에서 B4를 "
                        "검사하고, $A$4(절대)는 모든 행이 4행만 검사합니다.")
                    if cmp_order_swapped(stu_f, exp_f):
                        diffs.append(
                            "부등호의 좌우 순서만 바꾼 식(A>=B ↔ B<=A)은 같은 "
                            "조건이므로 정답으로 인정했습니다 — 감점 사유는 "
                            "참조 고정($)뿐입니다.")
                elif is_cellis:
                    diffs.append(f"조건 불일치: 정답 {exp_c} / 내 답 {stu_c} "
                                 f"({_rng_txt(v['formula_bad'])})")
                else:
                    diffs.append(f"수식 불일치: 정답 {exp_f} / 내 답 "
                                 f"{stu_f or '(수식 없음)'} "
                                 f"({_rng_txt(v['formula_bad'])})")
            if v["dxf_bad"]:
                got_d = " / ".join(_dxf_desc(s["dxf"]) for s in rules_s)
                diffs.append(f"적용 서식 불일치: 정답 {_dxf_desc(rule['dxf'])} "
                             f"/ 내 답 {got_d}")
            stu_txt = " / ".join(_cf_rule_text(s) for s in rules_s) or "(규칙 없음)"
            res.details.append(
                f"[조건부 서식] 누락/불일치: 범위 {sq}, 유형 {typ_ko}, 조건 {exp_c}"
                + f" / 내 규칙: {stu_txt}"
                + (" — " + " ".join(diffs) if diffs else ""))
            props = [
                {"name": "적용 범위", "expected": sq,
                 "got": " / ".join(s["sqref"] for s in rules_s) or "(규칙 없음)"},
                {"name": "조건" if is_cellis else "수식", "expected": exp_c,
                 "got": stu_c},
                {"name": "적용 서식", "expected": _dxf_desc(rule["dxf"]),
                 "got": " / ".join(_dxf_desc(s["dxf"]) for s in rules_s)
                 or "(규칙 없음)"},
            ]
            if v["type_bad"]:
                props.append({"name": "규칙 유형", "expected": typ_ko,
                              "got": ", ".join(sorted(
                                  {CF_TYPE_KO.get(s["type"], s["type"])
                                   for s in rules_s})) or "(규칙 없음)"})
            if v["operator_bad"]:
                props.append({"name": "비교 연산자",
                              "expected": _cf_op_text(rule["operator"]),
                              "got": op_got})
            if is_cellis:
                hint = ("홈 탭 → 조건부 서식 → 새 규칙 → '다음을 포함하는 셀만 "
                        "서식 지정' → '셀 값' / 연산자(이상 = >=, 초과 = >, "
                        "이하 = <=, 미만 = <) / 값 입력 → [서식]에서 채우기·"
                        "글꼴 지정. 규칙 수만큼 새 규칙을 반복합니다.")
                point = ("'70% 이상'은 70%를 포함하므로 '>=' (이상), '70% 초과'는 "
                         "포함하지 않으므로 '>' 입니다. 실제 시험은 화면 결과가 "
                         "아니라 규칙(연산자·값·서식) 자체를 채점합니다.")
            else:
                hint = ("홈 탭 → 조건부 서식 → 새 규칙 → '수식을 사용하여 "
                        "서식을 지정할 셀 결정'. 포인트: 행 전체에 적용하려면 "
                        "표 전체 범위를 먼저 선택하고, 수식은 열만 고정하는 "
                        "혼합참조($A4)를 쓰며, 텍스트 비교는 큰따옴표"
                        "(\"...\")로 감쌉니다.")
                point = cf_row_principle(rule, book_a.raw[asheet])
            add_card(res, "조건부 서식", res.alloc / len(target), "cf",
                     formula=None if is_cellis else exp_f, props=props,
                     note=f"정답 규칙: 적용 범위 {sq} / 유형 {typ_ko}"
                          + f" / 조건 {exp_c}"
                          + f" / {_dxf_desc(rule['dxf'])} ‖ 내 규칙: {stu_txt}",
                     hint=hint)
            card = res.wrong[-1]
            card["diff_notes"] = diffs
            card["student_formula"] = stu_f
            card["point"] = point
            if is_cellis:
                card["explain"] = [
                    f"① {sq} 범위를 선택하고 홈 탭 → 조건부 서식 → 새 규칙",
                    "② 규칙 유형 '다음을 포함하는 셀만 서식 지정' → '셀 값' 선택",
                    f"③ 조건: {exp_c} "
                    f"(연산자 {_cf_op_text(rule['operator'])})",
                    f"④ [서식] 단추 → {_dxf_desc(rule['dxf'])} → 확인",
                ]
        res.earned = int(round(res.alloc * matched / len(target)))
        return
    diffs = ctx["vdiffs"]
    if not diffs:
        res.earned = res.alloc
        res.notes.append("문제/정답 간 차이가 없어 자동 만점 처리")
        return
    judge = ctx["judge"]
    from openpyxl.utils import get_column_letter
    shp = book_p.raw[psheet]
    sha = book_a.raw[asheet]
    shs = book_s.raw[ssheet] if ssheet else None
    coord_of = {(r, c): co for (r, c, co) in diffs}
    clusters = cluster_cells([(r, c) for (r, c, _co) in diffs], target=99)
    header_pool = set()
    try:
        for row in shp.iter_rows(min_row=1,
                                 max_row=min(40, shp.max_row or 1)):
            for cell in row:
                if isinstance(cell.value, str) and cell.value.strip():
                    header_pool.add(_norm_field(cell.value))
    except Exception:
        pass
    blocks = []
    for cl in clusters:
        r1 = min(r for r, _c in cl)
        r2 = max(r for r, _c in cl)
        c1 = min(c for _r, c in cl)
        c2 = max(c for _r, c in cl)
        ga = _block_grid(sha, r1, r2, c1, c2)
        first = [v for v in ga[0] if v is not None] if ga else []
        headerish = bool(first) and all(
            isinstance(v, str) and _norm_field(v) in header_pool
            for v in first)
        cond_syntax = any(_is_cond_syntax(v) for row in ga[1:] for v in row)
        has_formula_cond = any(
            isinstance(v, str) and v.strip().startswith("=")
            for row in ga[1:] for v in row)
        if len(ga) >= 2 and (has_formula_cond or (
                headerish and (cond_syntax or len(ga) == 2))):
            btype = "cond"
        elif headerish and len(ga) >= 3:
            btype = "result"
        else:
            btype = "cells"
        blocks.append((btype, cl, (r1, r2, c1, c2), ga))
    passed_b = 0
    for btype, cl, (r1, r2, c1, c2), ga in blocks:
        rng_txt = (f"{get_column_letter(c1)}{r1}:{get_column_letter(c2)}{r2}"
                   if (r1, c1) != (r2, c2) else f"{get_column_letter(c1)}{r1}")
        gs = _block_grid(shs, r1, r2, c1, c2) if shs is not None else []
        if btype == "cond":
            ok, msgs = compare_condition_block(ga, gs)
            if ok:
                passed_b += 1
                continue
            equiv = filter_result_equivalent(shp, ga, gs) if gs else None
            note_txt = COND_EQUIV_NOTE if equiv else None
            res.details.append(
                f"[고급 필터 조건] {rng_txt}: " + "; ".join(msgs)
                + (" (필터 결과는 동일 — 조건식 표기 문제)" if equiv else ""))
            add_card(res, f"고급 필터 조건 ({rng_txt})",
                     res.alloc / len(blocks), "cell",
                     cells=[{"coord": rng_txt}],
                     props=[{"name": "조건", "expected": "문제 지시 조건",
                             "got": m} for m in msgs[:3]],
                     note=note_txt,
                     hint="필드명은 원본 표 머리글과 똑같이, AND 조건은 같은 "
                          "행, OR 조건은 다른 행에 적습니다. 열 순서는 달라도 "
                          "됩니다.")
        elif btype == "result":
            ok, msgs, order_diff = compare_result_block(ga, gs)
            if ok:
                passed_b += 1
                if order_diff:
                    res.notes.append("고급 필터 결과의 행 순서가 정답과 "
                                     "다르지만 내용이 같아 정답 처리했습니다.")
                continue
            res.details.append(f"[고급 필터 결과] {rng_txt}: "
                               + "; ".join(msgs))
            add_card(res, f"고급 필터 결과 ({rng_txt})",
                     res.alloc / len(blocks), "cell",
                     cells=[{"coord": rng_txt}],
                     props=[{"name": "결과 표", "expected": "지시 조건에 맞는 "
                             "행 추출", "got": m} for m in msgs[:3]],
                     hint="데이터 탭 → 고급에서 목록/조건/복사 위치 범위를 "
                          "다시 확인하세요.")
        else:
            cells3 = [(r, c, coord_of[(r, c)]) for (r, c) in cl]
            wrong = [(r, c, co) for (r, c, co) in cells3
                     if not judge.judge(r, c)[0]]
            if not wrong:
                passed_b += 1
                continue
            res.details.append(f"[결과 셀] {rng_txt}: {len(wrong)}개 불일치")
            add_wrong_cells(res, judge, wrong, limit=6)
            entries = make_cell_entries(judge, wrong, 3)
            add_card(res, f"결과 셀 ({rng_txt})", res.alloc / len(blocks),
                     "cell", cells=entries, more=max(0, len(wrong) - 3),
                     hint="텍스트 나누기/자동 필터 등 지시 결과를 원본과 "
                          "대조하세요.")
    res.earned = int(round(res.alloc * passed_b / len(blocks)))

CALC_FUNC_NOTE = ("이 채점기는 결과값 기준이라 문제가 제시한 함수를 썼는지는 "
                  "판정하지 못합니다. 실제 시험은 제시 함수 외 사용 시 값이 "
                  "맞아도 오답이니 문제지의 함수 지정을 확인하세요.")

def grade_calc(res, ctx):
    res.notes.append(CALC_FUNC_NOTE)
    diffs = ctx["vdiffs"]
    if not diffs:
        res.earned = res.alloc
        res.notes.append("문제/정답 간 차이가 없어 자동 만점 처리")
        return
    judge = ctx["judge"]
    key_groups = (ctx["key"].get("groups") or {}).get(norm_sheet_name(res.name))
    sha = ctx["book_a"].raw[ctx["asheet"]]
    if not key_groups:
        f_diffs = [(r, c, co) for (r, c, co) in diffs
                   if formula_text(sha.cell(r, c).value) is not None]
        d_diffs = [d for d in diffs if d not in f_diffs]
        if f_diffs and d_diffs:
            dc = compress_coords([co for (_r, _c, co) in d_diffs])
            res.notes.append(
                "정답이 수식이 아닌 셀 " + ", ".join(dc[:4])
                + (" 외" if len(dc) > 4 else "")
                + "은(는) 자료(데이터) 셀로 보아 계산 문제에서 제외했습니다 — "
                  "계산작업은 수식으로 답하는 문제만 채점합니다.")
            diffs = f_diffs
    if key_groups:
        from openpyxl.utils import get_column_letter
        groups = []
        for g in key_groups:
            cells = []
            for ref in g:
                for (r, c) in expand_ref(ref):
                    cells.append((r, c, f"{get_column_letter(c)}{r}"))
            if cells:
                groups.append(cells)
    else:
        clusters = cluster_cells([(r, c) for (r, c, _) in diffs], CALC_GROUP_TARGET)
        co = {(r, c): coord for (r, c, coord) in diffs}
        groups = [[(r, c, co[(r, c)]) for (r, c) in cl] for cl in clusters]
    n = len(groups)
    per = res.alloc / n
    contaminated = []
    contaminated_diff = []
    try:
        shp = ctx["book_p"].raw[ctx["psheet"]]
        diff_set = {(r, c) for (r, c, _) in ctx["vdiffs"]}
        for row in sha.iter_rows(max_row=min(sha.max_row or 1, MAX_SCAN_ROWS),
                                 max_col=min(sha.max_column or 1,
                                             MAX_SCAN_COLS)):
            for cell in row:
                fa = formula_text(cell.value)
                if fa is None:
                    continue
                fp = formula_text(shp.cell(cell.row, cell.column).value)
                if fp is None:
                    continue
                if (cell.row, cell.column) in diff_set:
                    if norm_formula(fp) != norm_formula(fa):
                        contaminated_diff.append(cell.coordinate)
                elif norm_formula(fp) == norm_formula(fa):
                    contaminated.append(cell.coordinate)
    except Exception:
        contaminated, contaminated_diff = [], []
    if key_groups or n >= CALC_GROUP_TARGET:
        contaminated = []
    passed = 0
    for gi, cells in enumerate(groups, 1):
        wrong = [(r, c, co) for (r, c, co) in cells if not judge.judge(r, c)[0]]
        rng = f"{cells[0][2]}~{cells[-1][2]}" if len(cells) > 1 else cells[0][2]
        if not wrong:
            passed += 1
        else:
            res.details.append(
                f"[문제 {gi}] ({rng}) {len(cells)}개 셀 중 {len(wrong)}개 불일치 "
                f"→ 0/{per:.0f}점 (문제 단위 전부 정답 원칙)")
            add_wrong_cells(res, judge, wrong, limit=6)
            entries = make_cell_entries(judge, wrong, 3)
            rep_formula = next((e["formula"] for e in entries if e["formula"]),
                               None)
            if rep_formula is None:
                for (r, c, co) in cells:
                    ef = judge.judge(r, c)[3]
                    if ef:
                        rep_formula = ef
                        break
            note = (f"대상 범위 {rng} · 인식된 문제 {n}개 중 {gi}번 "
                    f"(문제당 {per:.0f}점)")
            if n != CALC_GROUP_TARGET:
                note += (f" — 표준 {CALC_GROUP_TARGET}문제와 달라 배점이 "
                         "실제 시험과 다르게 배분됐습니다")
            guide = cell_state_guide(judge, wrong)
            if guide:
                note += ". " + guide
            errn = wrong_error_note(judge, wrong)
            if errn:
                note += " " + errn
            gr_ref = [m for m in judge.ref_mismatch
                      if any(m[0] == co for (_r, _c, co) in cells)]
            if gr_ref:
                coord, ef2, sf2 = gr_ref[0]
                more_a, more_b = ref_usage_diff(ef2, sf2)
                note += (f" 다시 계산한 결과값은 정답과 같지만 참조하는 셀이 "
                         f"다릅니다 — {coord}에서 정답은 "
                         f"{', '.join(more_a) or '(없음)'}, 내 답은 "
                         f"{', '.join(more_b) or '(없음)'}을(를) 참조합니다. "
                         "이 자료에서 우연히 결과가 같았을 뿐이라 정답으로 "
                         "인정하지 않습니다.")
            add_card(res, f"계산 문제 {gi}", per, "cell", cells=entries,
                     formula=rep_formula, note=note,
                     more=max(0, len(wrong) - 3),
                     hint="한 셀이라도 틀리면 문제 전체가 0점입니다. 첫 셀에 "
                          "수식을 완성한 뒤 채우기 핸들로 복사하고, 결과가 "
                          "이상한 셀이 없는지 훑어보세요.")
    res.earned = int(round(per * passed))
    cmn = compare_mode_note(judge)
    if cmn:
        res.notes.append(cmn)
    rmn = ref_mismatch_note(judge)
    if rmn:
        res.notes.append(rmn)
    res.notes.append(f"수식 문제를 {n}개 그룹으로 인식 (그룹당 {per:.0f}점)"
                     + ("" if n == CALC_GROUP_TARGET else
                        f" — 표준은 {CALC_GROUP_TARGET}문제입니다. 배점 배분이 "
                        "실제 시험과 다를 수 있으니 문제 파일이 원본인지 "
                        "확인하세요"))
    if contaminated:
        cc = compress_coords(contaminated)
        res.notes.append(
            f"주의: 문제 파일의 계산작업에 정답과 같은 수식이 이미 "
            f"{len(contaminated)}개 들어 있습니다 ({', '.join(cc[:4])}"
            f"{' 외' if len(cc) > 4 else ''}). 원본 문제 파일이 이전 풀이로 "
            "덮어써진 것(문제 파일 오염) 같습니다. 이 상태에서는 남은 부분만 "
            f"문제로 인식돼 배점이 {n}개 문제에 몰립니다(문제당 {per:.0f}점). "
            "원본 문제 파일을 다시 받아 채점하세요.")
    if contaminated_diff:
        cc = compress_coords(contaminated_diff)
        res.notes.append(
            f"주의: 문제 파일의 계산작업에 정답과 다른 수식이 이미 "
            f"{len(contaminated_diff)}개 들어 있습니다 ({', '.join(cc[:4])}"
            f"{' 외' if len(cc) > 4 else ''}). 원본 문제에는 없는 수식이므로 "
            "문제 파일이 이전 풀이로 덮어써진 것(문제 파일 오염)으로 보입니다. "
            "문제↔정답 차이가 줄어 문제 수·배점·다른 시트의 채점 기준이 실제와 "
            "달라질 수 있으니 원본 문제 파일을 다시 받아 채점하세요"
            + ("" if key_groups else
               " (세트 기대값 파일이 있으면 문제 구성은 그 파일을 따릅니다)")
            + ".")
    if contaminated or contaminated_diff:
        res.contaminated = compress_coords(contaminated + contaminated_diff)

def grade_analysis(res, ctx):
    book_a, book_s = ctx["book_a"], ctx["book_s"]
    asheet, ssheet = ctx["asheet"], ctx["ssheet"]
    diffs = ctx["vdiffs"]
    judge = ctx["judge"]
    if diffs:
        wrong = [(r, c, co) for (r, c, co) in diffs if not judge.judge(r, c)[0]]
        matched = len(diffs) - len(wrong)
        if wrong:
            res.earned = 0
            res.details.append(
                f"결과 셀 {len(diffs)}개 중 {matched}개 일치, {len(wrong)}개 "
                f"불일치 → 0점 (실제 시험 규정: 분석작업은 부분점수 없음)")
            add_wrong_cells(res, judge, wrong, limit=15)
            entries = make_cell_entries(judge, wrong, 3)
            add_card(res, "분석 결과", res.alloc, "cell", cells=entries,
                     formula=next((e["formula"] for e in entries
                                   if e["formula"]), None),
                     more=max(0, len(wrong) - 3),
                     hint="분석작업은 한 곳만 틀려도 0점입니다. 정렬 -> 부분합 "
                          "같은 지시 순서를 지키고, 완료 후 결과 표를 문제지의 "
                          "그림과 대조하세요.")
        else:
            res.earned = res.alloc
            res.details.append(f"결과 셀 {len(diffs)}개 전부 일치")
    else:
        res.earned = res.alloc
    for has_a, has_s, label in (
        (book_a.has_scenarios(asheet),
         ssheet and book_s.has_scenarios(ssheet), "시나리오"),
        (book_a.has_pivot_for_sheet(asheet),
         ssheet and book_s.has_pivot_for_sheet(ssheet), "피벗 테이블"),
    ):
        if has_a:
            if not has_s:
                res.notes.append(
                    f"[{label}] 학생 파일에서 {label}을(를) 찾지 못했습니다. "
                    "자동 감점은 하지 않으니 직접 확인하세요.")
            res.notes.append(f"{label} 세부 설정은 자동 비교하지 않습니다. "
                             "수동 확인을 권장합니다.")

def _macro_fmt_desc(kind, sig, raw=None):
    if sig == "?" or is_unknown(sig):
        return None
    if kind == "font":
        return describe_font(sig)
    if kind == "fill":
        return describe_fill(sig)
    if kind == "border":
        return describe_border(sig)
    if kind == "number_format":
        return nf_display(raw if raw not in (None, "") else sig)
    if kind == "alignment":
        sa = sig if isinstance(sig, tuple) else (None, None, False, 0)
        return (f"가로 {H_ALIGN_KO.get(sa[0], sa[0])} / "
                f"세로 {V_ALIGN_KO.get(sa[1], sa[1])}"
                + (", 텍스트 줄 바꿈" if len(sa) > 2 and sa[2] else ""))
    return str(sig)

def _font_color_txt(color):
    if color is None:
        return "자동(검정)"
    if isinstance(color, tuple) and color[0] == "rgb":
        return color_name_ko(color[1])
    if isinstance(color, tuple) and color[0] == "theme":
        return f"테마 색 {color[1]}"
    return "지정 색"

def font_pair_desc(sig_a, sig_s):
    a, s = describe_font(sig_a), describe_font(sig_s)
    try:
        ca, cs = sig_a[6], sig_s[6]
    except Exception:
        return a, s
    if ca == cs:
        return a, s
    if ca is None and a != "확인 불가":
        a += f", 글꼴 색 {_font_color_txt(ca)}"
    if cs is None and s != "확인 불가":
        s += f", 글꼴 색 {_font_color_txt(cs)}"
    return a, s

def macro_objects(book, sheet):
    out = []
    if not sheet:
        return out
    try:
        ctrls = book.form_controls(sheet)
    except Exception as e:
        ctrls = Unknown(f"{book.label} 파일의 양식 컨트롤을 읽지 "
                        f"못했습니다({type(e).__name__})")
    if is_unknown(ctrls):
        return Unknown(unknown_why(ctrls))
    for b in ctrls:
        out.append({"kind": "단추", "text": b.get("text", ""),
                    "anchor": b.get("anchor", "?"),
                    "macro": b.get("macro", ""), "geom": ""})
    try:
        shapes = book.drawing_shapes(sheet)
    except Exception as e:
        shapes = Unknown(f"{book.label} 파일의 도형을 읽지 "
                         f"못했습니다({type(e).__name__})")
    if is_unknown(shapes):
        return Unknown(unknown_why(shapes))
    for s in shapes:
        if s.get("hidden"):
            continue
        out.append({"kind": "도형", "text": s.get("text", ""),
                    "anchor": s.get("anchor", "?"),
                    "macro": s.get("macro", ""),
                    "geom": s.get("geom", "")})
    return out

SHAPE_GEOM_KO = {"hexagon": "육각형", "rect": "직사각형",
                 "roundRect": "모서리가 둥근 직사각형", "ellipse": "타원",
                 "flowChartProcess": "순서도: 처리", "pentagon": "오각형",
                 "diamond": "다이아몬드", "octagon": "팔각형"}

def _obj_label(o):
    g = SHAPE_GEOM_KO.get(o.get("geom"), o.get("geom") or "")
    head = f"{o['kind']}({g})" if g else o["kind"]
    return (f"[{o['anchor']}] {head} 텍스트 '{o['text'] or '(없음)'}'"
            + (f" → {o['macro']} 매크로" if o.get("macro") else ""))

def compare_macro_objects(objs_a, objs_s):
    left = list(objs_s)
    pairs = []
    for a in objs_a:
        mate = None
        for pick in (lambda s: s["anchor"] != "?" and s["anchor"] == a["anchor"],
                     lambda s: a.get("macro") and s.get("macro") == a["macro"],
                     lambda s: s["kind"] == a["kind"]):
            for s in left:
                if pick(s):
                    mate = s
                    break
            if mate:
                break
        if mate:
            left.remove(mate)
        diffs = []
        if mate is None:
            diffs.append("찾지 못함")
        else:
            if (a.get("text") or "") != (mate.get("text") or ""):
                diffs.append(f"텍스트 '{mate.get('text') or '(없음)'}'"
                             f" → '{a.get('text') or '(없음)'}'이어야 합니다")
            if a.get("geom") and mate.get("geom") and \
                    a["geom"] != mate["geom"]:
                diffs.append(
                    f"도형 종류 {SHAPE_GEOM_KO.get(mate['geom'], mate['geom'])}"
                    f" → {SHAPE_GEOM_KO.get(a['geom'], a['geom'])}")
            if a.get("macro") and mate.get("macro") and \
                    a["macro"] != mate["macro"]:
                diffs.append(f"연결 매크로 {mate['macro']} → {a['macro']}")
        pairs.append((a, mate, diffs))
    return pairs

def _macro_evidence_notes(res, book_a, book_s, asheet, ssheet):
    names_s = [u["name"] for u in book_s.vba_units()] if book_s else []
    if names_s:
        res.notes.append("내 파일 VBA 코드에서 찾은 매크로: "
                         + ", ".join(dict.fromkeys(names_s)))
    elif book_s and book_s.has_vba():
        why = book_s.vba_diag() or "원인을 특정하지 못했습니다"
        res.notes.append("VBA를 읽지 못했습니다 — 내 파일에 매크로"
                         f"(vbaProject)는 있는데 {why}. 매크로 이름·코드는 "
                         "판정에 쓰지 않았고, 채점은 실행 결과(값·서식)만 "
                         "보고 했습니다.")
    btns = book_s.form_controls(ssheet) if (book_s and ssheet) else []
    if btns:
        res.notes.append("내 파일 단추: " + " / ".join(
            f"[{b['anchor']}] 텍스트 '{b['text'] or '(없음)'}'"
            + (f" → {b['macro']} 매크로" if b["macro"] else " → 연결 매크로 확인 불가")
            for b in btns[:4]))
    elif ssheet and book_s.sheet_has_drawing(ssheet):
        res.notes.append("내 파일에 단추/도형은 있지만 앵커·연결 매크로를 "
                         "읽지 못했습니다.")
    shp = [s for s in (book_s.drawing_shapes(ssheet) if (book_s and ssheet)
                       else []) if not s.get("hidden")]
    if shp:
        res.notes.append("내 파일 도형: " + " / ".join(
            f"[{s['anchor']}] {SHAPE_GEOM_KO.get(s['geom'], s['geom'] or '도형')}"
            f" 텍스트 '{s['text'] or '(없음)'}'"
            + (f" → {s['macro']} 매크로" if s["macro"] else "")
            for s in shp[:4]))
    return names_s

def grade_macro(res, ctx):
    book_a, book_s = ctx["book_a"], ctx["book_s"]
    asheet, ssheet = ctx["asheet"], ctx["ssheet"]
    judge = ctx["judge"]
    shs = book_s.raw[ssheet] if ssheet else None
    sha = book_a.raw[asheet]
    result_pts = max(0, res.alloc - 2)
    units_total = 0
    units_ok = 0
    units_held = 0
    vba_units = book_s.vba_units() if book_s else []
    vba_blind = is_unknown(vba_units)
    vba_blind_why = unknown_why(vba_units) if vba_blind else None
    vba_credit = []
    held_units = []
    blind_whys = []
    wrong_cells = []
    for (r, c, co) in ctx["vdiffs"]:
        units_total += 1
        if judge.judge(r, c)[0]:
            units_ok += 1
            continue
        mac = vba_covers(vba_units, co, "value")
        if mac:
            units_ok += 1
            vba_credit.append((co, "value", mac))
            continue
        if vba_blind:
            units_held += 1
            held_units.append((co, "value"))
            continue
        wrong_cells.append((r, c, co))
    from openpyxl.utils.cell import coordinate_to_tuple
    a_members, a_anchors = merge_maps(sha)
    s_members, s_anchors = merge_maps(shs) if shs is not None else ({}, {})
    s_edges = None
    if "border" in ctx["fdiffs"] and shs is not None:
        _mr, _mc = scan_bounds(sha, shs)
        s_edges = sheet_edge_map(shs, _mr, _mc)
    fmt_bad = []
    fmt_ok = {}
    for (where, ukind, why) in (ctx.get("fmt_unknown") or []):
        units_total += 1
        units_held += 1
        held_units.append((where, ukind))
        if why not in blind_whys:
            blind_whys.append(why)
    for kind, payload in ctx["fdiffs"].items():
        if kind in ("rowheight", "colwidth", "merge", "merge_del", "names",
                    "border_edges", "_unknown"):
            continue
        if kind == "border":
            edge_info = ctx["fdiffs"].get("border_edges") or {}
            for coord in payload:
                units_total += 1
                ok = (s_edges is not None and not is_unknown(s_edges)
                      and all((s_edges.get(key) or None) == (style or None)
                              for key, style in edge_info.get(coord, [])))
                mac = None if ok else vba_covers(vba_units, coord, "border")
                if ok or mac:
                    units_ok += 1
                    fmt_ok.setdefault("border", []).append(coord)
                    if mac:
                        vba_credit.append((coord, "border", mac))
                    continue
                if s_edges is None or is_unknown(s_edges) or vba_blind:
                    units_held += 1
                    held_units.append((coord, "border"))
                    if is_unknown(s_edges) and \
                            hold_reason(s_edges) not in blind_whys:
                        blind_whys.append(hold_reason(s_edges))
                    continue
                exp_b, got_b = _border_props(
                    [{"coord": coord, "edges": edge_info.get(coord, [])}],
                    s_edges)
                fmt_bad.append({
                    "coord": coord, "kind": "border",
                    "expected": exp_b, "got": got_b,
                    "why": None if shs is not None else
                    f"학생 파일에 '{asheet}' 시트가 없습니다"})
            continue
        for coord in payload:
            units_total += 1
            sig_a = sig_s = None
            blind = None
            try:
                r, c = coordinate_to_tuple(coord)
                sig_a = fmt_signature_m(sha, r, c, kind, a_members, a_anchors)
                sig_s = fmt_signature_m(shs, r, c, kind, s_members,
                                        s_anchors) if shs is not None \
                    else Unknown(f"학생 파일에 '{asheet}' 시트가 없습니다")
                if any_unknown(sig_a, sig_s):
                    blind = hold_reason(sig_a, sig_s)
                ok = (not blind) and shs is not None and sig_a == sig_s
                if not ok and not blind and kind == "number_format" \
                        and shs is not None \
                        and nf_display_equal(sha, shs, r, c, a_members,
                                             s_members):
                    ok = True
            except Exception as e:
                ok = False
                blind = (f"{coord} 의 서식을 비교하지 못했습니다"
                         f"({type(e).__name__})")
            mac = None if ok else vba_covers(vba_units, coord, kind)
            if ok or mac:
                units_ok += 1
                fmt_ok.setdefault(kind, []).append(coord)
                if mac:
                    vba_credit.append((coord, kind, mac))
                continue
            if blind or vba_blind:
                units_held += 1
                held_units.append((coord, kind))
                if blind and blind not in blind_whys:
                    blind_whys.append(blind)
                continue
            raw_a = raw_s = None
            if kind == "number_format":
                try:
                    raw_a = cell_raw_m(sha, r, c, a_members, "number_format")
                    raw_s = cell_raw_m(shs, r, c, s_members, "number_format")
                except Exception:
                    pass
            exp_d = _macro_fmt_desc(kind, sig_a, raw_a)
            got_d = _macro_fmt_desc(kind, sig_s, raw_s)
            if kind == "font" and exp_d and got_d:
                exp_d, got_d = font_pair_desc(sig_a, sig_s)
            why = None
            if shs is None:
                why = f"학생 파일에 '{asheet}' 시트가 없습니다"
            elif exp_d is None:
                why = "정답 파일의 서식을 읽을 수 없습니다"
            elif got_d is None:
                why = "내 파일의 서식을 읽을 수 없습니다"
            fmt_bad.append({"coord": coord, "kind": kind,
                            "expected": exp_d, "got": got_d, "why": why})
    if units_total:
        per_unit = result_pts / units_total
        earned = int(round(per_unit * units_ok))
        if units_held:
            kinds_txt = ", ".join(dict.fromkeys(
                f"{co}({FMT_KIND_LABEL.get(k, k)})" for co, k in held_units
            ))[:120]
            why = " / ".join(
                [w for w in ([vba_blind_why] if vba_blind_why else [])
                 + blind_whys]) or "판정 근거를 읽지 못했습니다"
            add_hold(res, "매크로 실행 결과·서식",
                     why + f" — 판정하지 못한 항목: {kinds_txt}",
                     per_unit * units_held, what=kinds_txt)
    else:
        earned = result_pts
        res.notes.append("매크로 결과에 대한 diff가 없어 결과 점수 자동 부여")
    if wrong_cells:
        res.details.append(f"매크로 결과 셀 {len(wrong_cells)}개 불일치")
        add_wrong_cells(res, judge, wrong_cells, limit=10)
        entries = make_cell_entries(judge, wrong_cells, 3)
        add_card(res, "매크로 실행 결과",
                 max(0.0, result_pts - (earned if units_total else result_pts)),
                 "macro", cells=entries,
                 formula=next((e["formula"] for e in entries
                               if e["formula"]), None),
                 more=max(0, len(wrong_cells) - 3),
                 hint="매크로 기록(개발 도구 → 매크로 기록)을 시작한 뒤 지시 "
                      "작업을 수행하고 기록을 중지했는지, 기록 후 단추로 실행해 "
                      "결과가 나오는지 확인하세요.")
    if fmt_bad:
        by_kind = {}
        for b in fmt_bad:
            by_kind.setdefault(b["kind"], []).append(b["coord"])
        chips = [f"{FMT_KIND_LABEL.get(k, k)}: {', '.join(compress_coords(v)[:3])}"
                 for k, v in by_kind.items()]
        res.details.append("매크로 서식 불일치: " + ", ".join(chips))
        cells = []
        for b in fmt_bad[:6]:
            label = FMT_KIND_LABEL.get(b["kind"], b["kind"])
            if b["why"]:
                exp = got = f"판정 불가: {b['why']}"
                if b["expected"]:
                    exp = b["expected"]
                cells.append({"coord": f"{b['coord']} ({label})",
                              "expected": exp, "got": got, "formula": None})
            else:
                cells.append({"coord": f"{b['coord']} ({label})",
                              "expected": b["expected"],
                              "got": b["got"] or "(서식 없음)",
                              "formula": None})
        props = []
        for k, coords in by_kind.items():
            first = next(b for b in fmt_bad if b["kind"] == k)
            props.append({
                "name": f"{FMT_KIND_LABEL.get(k, k)}"
                        f"({', '.join(compress_coords(coords)[:2])})",
                "expected": first["expected"]
                or f"판정 불가: {first['why'] or '알 수 없음'}",
                "got": first["got"]
                or f"판정 불가: {first['why'] or '알 수 없음'}"})
        okay_note = ""
        ok_chips = []
        for k in by_kind:
            good = [co for co in fmt_ok.get(k, [])
                    if co not in by_kind[k]]
            if good:
                ok_chips.append(f"{FMT_KIND_LABEL.get(k, k)}: "
                                + ", ".join(compress_coords(good)[:3]))
        if ok_chips:
            okay_note = (" 같은 지시에서 " + " / ".join(ok_chips)
                         + " 영역은 정확히 적용됐고, 위에 적은 범위에만 "
                         "적용되지 않았습니다.")
        code_note = ""
        if vba_units:
            bad_kinds = set(by_kind)
            rel = [u for u in vba_units if u["kinds"] & bad_kinds]
            if rel:
                parts = []
                for u in rel[:2]:
                    rngs = ", ".join(dict.fromkeys(u["ranges"]))[:80] or "없음"
                    parts.append(f"'{u['name']}' 매크로의 적용 범위 [{rngs}]")
                code_note = (" 내 VBA 코드에 " + " / ".join(parts)
                             + "에는 이 셀이 들어 있지 않습니다 — 매크로는 "
                               "있지만 적용 범위가 지시와 다릅니다.")
            else:
                code_note = (" 내 VBA 코드에는 이 셀·서식을 지정하는 문장이 "
                             "없었습니다.")
        elif book_s.has_vba():
            why = book_s.vba_diag() or "원인을 특정하지 못했습니다"
            code_note = (f" VBA를 읽지 못했습니다({why}) — 코드 인정은 "
                         "못 하고 실행 결과만 보고 판정했습니다.")
        else:
            code_note = (" 내 파일에 매크로(vbaProject)가 없습니다 — 매크로는 "
                         "xlsm(매크로 사용 통합 문서)으로 저장해야 남습니다.")
        add_card(res, "매크로 서식 결과", 0, "macro", cells=cells, props=props,
                 note="정답 파일의 매크로 실행 결과 서식과 내 파일을 셀 단위로 "
                      "비교했습니다." + okay_note + code_note,
                 hint="서식 매크로가 지시한 서식(채우기·글꼴 등)을 정확히 "
                      "기록하고, 기록 후 단추를 눌러 실제로 실행했는지 "
                      "확인하세요.")
    if vba_credit:
        txt = ", ".join(f"{co}({FMT_KIND_LABEL.get(k, k)}) ← {m}"
                        for co, k, m in vba_credit[:5])
        res.notes.append(
            "실행 결과 서식/값은 정답과 다르지만 매크로 코드에 해당 동작이 "
            f"있어 인정했습니다: {txt}. 실제 시험은 '생성하여 실행'까지 "
            "요구하므로 단추를 눌러 결과까지 남기세요.")
    btn_a = book_a.sheet_has_drawing(asheet)
    btn_s = ssheet and book_s.sheet_has_drawing(ssheet)
    obj_held = False
    if any_unknown(btn_a, btn_s):
        add_hold(res, "매크로 단추·도형",
                 hold_reason(btn_a, btn_s)
                 + " — 단추(도형)가 있는지 확인하지 못해 이 2점은 점수에서 "
                   "뺐습니다.",
                 2, what="단추/도형")
        obj_held = True
        btn_a = False
    if btn_a:
        objs_a = macro_objects(book_a, asheet)
        objs_s = macro_objects(book_s, ssheet)
        if any_unknown(objs_a, objs_s):
            add_hold(res, "매크로 단추·도형",
                     hold_reason(objs_a, objs_s)
                     + " — 단추·도형의 텍스트/연결 매크로를 대조하지 못해 "
                       "이 2점은 점수에서 뺐습니다.",
                     2, what="단추/도형")
            obj_held = True
        elif btn_s:
            obj_pts = 2
            bad_objs = [(a, m, d) for (a, m, d)
                        in compare_macro_objects(objs_a, objs_s) if d]
            if objs_a and bad_objs:
                obj_pts = 1
                lines = [f"{_obj_label(a)} — " + ", ".join(d)
                         for (a, m, d) in bad_objs[:3]]
                res.details.append("[단추/도형] 지시와 다름: "
                                   + " / ".join(lines))
                add_card(res, "매크로 단추·도형", 1, "macro",
                         props=[{"name": f"{a['kind']} {a['anchor']}",
                                 "expected": _obj_label(a),
                                 "got": _obj_label(m) if m
                                 else "찾지 못했습니다"}
                                for (a, m, d) in bad_objs[:3]],
                         note="문제지는 단추·도형의 텍스트와 연결 매크로까지 "
                              "지정합니다 — 정답 파일의 개체와 대조했습니다: "
                              + " / ".join(lines),
                         hint="도형(단추)을 우클릭 → 텍스트 편집으로 지시된 "
                              "글자를 정확히 입력하고, 우클릭 → 매크로 지정으로 "
                              "지시된 매크로를 연결했는지 확인하세요.")
            earned += obj_pts
        else:
            res.details.append("[단추/도형] 매크로 실행용 단추(도형)를 찾지 못함 (-2점)")
            add_card(res, "매크로 단추(도형)", 2, "macro",
                     hint="개발 도구 → 삽입 → 단추(양식 컨트롤) 또는 삽입 → "
                          "도형을 그린 뒤 우클릭 → 매크로 지정으로 연결하고, "
                          "단추 텍스트를 지시대로 입력합니다.")
    elif not obj_held:
        gradable_units = units_total - units_held
        earned += 2 if gradable_units <= 0 or units_ok == gradable_units else 0
    res.earned = min(res.alloc - res.held, earned)
    _macro_evidence_notes(res, book_a, book_s, asheet, ssheet)
    if book_a.has_vba() and not book_s.has_vba():
        res.notes.append("학생 파일에 매크로(vbaProject)가 없습니다. "
                         "매크로는 xlsm 형식으로 저장해야 보존됩니다.")
    res.notes.append("매크로는 실행 결과(값·서식)를 기준으로 채점하고, 결과가 "
                     "없으면 VBA 코드에 지시 동작이 있는지까지 확인합니다.")

def grade_chart(res, ctx):
    book_a, book_s = ctx["book_a"], ctx["book_s"]
    asheet, ssheet = ctx["asheet"], ctx["ssheet"]
    fa = chart_features(book_a, asheet)
    fs = chart_features(book_s, ssheet) if ssheet else None
    fp = chart_features(ctx["book_p"], ctx["psheet"])
    per = res.alloc / 5.0
    if any_unknown(fa, fs):
        add_hold(res, "차트작업",
                 hold_reason(fa, fs)
                 + " — 차트를 대조하지 못해 이 문항은 점수에서 뺐습니다.",
                 res.alloc, what="차트")
        res.earned = 0
        res.details.append("차트를 읽지 못해 판정을 보류했습니다 (확인 필요)")
        return
    if is_unknown(fp):
        fp = None
        res.notes.append("문제 파일의 차트를 읽지 못해 '문제·정답이 같은 "
                         "특성은 지시가 아니다' 규칙을 적용하지 못했습니다.")
    if fa is None:
        res.earned = res.alloc
        res.notes.append("정답 파일에서 차트를 찾지 못해 자동 만점 처리")
        return
    if fs is None:
        res.earned = 0
        res.details.append("학생 파일에서 차트를 찾지 못했습니다 (0점)")
        add_card(res, "차트 없음", res.alloc, "chart",
                 hint="차트가 저장되지 않았습니다. 문제의 기본 차트를 지우지 "
                      "말고 그 위에서 지시 사항만 수정하세요.")
        return
    passed = 0

    def fail(label, exp_txt, got_txt, hint):
        res.details.append(f"[{label}] 기대 {exp_txt} / 학생 {got_txt}")
        add_card(res, f"차트 - {label}", per, "chart",
                 cells=[{"coord": None, "expected": str(exp_txt),
                         "got": str(got_txt), "formula": None}],
                 hint=hint)

    def _title_key(f):
        t = f["title"][1] if f and f["title"] else None
        return t.replace(" ", "") if isinstance(t, str) else None

    def _series_key(f):
        return (len(f["series"]), sorted(x[0] or "" for x in f["series"]))

    def _dlbls_key(f):
        return sorted((x[0] or "") for x in f["series"] if x[1])

    def same_as_problem(key_fn):
        return fp is not None and key_fn(fp) == key_fn(fa)

    if same_as_problem(lambda f: f["types"]) or fa["types"] == fs["types"]:
        passed += 1
    else:
        fail("차트 종류", fa["types"], fs["types"],
             "차트 영역 우클릭 → 차트 종류 변경. 콤보(혼합)형이면 계열별 "
             "차트 종류와 보조 축 체크를 확인하세요.")
    ta, ts = _title_key(fa), _title_key(fs)
    if same_as_problem(_title_key) or \
            ((ta is None) == (ts is None) and (ta is None or ta == ts)):
        passed += 1
    else:
        fail("차트 제목",
             (fa["title"][1] if fa["title"] else "(제목 없음)"),
             (fs["title"][1] if fs["title"] else "(제목 없음)"),
             "차트 선택 → 차트 요소(+) → 차트 제목 체크 후 텍스트를 "
             "입력합니다(제거 지시면 체크 해제).")
    if same_as_problem(_series_key) or _series_key(fa) == _series_key(fs):
        passed += 1
    else:
        refs_a = sorted(x[0] or "" for x in fa["series"])
        refs_s = sorted(x[0] or "" for x in fs["series"])
        fail("데이터 계열", f"{len(fa['series'])}개 {refs_a}",
             f"{len(fs['series'])}개 {refs_s}",
             "차트 우클릭 → 데이터 선택에서 계열을 문제 지시대로 추가/제거 "
             "하고 각 계열의 값 범위를 확인합니다.")
    dl_a, dl_s = _dlbls_key(fa), _dlbls_key(fs)
    if same_as_problem(_dlbls_key) or dl_a == dl_s:
        passed += 1
    else:
        fail("데이터 레이블", f"계열 {dl_a or '없음'}", f"계열 {dl_s or '없음'}",
             "레이블을 붙일 계열만 한 번 클릭해 선택한 뒤 차트 요소(+) → "
             "데이터 레이블에서 위치를 지정합니다.")

    def axis_item():
        A, S = _axes_by_role(fa), _axes_by_role(fs)
        P = _axes_by_role(fp) if fp is not None else {}
        graded_any, ok, exp, got = False, True, [], []

        def _txt(role, attr, v, missing=False):
            return (f"{AXIS_ROLE_KO.get(role, role)} {AXIS_ATTR_KO[attr]} "
                    f"{_num_txt(v)}" + (" (축 없음)" if missing else ""))

        spec = ((ctx["key"].get("chart") or {}).get(norm_sheet_name(res.name))
                or {}).get("axes")
        if spec:
            for item in (spec if isinstance(spec, list) else [spec]):
                if not isinstance(item, dict):
                    continue
                role = _axis_role_name(item.get("axis"))
                a_s = S.get(role)
                for attr, keys in (("max", ("max",)), ("min", ("min",)),
                                   ("unit", ("major_unit", "unit"))):
                    want = next((item[k] for k in keys
                                 if item.get(k) is not None), None)
                    if want is None:
                        continue
                    graded_any = True
                    have = a_s.get(attr) if a_s else None
                    if _axis_attr_ok(attr, float(want), have, a_s):
                        continue
                    ok = False
                    exp.append(_txt(role, attr, want))
                    got.append(_txt(role, attr, have, missing=a_s is None))
            return graded_any, ok, ", ".join(exp), ", ".join(got)
        for role, a_a in A.items():
            a_p, a_s = P.get(role), S.get(role)
            for attr in ("max", "min", "unit"):
                aa = a_a.get(attr)
                if aa is None:
                    continue
                pa = a_p.get(attr) if a_p else None
                if pa is not None and abs(pa - aa) <= 1e-9 * max(1.0, abs(aa)):
                    continue
                graded_any = True
                sa = a_s.get(attr) if a_s else None
                if _axis_attr_ok(attr, aa, sa, a_s, lenient_unit=(pa is None)):
                    continue
                ok = False
                exp.append(_txt(role, attr, aa))
                got.append(_txt(role, attr, sa, missing=a_s is None))
        return graded_any, ok, ", ".join(exp), ", ".join(got)

    ax_graded, ax_ok, ax_exp, ax_got = axis_item()
    if ax_graded:
        if ax_ok:
            passed += 1
        else:
            fail("축 설정", ax_exp, ax_got,
                 "축(기본 축은 왼쪽, 보조 축은 오른쪽 세로 축)을 더블클릭 → "
                 "축 서식 → 축 옵션에서 경계 최소/최대와 단위(기본)를 지시 "
                 "값으로 입력합니다. 지시에 없는 항목은 '자동'으로 두세요.")
    else:
        legend_graded = fp is None or fp["legend"] != fa["legend"]
        if not legend_graded or fa["legend"] == fs["legend"]:
            passed += 1
        else:
            fail("범례 위치", fa["legend"], fs["legend"],
                 "차트 요소(+) → 범례에서 위치(아래쪽/오른쪽 등)를 "
                 "지정합니다.")
    res.earned = int(round(per * passed))
    if passed < 5:
        res.details.insert(0, f"차트 검사 5개 항목 중 {passed}개 통과 "
                              f"(항목당 {per:.0f}점)")

def grade_generic(res, ctx):
    diffs = ctx["vdiffs"]
    if not diffs:
        res.earned = res.alloc
        return
    judge = ctx["judge"]
    wrong = [(r, c, co) for (r, c, co) in diffs if not judge.judge(r, c)[0]]
    res.earned = int(round(res.alloc * (len(diffs) - len(wrong)) / len(diffs)))
    if wrong:
        res.details.append(f"diff 셀 {len(diffs)}개 중 {len(wrong)}개 불일치")
        add_wrong_cells(res, judge, wrong, limit=15)
        entries = make_cell_entries(judge, wrong, 3)
        add_card(res, "오답 셀", res.alloc - res.earned, "cell",
                 cells=entries, more=max(0, len(wrong) - 3),
                 formula=next((e["formula"] for e in entries
                               if e["formula"]), None))

GRADERS = {
    "기본작업-1": grade_basic1,
    "기본작업-2": grade_basic2,
    "기본작업-3": grade_basic3,
    "계산작업": grade_calc,
    "분석작업-1": grade_analysis,
    "분석작업-2": grade_analysis,
    "매크로작업": grade_macro,
    "차트작업": grade_chart,
}

NEEDS_FMT_DIFF = {"기본작업-2", "매크로작업"}

def run_grading(problem_path, answer_path, student_path, key, sheets=None):
    book_p = Book(problem_path, "문제")
    book_a = Book(answer_path, "정답")
    book_s = Book(student_path, "학생")

    points = dict(DEFAULT_POINTS)
    for k, v in (key.get("points") or {}).items():
        points[norm_sheet_name(k)] = v

    global_notes = []
    results = []

    matched = []
    for psheet in book_p.raw.sheetnames:
        n = norm_sheet_name(psheet)
        asheet = book_a.norm_map.get(n)
        if asheet is None:
            global_notes.append(f"정답 파일에 '{psheet}' 시트가 없어 채점에서 제외")
            continue
        ssheet = book_s.norm_map.get(n)
        matched.append((n, psheet, asheet, ssheet))
    for asheet in book_a.raw.sheetnames:
        n = norm_sheet_name(asheet)
        if n not in {m[0] for m in matched} and n not in \
                {norm_sheet_name(s) for s in book_p.raw.sheetnames}:
            global_notes.append(f"정답 전용 시트 '{asheet}'는 채점 대상이 아닙니다 "
                                "(예: 시나리오 요약)")

    if not matched:
        raise RuntimeError("문제/정답 파일에서 매칭되는 시트를 찾지 못했습니다.")

    if sheets:
        wanted = [norm_sheet_name(s) for s in sheets if str(s).strip()]
        available = [m[0] for m in matched]
        unknown = [w for w in wanted if w not in available]
        if unknown:
            raise RuntimeError(
                "--sheets에 지정한 시트를 찾을 수 없습니다: "
                + ", ".join(unknown)
                + "\n사용 가능한 시트: " + ", ".join(available))
        matched = [m for m in matched if m[0] in wanted]
        global_notes.append("부분 채점: " + ", ".join(m[1] for m in matched)
                            + " 시트만 채점했습니다.")

    ctxs = {}
    nonstd = []
    for (n, psheet, asheet, ssheet) in matched:
        shp, sha = book_p.raw[psheet], book_a.raw[asheet]
        vdiffs = value_diff_cells(shp, sha)
        fdiffs = format_diff_items(book_p, book_a, psheet, asheet) \
            if n in NEEDS_FMT_DIFF else {}
        fmt_unknown = fdiffs.pop("_unknown", []) if fdiffs else []
        key_cells = {k.upper(): v for k, v in
                     ((key.get("cells") or {}).get(n) or {}).items()}
        excl_notes = []
        memos = detect_author_memo_cells(shp, sha, vdiffs)
        if memos:
            vdiffs, fdiffs, memo_notes = apply_key_exclusions(
                vdiffs, fdiffs, memos, label="제작 메모로 보아 채점 제외")
            excl_notes.extend(memo_notes)
        excl = (key.get("exclude") or {}).get(n)
        if excl:
            vdiffs, fdiffs, more_notes = apply_key_exclusions(vdiffs, fdiffs, excl)
            excl_notes.extend(more_notes)
        judge = CellJudge(book_a, book_s, asheet, ssheet, key_cells,
                          display_tolerant=(n == "기본작업-1"))
        ctxs[n] = {
            "book_p": book_p, "book_a": book_a, "book_s": book_s,
            "psheet": psheet, "asheet": asheet, "ssheet": ssheet,
            "vdiffs": vdiffs, "fdiffs": fdiffs, "judge": judge, "key": key,
            "exclude_notes": excl_notes, "fmt_unknown": fmt_unknown,
        }
        if n not in points:
            nonstd.append(n)

    std_total = sum(points[n] for (n, *_ ) in matched if n in points)
    if nonstd:
        remain = max(0, 100 - std_total)
        weights = {n: max(1, len(ctxs[n]["vdiffs"])) for n in nonstd}
        wsum = sum(weights.values())
        base = remain if std_total else 100
        for n in nonstd:
            points[n] = base * weights[n] / wsum
        global_notes.append("표준 시트명이 아닌 시트는 diff 셀 수 비례로 "
                            "배점을 자동 배분했습니다.")

    formula_only = False
    recalc_any = False
    for (n, psheet, asheet, ssheet) in matched:
        alloc = points.get(n, 0)
        res = SheetResult(psheet, alloc)
        ctx = ctxs[n]
        try:
            res.diff_total = len(ctx["vdiffs"])
            if ssheet is not None:
                res.diff_matched = sum(
                    1 for (rr, cc, _co) in ctx["vdiffs"]
                    if ctx["judge"].judge(rr, cc)[0])
        except Exception:
            pass
        if ssheet is None:
            res.missing = True
            res.earned = 0
            res.details.append("학생 파일에 이 시트가 없습니다 (0점)")
            results.append(res)
            continue
        grader = GRADERS.get(n, grade_generic)
        try:
            grader(res, ctx)
        except Exception as e:
            res.notes.append(f"채점 중 오류가 발생해 값 비교로 대체: {e}")
            add_hold(res, f"{psheet} 채점",
                     f"채점기가 이 시트를 끝까지 처리하지 못했습니다"
                     f"({type(e).__name__}: {e}) — 점수에서 뺐습니다.",
                     res.alloc, what=psheet)
            try:
                grade_generic(res, ctx)
            except Exception:
                pass
        res.earned = max(0, min(res.alloc if isinstance(res.alloc, int)
                                else res.alloc, res.earned))
        res.earned = max(0, min(res.earned, res.alloc - res.held))
        for note in ctx.get("exclude_notes") or []:
            res.notes.append(note)
        if ctx["judge"].formula_only_used:
            formula_only = True
        if getattr(ctx["judge"], "recalc_cells", None):
            recalc_any = True
        results.append(res)

    if recalc_any:
        global_notes.append("일부 셀은 계산된 값이 파일에 없어 LibreOffice로 다시 "
                            "계산해 결과값으로 채점했습니다. Excel에서 파일을 열어 "
                            "저장(Ctrl+S)하면 저장된 계산값으로 채점됩니다.")
    if formula_only:
        global_notes.append("일부 셀은 계산된 값이 없어 수식 기준으로 채점했습니다. "
                            "Excel에서 파일을 열어 저장하면 값 기준 채점이 가능합니다.")

    total_alloc = sum(r.alloc for r in results)
    total_earned = sum(r.earned for r in results)
    total_held = round(sum(getattr(r, "held", 0) for r in results), 2)
    gradable = max(0.0, total_alloc - total_held)
    if sheets:
        score100 = int(round(total_earned))
    elif not gradable:
        score100 = 0
        global_notes.append("채점 가능한 항목이 하나도 없었습니다 — 모든 "
                            "항목이 '확인 필요'입니다. 점수는 참고하지 "
                            "마세요.")
    elif abs(gradable - 100) > 0.01:
        score100 = round(total_earned / gradable * 100)
        if total_held:
            global_notes.append(
                f"'확인 필요' {len(hold_items(results))}건({total_held:g}점 "
                f"분량)은 판정 근거를 읽지 못한 항목이라 감점도 가점도 하지 "
                f"않고 만점에서 뺐습니다. 점수 = 득점 {total_earned:g} ÷ "
                f"채점된 배점 {gradable:g} × 100.")
        else:
            global_notes.append(f"채점 대상 배점 합계가 {total_alloc:.0f}점이라 "
                                "100점 만점으로 환산했습니다.")
    else:
        score100 = int(round(total_earned))
    try:
        enrich_wrong_cards(results)
    except Exception:
        pass
    return results, score100, global_notes, (book_p, book_a, book_s)

_CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮"

_FTOKEN = re.compile(r'''
    (?P<str>"(?:[^"]|"")*")
  | (?P<funcopen>[A-Za-z_][A-Za-z0-9_.]*\()
  | (?P<ref>\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)
  | (?P<name>[A-Za-z가-힣_][A-Za-z가-힣0-9_.]*)
  | (?P<num>\d+(?:\.\d+)?%?)
  | (?P<op><>|<=|>=|[=<>+\-*/&^%])
  | (?P<comma>,)
  | (?P<open>\()
  | (?P<close>\))
  | (?P<ws>\s+)
''', re.VERBOSE)

def _parse_formula(f):
    s = str(f).lstrip("=").strip()
    s = re.sub(r"'[^']*'!", "", s)
    s = re.sub(r"\[\d+\]", "", s)
    s = re.sub(r"[A-Za-z가-힣0-9_.]+!", "", s)
    toks = []
    for m in _FTOKEN.finditer(s):
        if m.lastgroup != "ws":
            toks.append((m.lastgroup, m.group()))
    pos = [0]

    def peek():
        return toks[pos[0]] if pos[0] < len(toks) else (None, None)

    def parse_expr():
        parts = []
        while True:
            k, v = peek()
            if k is None or k in ("comma", "close"):
                break
            if k == "funcopen":
                pos[0] += 1
                name = v[:-1].strip().upper().replace("_XLFN.", "")
                if name == "RANK":
                    name = "RANK.EQ"
                parts.append(("func", name, parse_args()))
            elif k == "open":
                pos[0] += 1
                inner = parse_expr()
                if peek()[0] == "close":
                    pos[0] += 1
                parts.append(inner)
            else:
                parts.append((k, v))
                pos[0] += 1
        if not parts:
            return ("empty", "")
        if len(parts) == 1:
            return parts[0]
        return ("expr", parts)

    def parse_args():
        args = []
        if peek()[0] == "close":
            pos[0] += 1
            return args
        while True:
            args.append(parse_expr())
            k, _v = peek()
            if k == "comma":
                pos[0] += 1
                continue
            if k == "close":
                pos[0] += 1
            break
        return args

    return parse_expr()

def _collect_funcs(node, out):
    if node[0] == "func":
        for a in node[2]:
            _collect_funcs(a, out)
        out.append(node)
    elif node[0] == "expr":
        for p in node[1]:
            _collect_funcs(p, out)

def _render(node, marks):
    kind = node[0]
    if kind == "func":
        if id(node) in marks:
            return marks[id(node)]
        return _fn_display(node, marks)
    if kind == "expr":
        return "".join(_render(p, marks) for p in node[1])
    if kind == "empty":
        return "(생략)"
    if kind == "ref":
        return node[1].replace("$", "")
    return node[1]

def _fn_display(node, marks):
    inner = ", ".join(_render(a, marks) for a in node[2])
    return f"{node[1]}({inner})"

def _lit_num(node):
    try:
        if node[0] == "num":
            v = node[1].rstrip("%")
            return float(v) if "." in v else int(v)
        if node[0] == "expr" and len(node[1]) == 2 and \
                node[1][0] == ("op", "-") and node[1][1][0] == "num":
            return -_lit_num(node[1][1])
    except Exception:
        pass
    return None

def _rng(text):
    return f"{text} 범위" if ":" in text else text

_DB_VERBS = {"DSUM": "합계", "DAVERAGE": "평균", "DCOUNT": "숫자 셀 개수",
             "DCOUNTA": "개수", "DMAX": "최대값", "DMIN": "최소값",
             "DGET": "값", "DPRODUCT": "곱"}
_AGG_KO = {"SUM": "합계", "AVERAGE": "평균", "MAX": "최대값", "MIN": "최소값",
           "MEDIAN": "중간값(중앙값)", "COUNT": "숫자 셀 개수",
           "COUNTA": "비어 있지 않은 셀 개수", "STDEV": "표준편차",
           "MODE": "최빈값", "MODE.SNGL": "최빈값", "VAR": "분산"}

def _round_pos(node):
    v = _lit_num(node)
    if v is None:
        return "지정한 자리로"
    v = int(v)
    if v > 0:
        return f"소수 {v}째 자리까지"
    if v == 0:
        return "일의 자리까지(정수로)"
    unit = {-1: "십", -2: "백", -3: "천", -4: "만"}.get(v, f"10^{-v}")
    return f"{unit} 단위로"

def _fn_sentence(name, args, R):
    n = len(args)

    def A(i, default=""):
        return R(args[i]) if i < n else default

    if name == "IF":
        return (f"조건 {A(0)} 이(가) 참이면 {A(1)}, 아니면 "
                f"{A(2, '빈 값')} 을(를) 표시합니다.")
    if name == "IFS":
        pairs = [f"{A(i)} → {A(i + 1)}" for i in range(0, n - 1, 2)]
        return ("조건을 앞에서부터 차례로 검사해 처음 참이 되는 조건의 값을 "
                "표시합니다: " + ", ".join(pairs))
    if name == "AND":
        return f"{', '.join(A(i) for i in range(n))} 조건이 모두 참인지 확인합니다."
    if name == "OR":
        return (f"{', '.join(A(i) for i in range(n))} 조건 중 하나라도 참인지 "
                "확인합니다.")
    if name == "IFERROR":
        return f"{A(0)} 이(가) 오류이면 {A(1)} 을(를) 대신 표시합니다."
    if name in ("VLOOKUP", "HLOOKUP"):
        where = "첫 열(세로)" if name == "VLOOKUP" else "첫 행(가로)"
        which = "열" if name == "VLOOKUP" else "행"
        mode = ""
        if n >= 4:
            last = A(3).upper()
            mode = (" (정확히 일치하는 값 찾기)" if last in ("0", "FALSE")
                    else " (구간/근사 일치: 찾는 값이 속하는 구간)")
        else:
            mode = " (구간/근사 일치: 찾는 값이 속하는 구간)"
        return (f"{A(0)} 값을 {_rng(A(1))}의 {where}에서 찾아, 같은 줄의 "
                f"{A(2)}번째 {which} 값을 가져옵니다{mode}.")
    if name == "INDEX":
        if n >= 3:
            return (f"{_rng(A(0))}에서 {A(1)}번째 행, {A(2)}번째 열이 만나는 "
                    "값을 가져옵니다.")
        return f"{_rng(A(0))}에서 {A(1)}번째 값을 가져옵니다."
    if name == "MATCH":
        mode = " (정확히 일치)" if n >= 3 and A(2) == "0" else ""
        return (f"{A(0)} 값이 {_rng(A(1))}에서 몇 번째에 있는지 위치(순번)를 "
                f"구합니다{mode}.")
    if name == "CHOOSE":
        opts = ", ".join(f"{i}이면 {A(i)}" for i in range(1, n))
        return f"{A(0)} 값이 {opts} 을(를) 선택합니다."
    if name == "LEFT":
        return f"{A(0)}의 왼쪽에서 {A(1, '1')}글자를 꺼냅니다."
    if name == "RIGHT":
        return f"{A(0)}의 오른쪽에서 {A(1, '1')}글자를 꺼냅니다."
    if name == "MID":
        return f"{A(0)}의 {A(1)}번째 글자부터 {A(2)}글자를 꺼냅니다."
    if name in ("SEARCH", "FIND"):
        extra = "(대소문자 구분 없음)" if name == "SEARCH" else "(대소문자 구분)"
        return f"{A(1)}에서 {A(0)} 이(가) 몇 번째 글자에 있는지 찾습니다{extra}."
    if name == "PROPER":
        return f"{A(0)}의 각 단어 첫 글자만 대문자로 바꿉니다."
    if name == "UPPER":
        return f"{A(0)}을(를) 모두 대문자로 바꿉니다."
    if name == "LOWER":
        return f"{A(0)}을(를) 모두 소문자로 바꿉니다."
    if name == "TRIM":
        return f"{A(0)}의 불필요한 공백을 제거합니다."
    if name == "LEN":
        return f"{A(0)}의 글자 수를 셉니다."
    if name in ("HOUR", "MINUTE", "SECOND"):
        ko = {"HOUR": "시", "MINUTE": "분", "SECOND": "초"}[name]
        return f"{A(0)}에서 '{ko}' 부분의 숫자만 꺼냅니다."
    if name in ("YEAR", "MONTH", "DAY"):
        ko = {"YEAR": "연도", "MONTH": "월", "DAY": "일"}[name]
        return f"{A(0)}에서 {ko} 숫자만 꺼냅니다."
    if name == "WEEKDAY":
        t = _lit_num(args[1]) if n >= 2 else None
        base = {2: " (옵션 2: 월요일=1 ~ 일요일=7)",
                3: " (옵션 3: 월요일=0 시작)"}.get(
            t, " (기본: 일요일=1 ~ 토요일=7)")
        return f"{A(0)}의 요일 번호를 구합니다{base}."
    if name == "MOD":
        return f"{A(0)}을(를) {A(1)}(으)로 나눈 나머지를 구합니다."
    if name == "QUOTIENT":
        return f"{A(0)}을(를) {A(1)}(으)로 나눈 몫(정수)을 구합니다."
    if name == "TODAY":
        return "오늘 날짜를 구합니다."
    if name == "NOW":
        return "현재 날짜와 시각을 구합니다."
    if name == "DATE":
        return f"연 {A(0)}, 월 {A(1)}, 일 {A(2)} 로 날짜를 만듭니다."
    if name == "TIME":
        return f"시 {A(0)}, 분 {A(1)}, 초 {A(2)} 로 시각을 만듭니다."
    if name in _DB_VERBS:
        return (f"{_rng(A(0))}(데이터 전체, 필드명 포함)에서 조건 범위 "
                f"{_rng(A(2))}의 조건에 맞는 행만 골라 {A(1)} 필드의 "
                f"{_DB_VERBS[name]}을(를) 구합니다.")
    if name == "COUNTIF":
        return f"{_rng(A(0))}에서 조건 {A(1)} 에 맞는 셀의 개수를 셉니다."
    if name == "COUNTIFS":
        pairs = [f"{_rng(A(i))}가 {A(i + 1)}" for i in range(0, n - 1, 2)]
        return "조건 " + " 그리고 ".join(pairs) + " 을(를) 모두 만족하는 행의 개수를 셉니다."
    if name == "SUMIF":
        if n >= 3:
            return (f"{_rng(A(0))}이 조건 {A(1)} 에 맞는 행의 {_rng(A(2))} "
                    "값을 모두 더합니다.")
        return f"{_rng(A(0))}에서 조건 {A(1)} 에 맞는 값을 모두 더합니다."
    if name in ("SUMIFS", "AVERAGEIFS"):
        verb = "더합니다" if name == "SUMIFS" else "평균을 구합니다"
        pairs = [f"{_rng(A(i))}가 {A(i + 1)}" for i in range(1, n - 1, 2)]
        return (f"조건 {' 그리고 '.join(pairs)} 을(를) 모두 만족하는 행의 "
                f"{_rng(A(0))} 값을 {verb}.")
    if name == "AVERAGEIF":
        if n >= 3:
            return (f"{_rng(A(0))}이 조건 {A(1)} 에 맞는 행의 {_rng(A(2))} "
                    "값의 평균을 구합니다.")
        return f"{_rng(A(0))}에서 조건 {A(1)} 에 맞는 값의 평균을 구합니다."
    if name == "LARGE":
        return f"{_rng(A(0))}에서 {A(1)}번째로 큰 값을 구합니다."
    if name == "SMALL":
        return f"{_rng(A(0))}에서 {A(1)}번째로 작은 값을 구합니다."
    if name == "RANK.EQ":
        order = ""
        if n >= 3 and _lit_num(args[2]) == 1:
            order = " (오름차순: 작을수록 1위)"
        else:
            order = " (내림차순: 클수록 1위)"
        return f"{A(0)} 값이 {_rng(A(1))}에서 몇 위인지 순위를 구합니다{order}."
    if name in _AGG_KO:
        rngs = ", ".join(_rng(A(i)) for i in range(n))
        return f"{rngs}의 {_AGG_KO[name]}을(를) 구합니다."
    if name == "ROUND":
        return f"{A(0)}을(를) {_round_pos(args[1] if n > 1 else None)} 반올림합니다."
    if name == "ROUNDUP":
        return f"{A(0)}을(를) {_round_pos(args[1] if n > 1 else None)} 올림합니다."
    if name == "ROUNDDOWN":
        return f"{A(0)}을(를) {_round_pos(args[1] if n > 1 else None)} 버림(내림)합니다."
    if name == "TRUNC":
        return f"{A(0)}의 소수 부분을 버립니다."
    if name == "INT":
        return f"{A(0)}보다 크지 않은 정수로 내립니다(소수 버림)."
    if name == "ABS":
        return f"{A(0)}의 절댓값(음수면 부호를 뗀 값)을 구합니다."
    if name == "POWER":
        return f"{A(0)}의 {A(1)}제곱을 구합니다."
    if name == "SUMPRODUCT":
        return "짝이 되는 값끼리 곱한 뒤 모두 더합니다."
    if name in ("CONCAT", "CONCATENATE"):
        return f"{', '.join(A(i) for i in range(n))} 을(를) 이어 붙입니다."
    if name == "TEXT":
        return f"{A(0)}을(를) 표시 형식 {A(1)} 문자열로 바꿉니다."
    args_txt = ", ".join(A(i) for i in range(n))
    return f"{name} 함수로 {args_txt} 을(를) 계산합니다."

def explain_formula(formula):
    try:
        f = str(formula or "")
        if not f.startswith("="):
            return [], None
        if "__DATATABLE__" in f:
            return ["데이터 표 수식입니다. 데이터 탭 → 가상 분석 → 데이터 "
                    "표에서 행/열 입력 셀을 지정해 만듭니다."], None
        if "__ARRAY__" in f:
            return ["배열 수식입니다. 수식 입력 후 Ctrl+Shift+Enter로 "
                    "확정합니다."], None
        tree = _parse_formula(f)
        funcs = []
        _collect_funcs(tree, funcs)
        marks = {}
        steps = []
        for i, fn in enumerate(funcs[:15]):
            mark = _CIRCLED[i] if i < len(_CIRCLED) else f"({i + 1})"
            disp = _fn_display(fn, marks)
            sent = _fn_sentence(fn[1], fn[2], lambda nd: _render(nd, marks))
            steps.append(f"{mark} {disp}: {sent}")
            marks[id(fn)] = mark
        if tree[0] == "expr" and any(
                p[0] == "op" and p[1] == "&" for p in tree[1]):
            i = len(funcs)
            mark = _CIRCLED[i] if i < len(_CIRCLED) else f"({i + 1})"
            steps.append(f"{mark} {_render(tree, marks)}: 위 값들을 & 로 이어 "
                         "붙여 하나의 문자열로 만듭니다.")
        elif not funcs and tree[0] == "expr":
            steps.append(f"{_render(tree, marks)} 을(를) 계산합니다.")
        point = ABS_POINT if "$" in f else None
        return steps, point
    except Exception:
        return ["정답 수식을 한 단계씩 그대로 입력하며 구조를 익혀 보세요."], None

ABS_POINT = ("$ 절대참조는 채우기 핸들로 수식을 복사할 때 참조 범위가 "
             "밀리지 않도록 고정하는 표시입니다. 표/기준 범위에만 $를 "
             "붙이고, 행마다 바뀌어야 하는 셀에는 붙이지 않습니다.")
TEXT2NUM_POINT = (
    "MID·LEFT·RIGHT가 꺼낸 값은 '1'처럼 보여도 숫자가 아니라 문자입니다. "
    "Excel은 문자 \"1\" 과 숫자 1을 서로 다른 값으로 보기 때문에, 숫자로 된 "
    "표에서 찾거나 계산에 쓰려면 뒤에 *1을 붙이거나 VALUE(...)로 감싸 숫자로 "
    "바꿔야 합니다. (반대로 숫자를 문자처럼 붙일 때는 &\"\" 를 씁니다.)")
LOOKUP_POINT = (
    "VLOOKUP/HLOOKUP의 번호는 시트의 열·행 번호가 아니라 '찾을 범위 안에서' "
    "세는 번호입니다. 범위의 첫 열(HLOOKUP은 첫 행)이 1번이자 찾는 기준 줄이라, "
    "가져올 값이 있는 줄은 2번부터인 경우가 많습니다. 마지막 인수는 구간을 "
    "찾으면 TRUE, 딱 맞는 값을 찾으면 FALSE입니다.")
CRITERIA_POINT = (
    "COUNTIF(S)·SUMIF(S)의 조건은 통째로 하나의 '문자열'이어야 합니다. "
    "800000처럼 값이 고정돼 있으면 \">800000\" 처럼 따옴표 안에 그대로 적을 수 "
    "있지만, AVERAGE(...) 같은 계산 결과는 따옴표 안에 넣으면 글자 그대로 "
    "취급돼 계산되지 않습니다. 그래서 부등호만 따옴표로 적고 \">=\"&AVERAGE(...) "
    "처럼 & 로 이어 붙여 하나의 문자열을 만듭니다.")
CONCAT_POINT = (
    "계산 결과 뒤에 글자를 붙이라는 지시(예: 3 → 3명)는 & 연결로 처리합니다. "
    "표시 형식으로 꾸미면 값 자체는 숫자라 채점에서 다르게 볼 수 있으니, "
    "지시에 '붙여 표시'가 있으면 수식 끝에 &\"명\" 을 붙이세요.")

def formula_point(expected_f, student_f=None):
    try:
        ef = str(expected_f or "")
        sf = str(student_f or "")
        if not ef.startswith("="):
            return None
        eu, su = _flat(ef), _flat(sf)
        e_names = set(re.findall(r"([A-Z][A-Z.]{1,15})\(",
                                 re.sub(r'"[^"]*"', "", eu)))
        cut_used = any(fn + "(" in eu for fn in TEXT_CUT_FUNCS)
        e_conv = bool(re.search(r"\)\*1(?![\d.])", eu)) or "VALUE(" in eu
        s_conv = bool(re.search(r"\)\*1(?![\d.])", su)) or "VALUE(" in su
        if cut_used and e_conv and (not sf or not s_conv):
            return TEXT2NUM_POINT
        if e_names & {"VLOOKUP", "HLOOKUP"}:
            return LOOKUP_POINT
        if any(n.startswith(("COUNTIF", "SUMIF", "AVERAGEIF"))
               for n in e_names) and '"&' in eu.replace(" ", ""):
            return CRITERIA_POINT
        tail = _tail_concat_strings(_parse_formula(ef))
        if tail and (not sf or not _tail_concat_strings(_parse_formula(sf))):
            return CONCAT_POINT
        if "$" in ef:
            return ABS_POINT
        return None
    except Exception:
        return None

def explain_number_format(nf):
    try:
        nf = nf_display(nf)
        if not nf or nf in ("General", _NF_GENERAL_KO):
            return f"기본 표시 형식({_NF_GENERAL_KO})입니다."
        lits = [x for x in re.findall(r'"([^"]*)"', nf) if x.strip()]
        s = re.sub(r'"[^"]*"', "", nf)
        low = s.lower()
        feats = []
        if "#,##" in s or ",##0" in s:
            feats.append("천 단위 구분 기호(,) 표시")
        m = re.search(r"0\.(0+)", s)
        if m:
            feats.append(f"소수 {len(m.group(1))}자리까지 표시")
        if "%" in s:
            feats.append("백분율(%)로 표시")
        is_date = bool(re.search(r"[ymdhsa]", low))
        if not is_date:
            m = re.search(r"(?<![#,0.])0{2,}(?![#,0.])", s)
            if m:
                feats.append(f"숫자를 {len(m.group(0))}자리로 채워 표시"
                             "(빈 자리는 0)")
        if "@" in s:
            feats.append("입력한 텍스트는 그대로 두고")
        if is_date:
            bits = []
            if "yyyy" in low:
                bits.append("네 자리 연도")
            elif "yy" in low:
                bits.append("두 자리 연도")
            if re.search(r"(?<!a)m", low):
                bits.append("월(또는 분)")
            if re.search(r"(?<!d)d{1,2}(?!d)", low):
                bits.append("일")
            if "aaaa" in low:
                bits.append("한글 요일 전체(예: 월요일)")
            elif "aaa" in low:
                bits.append("한글 요일 약자(예: 월)")
            elif "ddd" in low:
                bits.append("영문 요일")
            if "h" in low:
                bits.append("시각")
            feats.append("날짜/시간을 " + "·".join(bits) + " 형태로 표시")
        for lit in lits:
            if is_date and lit.strip() in ("년", "월", "일", "시", "분", "초"):
                continue
            feats.append(f"문자 '{lit.strip()}' 붙여 표시")
        if not feats:
            return f"사용자 지정 형식 코드 {nf} 를 그대로 입력합니다."
        return f"형식 코드 {nf} 의 뜻: " + ", ".join(feats) + "."
    except Exception:
        return f"사용자 지정 형식 코드 {nf} 를 그대로 입력합니다."

def _walk_funcs(node, depth=0, out=None):
    if out is None:
        out = []
    if node[0] == "func":
        out.append((node[1], depth, node))
        for a in node[2]:
            _walk_funcs(a, depth + 1, out)
    elif node[0] == "expr":
        for p in node[1]:
            _walk_funcs(p, depth, out)
    return out

def _strip_strings(f):
    return re.sub(r'"[^"]*"', '""', str(f))

def _cmp_ops(f):
    return re.findall(r"<=|>=|<>|<|>", _strip_strings(f))

def _top_func_name(tree):
    if tree[0] == "func":
        return tree[1]
    if tree[0] == "expr":
        for p in tree[1]:
            if p[0] == "func":
                return p[1]
    return None

def _is_ref_node(n):
    return n[0] == "ref"

def _josa_wa(word):
    return f"{word}와(과)"

LOOKUP_ARG_KO = {
    "VLOOKUP": {1: "찾을 값", 2: "찾을 범위", 3: "가져올 열 번호",
                4: "일치 옵션"},
    "HLOOKUP": {1: "찾을 값", 2: "찾을 범위", 3: "가져올 행 번호",
                4: "일치 옵션"},
}
LOOKUP_FIRST_KO = {"VLOOKUP": "열", "HLOOKUP": "행"}
TEXT_CUT_FUNCS = ("MID", "LEFT", "RIGHT")

def _flat(s):
    return re.sub(r"\s+", "", str(s or "")).upper()

def _num_conv_kind(sa, ea):
    base, got = _flat(_render(sa, {})), _flat(_render(ea, {}))
    if not base or base == got:
        return None
    if got in (base + "*1", "1*" + base, base + "*1.0"):
        return "*1"
    if got == f"VALUE({base})":
        return "VALUE"
    return None

def _cond_arg_text(node):
    if node[0] == "str":
        return node[1].strip('"')
    if node[0] == "expr":
        out = []
        for p in node[1]:
            if p[0] == "op" and p[1] == "&":
                continue
            if p[0] == "str":
                out.append(p[1].strip('"'))
            elif p[0] == "num":
                out.append(p[1])
            else:
                return None
        if out:
            return "".join(out)
    return None

def _tail_concat_strings(tree):
    if not tree or tree[0] != "expr":
        return []
    out = []
    parts = tree[1]
    for i, p in enumerate(parts):
        if p[0] == "op" and p[1] == "&" and i + 1 < len(parts) \
                and parts[i + 1][0] == "str":
            out.append(parts[i + 1][1])
    return out

def _edit_distance(a, b, cap=2):
    if abs(len(a) - len(b)) > cap:
        return cap + 1
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1,
                           prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]

def _typo_pairs(extra, missing):
    out = []
    for x in extra:
        for m in missing:
            if x != m and _edit_distance(x, m) <= 2 and abs(len(x) - len(m)) <= 2:
                out.append((x, m))
                break
    return out

def diagnose_formula_diff(student_f, expected_f):
    fallback = ["정답 수식과 접근이 다릅니다. 아래 풀이를 따라 처음부터 "
                "다시 작성해 보세요."]
    try:
        sf, ef = str(student_f or ""), str(expected_f or "")
        if not sf.startswith("=") or not ef.startswith("="):
            return []

        def _abs_ref_notes():
            out = []
            e_anchored = set(re.findall(
                r"\$[A-Z]{1,3}\$\d+:\$[A-Z]{1,3}\$\d+", ef.upper()))
            for anch in sorted(e_anchored):
                plain = anch.replace("$", "")
                if re.search(re.escape(plain),
                             sf.upper().replace("$", "")) \
                        and anch not in sf.upper().replace(" ", ""):
                    out.append(f"{anch}처럼 범위를 $로 고정해야 채우기 "
                               "핸들로 복사할 때 범위가 어긋나지 않습니다.")
                    break
            return out

        if norm_formula(sf) == norm_formula(ef):
            if re.sub(r"\s+", "", sf.upper()) != \
                    re.sub(r"\s+", "", ef.upper()):
                return _abs_ref_notes()
            return []
        notes = []
        st, et = _parse_formula(sf), _parse_formula(ef)
        sfuncs = _walk_funcs(st)
        efuncs = _walk_funcs(et)
        snames = [n for n, _d, _x in sfuncs]
        enames = [n for n, _d, _x in efuncs]
        sset, eset = set(snames), set(enames)
        s_top, e_top = _top_func_name(st), _top_func_name(et)
        if s_top and e_top and s_top != e_top \
                and s_top in eset and e_top in sset:
            notes.append(
                f"{_josa_wa(s_top)} {e_top}의 중첩 순서가 반대입니다 — "
                f"정답은 {e_top}(...) 안에 {s_top}(...)이 들어가는 구조입니다.")
        else:
            missing = sorted(eset - sset)
            extra = sorted(sset - eset)
            typos = _typo_pairs(extra, missing)
            for got_n, want_n in typos:
                notes.append(
                    f"함수 이름 오타입니다: {got_n} → {want_n}. Excel은 모르는 "
                    f"이름을 만나면 계산하지 못하고 #NAME? 오류를 냅니다.")
            typo_got = {g for g, _w in typos}
            typo_want = {w for _g, w in typos}
            missing = [m for m in missing if m not in typo_want]
            extra = [x for x in extra if x not in typo_got]
            if missing:
                notes.append("정답에 있는 " + ", ".join(missing)
                             + " 함수가 내 수식에 없습니다.")
            if extra:
                notes.append("정답에 없는 " + ", ".join(extra)
                             + " 함수를 사용했습니다.")
        from collections import Counter
        sops, eops = Counter(_cmp_ops(sf)), Counter(_cmp_ops(ef))
        for a, b in ((">=", "<="), ("<=", ">="), (">", "<"), ("<", ">")):
            if sops[a] and eops[b] and not eops[a] and not sops[b]:
                notes.append(f"'{a}' 를 썼지만 '{b}' 여야 합니다 "
                             "(비교 방향이 반대).")
                break
        seen = set()
        for name in enames:
            if name in seen or name not in sset:
                continue
            seen.add(name)
            sn = next(x for n, _d, x in sfuncs if n == name)
            en = next(x for n, _d, x in efuncs if n == name)
            sargs, eargs = sn[2], en[2]
            if len(sargs) != len(eargs):
                notes.append(f"{name}의 인수 개수가 다릅니다 "
                             f"(내 수식 {len(sargs)}개 / 정답 {len(eargs)}개).")
                continue
            role = LOOKUP_ARG_KO.get(name) or {}
            for i, (sa, ea) in enumerate(zip(sargs, eargs), 1):
                rs, re_ = _render(sa, {}), _render(ea, {})
                if norm_formula("=" + rs) == norm_formula("=" + re_):
                    continue
                label = f"{name}의 {i}번째 인수"
                if role.get(i):
                    label = f"{name}의 {i}번째 인수({role[i]})"
                cs, ce = _cond_arg_text(sa), _cond_arg_text(ea)
                if cs is not None and ce is not None and cs == ce:
                    if rs != re_:
                        notes.append(
                            f"조건을 {rs} 처럼 이어 붙인 것은 {re_} 와 같은 "
                            "뜻입니다 — 이 부분은 맞습니다.")
                    continue
                conv = _num_conv_kind(sa, ea)
                if conv:
                    cut = next((fn for fn in TEXT_CUT_FUNCS
                                if fn in _flat(rs)), None)
                    rng = _render(eargs[1], {}) if len(eargs) > 1 else "찾을 범위"
                    head = (f"{cut} 함수가 꺼낸 값은 '1'처럼 보여도 숫자가 "
                            "아니라 문자입니다" if cut else
                            "이 값은 문자로 취급됩니다")
                    tail = ("*1 을 붙여" if conv == "*1" else "VALUE(...)로 감싸")
                    notes.append(
                        f"{head} — {rng}의 숫자와는 서로 다른 값으로 보기 "
                        f"때문에 찾지 못합니다(#N/A 또는 엉뚱한 값). 정답처럼 "
                        f"{tail} 숫자로 바꿔야 합니다: {re_}")
                    continue
                if _num_conv_kind(ea, sa):
                    notes.append(f"{label}에는 숫자 변환(*1 · VALUE)이 "
                                 f"필요 없습니다 — 정답은 {re_} 입니다.")
                    continue
                if name in LOOKUP_ARG_KO and i == 3 and \
                        sa[0] == "num" and ea[0] == "num":
                    unit = LOOKUP_FIRST_KO[name]
                    notes.append(
                        f"{label}는 찾을 범위 안에서 몇 번째 {unit}을 가져올지 "
                        f"세는 값입니다 — {re_}이어야 하는데 {rs}을(를) "
                        f"썼습니다(1은 찾는 기준 {unit}이라 찾던 값이 그대로 "
                        "나옵니다).")
                    continue
                if name in LOOKUP_ARG_KO and i == 4:
                    want_t = _flat(re_) in ("TRUE", "1")
                    notes.append(
                        f"{label}이 다릅니다 — 정답은 {re_}"
                        + ("(근사 일치: 구간에서 찾기)" if want_t
                           else "(정확히 일치)")
                        + f"인데 {rs}을(를) 썼습니다.")
                    continue
                if sa[0] == "num" and ea[0] == "num":
                    notes.append(f"{label}가 {re_}이어야 "
                                 f"하는데 {rs}을(를) 썼습니다.")
                elif ea[0] == "str" and sa[0] in ("name", "ref") and \
                        ea[1].strip('"') == rs:
                    notes.append(f"문자는 따옴표로 감싸야 합니다: {ea[1]}")
                elif _is_ref_node(sa) and _is_ref_node(ea):
                    if rs.replace("$", "") != re_.replace("$", ""):
                        notes.append(f"{name}의 {i}번째 참조가 {re_}이어야 "
                                     f"하는데 {rs}을(를) 썼습니다.")
        e_tail = _tail_concat_strings(et)
        s_tail = _tail_concat_strings(st)
        for lit in e_tail:
            if lit not in s_tail:
                notes.append(
                    f"결과 뒤에 &{lit} 이어 붙이기가 빠졌습니다 — 내 수식은 "
                    f"숫자만 내놓아 표시 예처럼 '{lit.strip(chr(34))}'이(가) "
                    "붙지 않습니다.")
        if not e_tail and "&" in _strip_strings(ef) \
                and "&" not in _strip_strings(sf):
            notes.append("& 연결이 빠졌습니다 — 계산 결과 뒤에 &\"문자\" "
                         "형태로 이어 붙여야 합니다.")
        elif "&" in _strip_strings(sf) and "&" not in _strip_strings(ef):
            notes.append("정답에는 & 연결이 없습니다.")
        notes.extend(_abs_ref_notes())
        uniq = []
        for n in notes:
            if n not in uniq:
                uniq.append(n)
        return uniq[:6] if uniq else fallback
    except Exception:
        return fallback

CALC_CAT_FUNCS = [
    ("DB", {"DSUM", "DAVERAGE", "DCOUNT", "DCOUNTA", "DMAX", "DMIN", "DGET",
            "DPRODUCT", "DVAR", "DSTDEV"}),
    ("시간날짜", {"HOUR", "MINUTE", "SECOND", "YEAR", "MONTH", "DAY",
               "WEEKDAY", "DATE", "TIME", "TODAY", "NOW", "MOD", "EDATE",
               "EOMONTH", "DAYS", "DATEDIF"}),
    ("참조", {"VLOOKUP", "HLOOKUP", "INDEX", "MATCH", "CHOOSE", "OFFSET",
            "LOOKUP", "XLOOKUP"}),
    ("문자열", {"LEFT", "MID", "RIGHT", "SEARCH", "FIND", "PROPER", "UPPER",
             "LOWER", "TRIM", "LEN", "REPT", "CONCAT", "CONCATENATE", "TEXT",
             "FIXED", "VALUE", "SUBSTITUTE", "REPLACE"}),
    ("판정", {"IF", "IFS", "AND", "OR", "NOT", "IFERROR", "SWITCH", "TRUE",
            "FALSE"}),
    ("통계조건", {"COUNTIF", "COUNTIFS", "SUMIF", "SUMIFS", "AVERAGEIF",
              "AVERAGEIFS", "LARGE", "SMALL", "RANK", "RANK.EQ", "MAX", "MIN",
              "MEDIAN", "COUNTA", "COUNT", "SUM", "AVERAGE", "MODE",
              "MODE.SNGL", "STDEV", "ROUND", "ROUNDUP", "ROUNDDOWN", "INT",
              "ABS", "TRUNC", "POWER", "SUMPRODUCT"}),
]

SHEET_CATEGORY = {
    "기본작업-1": "입력", "기본작업-2": "서식", "기본작업-3": "조건부서식·필터",
    "분석작업-1": "분석", "분석작업-2": "분석", "매크로작업": "매크로",
    "차트작업": "차트",
}

CAT_INFO = {
    "판정": {
        "이름": "판정(IF 계열)",
        "진단": "IF 중첩 순서와 AND/OR 조건식 구성에서 실수가 나옵니다.",
        "처방": ["계산드릴 '드릴1_판정' 시트 재풀이",
                 "루틴 페이지 실수 노트의 '판정' 항목에 이번 오답 수식 기록",
                 "IF·IFS·AND·OR 치트시트 복습 - 조건 검사 순서와 문자 결과의 "
                 "큰따옴표"],
    },
    "참조": {
        "이름": "찾기/참조(VLOOKUP 계열)",
        "진단": "찾을 범위 고정($)과 열/행 번호, 일치 옵션 지정이 약합니다.",
        "처방": ["계산드릴 '드릴2_참조' 시트 재풀이",
                 "루틴 페이지 실수 노트의 '참조' 항목에 이번 오답 수식 기록",
                 "VLOOKUP·HLOOKUP·INDEX/MATCH·CHOOSE 치트시트 복습 - 범위 "
                 "$고정과 정확히 일치(0) 옵션"],
    },
    "시간날짜": {
        "이름": "날짜/시간",
        "진단": "WEEKDAY 옵션, 시·분 추출, MOD 활용 등 날짜·시간 계산이 "
               "흔들립니다.",
        "처방": ["계산드릴 '드릴3_시간날짜' 시트 재풀이",
                 "루틴 페이지 실수 노트의 '시간날짜' 항목에 이번 오답 수식 기록",
                 "HOUR·MINUTE·YEAR·WEEKDAY·MOD 치트시트 복습 - WEEKDAY 두 번째 "
                 "인수(2=월요일 시작) 확인"],
    },
    "DB": {
        "이름": "데이터베이스 함수(D 계열)",
        "진단": "D함수의 세 인수(전체 범위·필드·조건 범위) 구조가 약합니다.",
        "처방": ["계산드릴 '드릴4_DB' 시트 재풀이",
                 "루틴 페이지 실수 노트의 'DB' 항목에 이번 오답 수식 기록",
                 "DSUM·DAVERAGE·DCOUNTA 치트시트 복습 - 데이터 범위는 필드명 "
                 "행부터, 조건 범위는 필드명+조건 세로 2칸"],
    },
    "통계조건": {
        "이름": "통계/조건 집계",
        "진단": "COUNTIF·SUMIF 계열의 조건 표기(따옴표·부등호)와 RANK·LARGE "
               "지정에서 실수가 나옵니다.",
        "처방": ["계산드릴 '드릴5_통계조건' 시트 재풀이",
                 "루틴 페이지 실수 노트의 '통계조건' 항목에 이번 오답 수식 기록",
                 "COUNTIF(S)·SUMIF(S)·RANK.EQ·LARGE/SMALL 치트시트 복습 - "
                 "조건은 \">=90\" 처럼 따옴표로"],
    },
    "문자열": {
        "이름": "문자열 처리",
        "진단": "글자 추출 위치·개수 지정과 & 연결 구성이 약합니다.",
        "처방": ["계산드릴 '드릴6_문자열' 시트 재풀이",
                 "루틴 페이지 실수 노트의 '문자열' 항목에 이번 오답 수식 기록",
                 "LEFT·MID·RIGHT·SEARCH·& 연결 치트시트 복습 - MID의 시작 "
                 "위치와 글자 수"],
    },
    "입력": {
        "이름": "자료 입력(기본작업-1)",
        "진단": "입력 오타·누락이 있습니다. 이 영역은 한 셀만 틀려도 5점 "
               "전체를 잃습니다.",
        "처방": ["입력을 마친 뒤 문제지와 셀 단위 1:1 대조(위→아래, 왼→오른쪽)",
                 "루틴 페이지 실수 노트에 틀린 셀과 원인(오타/누락) 기록"],
    },
    "서식": {
        "이름": "셀 서식(기본작업-2)",
        "진단": "표시 형식 코드·병합·테두리 등 서식 지시 이행이 누락됩니다.",
        "처방": ["셀 서식(Ctrl+1) 사용자 지정 형식 코드 5개 손으로 다시 쓰기 "
                 "(#,##0\"원\" 등)",
                 "루틴 페이지 실수 노트의 '서식' 항목에 틀린 항목 기록",
                 "기본작업-2 유형 1회분 재풀이"],
    },
    "조건부서식·필터": {
        "이름": "조건부 서식/고급 필터(기본작업-3)",
        "진단": "조건부 서식의 혼합참조($E5)나 고급 필터 조건 범위 작성이 "
               "약합니다.",
        "처방": ["조건부 서식 수식 규칙을 혼합참조로 3회 반복 작성",
                 "고급 필터 AND(같은 행)/OR(다른 행) 조건 배치 복습",
                 "루틴 페이지 실수 노트의 '기본작업-3' 항목에 기록"],
    },
    "분석": {
        "이름": "분석작업",
        "진단": "부분합·피벗·시나리오·목표값 등 분석 도구의 지시 순서 이행이 "
               "약합니다. 이 영역은 부분점수가 없습니다.",
        "처방": ["오답 난 분석 유형을 같은 파일에서 처음부터 다시 수행",
                 "정렬 → 부분합 순서 등 절차 암기(루틴 페이지 실수 노트에 기록)",
                 "완료 후 문제지의 결과 그림과 표를 1:1 대조"],
    },
    "매크로": {
        "이름": "매크로작업",
        "진단": "매크로 기록/단추 연결 절차가 미완성입니다.",
        "처방": ["매크로 기록 → 작업 → 기록 중지 → 단추에 연결 절차를 2회 반복",
                 "단추 텍스트·매크로 이름을 문제 지시와 똑같이 입력했는지 확인",
                 "xlsm 형식으로 저장하는 습관"],
    },
    "차트": {
        "이름": "차트작업",
        "진단": "차트 세부 요소(종류·제목·레이블·축) 설정이 누락됩니다.",
        "처방": ["차트 요소(+) 메뉴에서 제목·레이블·범례 넣고 빼기 연습",
                 "데이터 선택으로 계열 추가/제거 연습",
                 "루틴 페이지 실수 노트의 '차트' 항목에 틀린 요소 기록"],
    },
    "기타": {
        "이름": "기타",
        "진단": "그 밖의 항목에서 실점이 있습니다.",
        "처방": ["오답 항목을 같은 파일에서 다시 수행", "루틴 페이지 실수 노트에 기록"],
    },
}

def classify_formula_category(formula):
    try:
        f = re.sub(r'"[^"]*"', "", str(formula or "").upper())
        names = set(re.findall(r"([A-Z][A-Z.]{1,15})\s*\(", f))
        names = {("RANK.EQ" if x == "RANK" else x).replace("_XLFN.", "")
                 for x in names}
        for cat, funcs in CALC_CAT_FUNCS:
            if cat == "DB" and names & funcs:
                return "DB"
        scores = {cat: len(names & funcs) for cat, funcs in CALC_CAT_FUNCS}
        if "&" in f:
            scores["문자열"] = scores.get("문자열", 0) + 0.5
        best = max(scores.values() or [0])
        if best == 0:
            return "통계조건"
        for cat, _funcs in CALC_CAT_FUNCS:
            if scores.get(cat) == best:
                return cat
    except Exception:
        pass
    return "통계조건"

def enrich_wrong_cards(results):
    for r in results:
        nname = norm_sheet_name(r.name)
        for card in r.wrong:
            if nname == "계산작업":
                card["category"] = classify_formula_category(card.get("formula"))
            else:
                card["category"] = SHEET_CATEGORY.get(nname, "기타")
            card["expected_formula"] = card.get("formula")
            card["student_formula"] = card.get("student_formula") or next(
                (c.get("got_formula") for c in (card.get("cells") or [])
                 if c.get("got_formula")), None)
            pre_notes = list(card.get("diff_notes") or [])
            card["diff_notes"] = pre_notes
            if card.get("formula"):
                steps, point = explain_formula(card["formula"])
                card["explain"] = steps
                if point and not card.get("point"):
                    card["point"] = point
                if card["student_formula"] and not pre_notes:
                    card["diff_notes"] = diagnose_formula_diff(
                        card["student_formula"], card["formula"])
                if nname == "계산작업":
                    pt = formula_point(card["formula"],
                                       card["student_formula"])
                    if pt:
                        card["point"] = pt

def build_diagnosis(results):
    losses = {}
    for r in results:
        for card in r.wrong:
            cat = card.get("category") or "기타"
            losses[cat] = losses.get(cat, 0) + (card.get("lost") or 0)
    ordered = sorted(losses.items(), key=lambda kv: -kv[1])
    top3 = []
    for cat, lost in ordered[:3]:
        info = CAT_INFO.get(cat, CAT_INFO["기타"])
        top3.append({"category": cat, "이름": info["이름"], "lost": lost,
                     "진단": info["진단"], "처방": info["처방"]})
    checklist = []
    has_calc_wrong = any(norm_sheet_name(r.name) == "계산작업" and r.wrong
                         for r in results)
    if has_calc_wrong:
        checklist.append("오답 난 계산 문제의 정답 수식을 보지 않고 손으로 "
                         "3번씩 다시 써 보기")
    for t in top3:
        if t["처방"]:
            checklist.append(t["처방"][0])
    checklist.append("내일 공부 첫 10분: 이 리포트의 오답노트만 다시 읽기")
    seen = set()
    uniq = []
    for c in checklist:
        if c not in seen:
            seen.add(c)
            uniq.append(c)
    return ordered, top3, uniq[:5]

def print_console(results, score100, global_notes, paths, partial=False):
    line = "─" * 56
    print()
    print(line)
    print(" 컴활 2급 실기 자동 채점 결과")
    print(line)
    print(f" 문제 파일 : {os.path.basename(paths[0])}")
    print(f" 정답 파일 : {os.path.basename(paths[1])}")
    print(f" 학생 파일 : {os.path.basename(paths[2])}")
    print(line)
    print(" " + pad("시트", 22) + pad("배점", 8, "right")
          + pad("득점", 8, "right") + pad("판정", 10, "right"))
    print(line)
    holds = hold_items(results)
    held_total = hold_total(results)
    for r in results:
        alloc = f"{r.alloc:.0f}"
        gradable = r.alloc - getattr(r, "held", 0)
        if r.missing:
            verdict = "시트없음"
        elif getattr(r, "holds", None) and r.earned >= gradable:
            verdict = "확인필요"
        elif r.earned >= gradable and gradable > 0:
            verdict = "정답"
        elif r.earned > 0:
            verdict = "부분정답"
        elif getattr(r, "holds", None):
            verdict = "확인필요"
        else:
            verdict = "오답"
        print(" " + pad(r.name, 22) + pad(alloc, 8, "right")
              + pad(f"{r.earned:.0f}", 8, "right") + pad(verdict, 10, "right"))
    print(line)
    max_alloc = sum(r.alloc for r in results)
    if partial:
        names = ", ".join(r.name for r in results)
        print(" " + pad("부분 채점 합계", 22)
              + pad(f"{max_alloc:.0f}", 8, "right")
              + pad(str(score100), 8, "right"))
        print(f" 부분 채점: {names} ({max_alloc:.0f}점 만점)")
    else:
        verdict = "합격권" if score100 >= PASS_LINE else "미달"
        print(" " + pad("총점", 22) + pad("100", 8, "right")
              + pad(str(score100), 8, "right")
              + (pad("확인필요", 10, "right") if holds else ""))
        if holds:
            print(f" {score100}점 · 확인 필요 {len(holds)}건")
        print(f" 합격선 {PASS_LINE}점 기준: {verdict}")
    print(line)
    print(f" ※ {SELF_GRADE_DISCLAIMER}")
    print("   틀렸다고 본 항목은 HTML 리포트에서 직접 판정을 바꿀 수 "
          "있습니다 (최종 판정은 사용자).")
    if holds:
        print()
        print(" [확인 필요] 판정하지 못해 점수에서 뺀 항목")
        for h in holds:
            print(f"    - {h['sheet']} · {h['label']} "
                  f"({h['points']:g}점 분량)")
            print(f"      {h['why']}")
        print(f"    * 이 {len(holds)}건({held_total:g}점 분량)은 감점도 가점도 "
              "하지 않고 만점에서 뺐습니다.")
        print(f"      점수 = 득점 ÷ 채점된 배점 × 100 "
              f"(채점된 배점 = {sum(r.alloc for r in results):g} - "
              f"{held_total:g} = "
              f"{max(0.0, sum(r.alloc for r in results) - held_total):g}).")

    detailed = [r for r in results if r.details or r.notes]
    if detailed:
        print()
        print(" [상세 내역]")
        for r in detailed:
            print(f" ▷ {r.name} ({r.earned:.0f}/{r.alloc:.0f}점)")
            for d in r.details:
                print(f"    - {d}")
            for nt in r.notes:
                print(f"    * (참고) {nt}")
    if global_notes:
        print()
        print(" [안내]")
        for nt in global_notes:
            print(f"    * {nt}")
    print()

def esc(s):
    return html_mod.escape(str(s))

def _svg_donut(score, pass_line=PASS_LINE, max_total=100, show_pass=True):
    r = 56
    c = 2 * math.pi * r
    max_total = max_total or 100
    frac = max(0.0, min(1.0, score / float(max_total)))
    if show_pass:
        color = "#107C41" if score >= pass_line else "#B45309"
    else:
        color = "#107C41"
    tick = ""
    if show_pass:
        ang = 2 * math.pi * (pass_line / 100.0) - math.pi / 2
        x1 = 70 + (r - 10) * math.cos(ang)
        y1 = 70 + (r - 10) * math.sin(ang)
        x2 = 70 + (r + 10) * math.cos(ang)
        y2 = 70 + (r + 10) * math.sin(ang)
        tick = (f'<line x1="{x1:.1f}" y1="{y1:.1f}" x2="{x2:.1f}" '
                f'y2="{y2:.1f}" stroke="#B45309" stroke-width="2"/>')
    return (
        f'<svg width="150" height="150" viewBox="0 0 140 140" role="img" '
        f'aria-label="점수 {score}/{max_total}점">'
        f'<circle cx="70" cy="70" r="{r}" fill="none" stroke="#E8EFEA" '
        f'stroke-width="13"/>'
        f'<circle cx="70" cy="70" r="{r}" fill="none" stroke="{color}" '
        f'stroke-width="13" stroke-dasharray="{c * frac:.1f} {c:.1f}" '
        f'transform="rotate(-90 70 70)"/>' + tick +
        f'<text x="70" y="68" text-anchor="middle" font-size="32" '
        f'font-weight="700" fill="{color}" '
        f'style="font-variant-numeric:tabular-nums">{score}</text>'
        f'<text x="70" y="90" text-anchor="middle" font-size="12" '
        f'fill="#57705F">/ {max_total:.0f}점</text></svg>')

def _svg_sheet_bars(results):
    lab_w, bar_w, row_h, pad = 150, 380, 30, 10
    width = lab_w + bar_w + 90
    height = pad * 2 + row_h * len(results)
    rows = []
    for i, r in enumerate(results):
        y = pad + i * row_h
        cy = y + row_h / 2 + 4
        frac = (r.earned / r.alloc) if r.alloc else 0
        fill_w = bar_w * max(0.0, min(1.0, frac))
        zero = r.earned <= 0
        lab_color = "#B45309" if zero else "#1B3A26"
        bars = (f'<rect x="{lab_w}" y="{y + 7}" width="{bar_w}" '
                f'height="{row_h - 14}" rx="4" fill="#ECF1ED"/>')
        if fill_w > 0.5:
            bars += (f'<rect x="{lab_w}" y="{y + 7}" width="{fill_w:.1f}" '
                     f'height="{row_h - 14}" rx="4" fill="#107C41"/>')
        else:
            bars += (f'<rect x="{lab_w}" y="{y + 7}" width="5" '
                     f'height="{row_h - 14}" fill="#B45309"/>')
        rows.append(
            f'<text x="{lab_w - 8}" y="{cy}" text-anchor="end" font-size="12" '
            f'fill="{lab_color}">{esc(r.name)}</text>' + bars +
            f'<text x="{lab_w + bar_w + 8}" y="{cy}" font-size="12" '
            f'font-weight="600" fill="{lab_color}" '
            f'style="font-variant-numeric:tabular-nums">'
            f'{r.earned:.0f} / {r.alloc:.0f}</text>')
    return (f'<svg width="{width}" height="{height}" viewBox="0 0 {width} '
            f'{height}" role="img" aria-label="영역별 점수">'
            + "".join(rows) + "</svg>")

def _svg_trend(points):
    width, height = 560, 200
    ml, mr, mt, mb = 42, 18, 16, 34
    pw, ph = width - ml - mr, height - mt - mb
    n = len(points)

    def px(i):
        return ml + (pw * (i / (n - 1)) if n > 1 else pw / 2)

    def py(v):
        return mt + ph * (1 - max(0, min(100, v)) / 100.0)

    parts = []
    for g in (0, 25, 50, 75, 100):
        parts.append(f'<line x1="{ml}" y1="{py(g):.1f}" x2="{width - mr}" '
                     f'y2="{py(g):.1f}" stroke="#E8EFEA" stroke-width="1"/>')
        parts.append(f'<text x="{ml - 7}" y="{py(g) + 4:.1f}" '
                     f'text-anchor="end" font-size="10" fill="#8AA093">'
                     f'{g}</text>')
    parts.append(f'<line x1="{ml}" y1="{py(PASS_LINE):.1f}" '
                 f'x2="{width - mr}" y2="{py(PASS_LINE):.1f}" '
                 f'stroke="#B45309" stroke-width="1.5" '
                 f'stroke-dasharray="5 4"/>')
    parts.append(f'<text x="{ml + 5}" y="{py(PASS_LINE) - 5:.1f}" '
                 f'text-anchor="start" font-size="10" fill="#B45309">'
                 f'합격선 {PASS_LINE}</text>')
    coords = [(px(i), py(s)) for i, (_lab, s) in enumerate(points)]
    poly = " ".join(f"{x:.1f},{y:.1f}" for x, y in coords)
    parts.append(f'<polyline points="{poly}" fill="none" stroke="#107C41" '
                 f'stroke-width="2.5"/>')
    for i, ((lab, s), (x, y)) in enumerate(zip(points, coords)):
        last = i == n - 1
        ring = 'stroke="#FFFFFF" stroke-width="1.5"' if last else ""
        parts.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" '
                     f'r="{5 if last else 3.5}" fill="#107C41" {ring}/>')
        parts.append(f'<text x="{x:.1f}" y="{y - 9:.1f}" text-anchor="middle" '
                     f'font-size="11" font-weight="{700 if last else 400}" '
                     f'fill="#1B3A26" '
                     f'style="font-variant-numeric:tabular-nums">{s}</text>')
        parts.append(f'<text x="{x:.1f}" y="{height - 12}" '
                     f'text-anchor="middle" font-size="10" fill="#57705F">'
                     f'{esc(lab)}</text>')
    return (f'<svg width="{width}" height="{height}" viewBox="0 0 {width} '
            f'{height}" role="img" aria-label="성적 추이">'
            + "".join(parts) + "</svg>")

def _svg_cat_bars(cat_losses):
    lab_w, bar_w, row_h, pad = 170, 320, 28, 8
    width = lab_w + bar_w + 80
    height = pad * 2 + row_h * len(cat_losses)
    peak = max((v for _c, v in cat_losses), default=1) or 1
    rows = []
    for i, (cat, lost) in enumerate(cat_losses):
        y = pad + i * row_h
        cy = y + row_h / 2 + 4
        w = bar_w * (lost / peak)
        name = CAT_INFO.get(cat, {}).get("이름", cat)
        rows.append(
            f'<text x="{lab_w - 8}" y="{cy}" text-anchor="end" font-size="12" '
            f'fill="#1B3A26">{esc(name)}</text>'
            f'<rect x="{lab_w}" y="{y + 6}" width="{bar_w}" '
            f'height="{row_h - 12}" rx="4" fill="#F4EFE6"/>'
            f'<rect x="{lab_w}" y="{y + 6}" width="{max(w, 3):.1f}" '
            f'height="{row_h - 12}" rx="4" fill="#B45309"/>'
            f'<text x="{lab_w + bar_w + 8}" y="{cy}" font-size="12" '
            f'font-weight="600" fill="#B45309" '
            f'style="font-variant-numeric:tabular-nums">-{lost:.0f}점</text>')
    return (f'<svg width="{width}" height="{height}" viewBox="0 0 {width} '
            f'{height}" role="img" aria-label="유형별 실점">'
            + "".join(rows) + "</svg>")

def _history_points(history, set_name, current_score):
    pts = []
    for rec in history or []:
        if rec.get("mode") == "부분연습":
            continue
        nm = str(rec.get("세트명", ""))
        if set_name and (nm == set_name or nm.startswith(set_name)):
            sc = rec.get("점수")
            if isinstance(sc, (int, float)):
                stamp = str(rec.get("일시", ""))
                lab = stamp[5:10].replace("-", "/") if len(stamp) >= 10 else ""
                pts.append((lab or "이전", int(sc)))
    pts = pts[-8:]
    pts.append(("이번", int(current_score)))
    return pts

SELF_GRADE_KEY_PREFIX = "kocoSelfGrade:"

SELF_GRADE_DISCLAIMER = "채점기에 오류가 있을 수 있습니다. 참고용으로 봐 주세요."

SELF_GRADE_DISCLAIMER_SUB = (
    "실제 시험 채점과 다를 수 있습니다. 채점기가 틀렸다고 본 항목은 아래 "
    "오답노트에서 직접 판정을 바꿀 수 있고, 점수는 바로 다시 계산됩니다."
)

def report_storage_id(student_path):
    rid = attempt_report_id(student_path)
    if rid:
        return rid
    base = os.path.basename(str(student_path or "")) or "report"
    return "f" + hashlib.md5(base.encode("utf-8")).hexdigest()[:12]

def self_grade_key(student_path):
    return SELF_GRADE_KEY_PREFIX + report_storage_id(student_path)

def self_grade_card_id(card, idx):
    raw = "%d|%s|%s|%s" % (idx, card.get("sheet", ""), card.get("label", ""),
                           card.get("kind", ""))
    return "c%d-%s" % (idx, hashlib.md5(raw.encode("utf-8")).hexdigest()[:8])

def self_grade_hold_id(hold, idx):
    raw = "%d|%s|%s" % (idx, hold.get("sheet", ""), hold.get("label", ""))
    return "h%d-%s" % (idx, hashlib.md5(raw.encode("utf-8")).hexdigest()[:8])

def self_grade_payload(results, score100, paths, partial=False):
    cards = [c for r in results for c in r.wrong]
    holds = hold_items(results)
    alloc_total = round(sum(r.alloc for r in results), 2)
    held_total = hold_total(results)
    earned_total = round(sum(r.earned for r in results), 2)
    return {
        "reportId": report_storage_id(paths[2]),
        "storageKey": self_grade_key(paths[2]),
        "setName": derive_set_name(paths[0]),
        "baseScore": score100,
        "partial": bool(partial),
        "maxTotal": (int(round(alloc_total)) if partial else 100),
        "earned": earned_total,
        "gradable": round(max(0.0, alloc_total - held_total), 2),
        "allocTotal": alloc_total,
        "heldTotal": held_total,
        "cards": [
            {"id": self_grade_card_id(c, i), "sheet": c.get("sheet"),
             "label": c.get("label"), "lost": round(float(c.get("lost") or 0), 2)}
            for i, c in enumerate(cards)
        ],
        "holds": [
            {"id": self_grade_hold_id(h, i), "sheet": h.get("sheet"),
             "label": h.get("label"),
             "points": round(float(h.get("points") or 0), 2)}
            for i, h in enumerate(holds)
        ],
    }

def js_json(obj):
    return (json.dumps(obj, ensure_ascii=False)
            .replace("<", "\\u003c").replace(">", "\\u003e")
            .replace("&", "\\u0026").replace("\u2028", "\\u2028")
            .replace("\u2029", "\\u2029"))

_SELF_GRADE_JS = r"""(function () {
  "use strict";
  var D = window.KOCO_SELF_GRADE_DATA;
  if (!D) { return; }
  var KEY = D.storageKey;
  var state = { cards: {}, holds: {} };

  function byId(id) {
    try { return document.getElementById(id); } catch (e) { return null; }
  }
  function qsa(sel, root) {
    try { return (root || document).querySelectorAll(sel) || []; }
    catch (e) { return []; }
  }
  function qs(root, sel) {
    try { return root && root.querySelector ? root.querySelector(sel) : null; }
    catch (e) { return null; }
  }
  function num(v) { var n = parseFloat(v); return isFinite(n) ? n : 0; }
  function pt(n) { return (Math.round(n * 100) / 100) + ""; }

  // 파이썬 round() 와 같은 자리맞춤(은행가 반올림) — 1차 점수와 어긋나지 않게
  function roundHalfEven(x) {
    var f = Math.floor(x), d = x - f;
    if (d > 0.5) { return f + 1; }
    if (d < 0.5) { return f; }
    return (f % 2 === 0) ? f : f + 1;
  }

  // 순수 계산 — DOM 없이도 돌아간다(단위 테스트용)
  function selfScore(data, st) {
    var earned = data.earned, gradable = data.gradable;
    var flips = 0, resolved = 0, i, c, h, v;
    for (i = 0; i < data.cards.length; i++) {
      c = data.cards[i];
      if (st.cards[c.id]) { earned += c.lost; flips++; }
    }
    for (i = 0; i < data.holds.length; i++) {
      h = data.holds[i];
      v = st.holds[h.id];
      if (v === "right") { earned += h.points; gradable += h.points; resolved++; }
      else if (v === "wrong") { gradable += h.points; resolved++; }
    }
    var score;
    if (!flips && !resolved) {
      score = data.baseScore;           // 손대지 않았으면 1차 점수 그대로
    } else if (data.partial) {
      score = roundHalfEven(earned);
    } else if (gradable > 0) {
      score = roundHalfEven(earned / gradable * 100);
    } else {
      score = 0;
    }
    if (score < 0) { score = 0; }
    if (score > data.maxTotal) { score = data.maxTotal; }
    return { score: score, earned: earned, gradable: gradable,
             flips: flips, resolved: resolved, changed: flips + resolved };
  }

  function load() {
    try {
      var raw = window.localStorage.getItem(KEY);
      if (!raw) { return; }
      var obj = JSON.parse(raw);
      if (!obj || typeof obj !== "object") { return; }
      var k;
      if (obj.cards && typeof obj.cards === "object") {
        for (k in obj.cards) {
          if (obj.cards[k] === true) { state.cards[k] = true; }
        }
      }
      if (obj.holds && typeof obj.holds === "object") {
        for (k in obj.holds) {
          if (obj.holds[k] === "right" || obj.holds[k] === "wrong") {
            state.holds[k] = obj.holds[k];
          }
        }
      }
    } catch (e) { /* 저장소를 못 쓰는 환경 — 화면은 그대로 돈다 */ }
  }

  function save() {
    try {
      var r = selfScore(D, state);
      window.localStorage.setItem(KEY, JSON.stringify({
        v: 1, reportId: D.reportId, set: D.setName,
        cards: state.cards, holds: state.holds,
        autoScore: D.baseScore, myScore: r.score,
        savedAt: new Date().toISOString()
      }));
    } catch (e) { /* 저장 실패는 조용히 넘어간다 */ }
  }

  function flash(msg) {
    var el = byId("sg-status");
    if (el) { el.textContent = msg; }
  }

  function render() {
    var r = selfScore(D, state), i, el, id, on, btn, st, lost;
    var mine = byId("sg-mine");
    if (mine) { mine.textContent = String(r.score); }
    var auto = byId("sg-auto");
    if (auto) { auto.textContent = String(D.baseScore); }
    var diff = byId("sg-diff");
    if (diff) {
      if (r.changed) {
        var d = r.score - D.baseScore;
        diff.textContent = "내가 바꾼 판정 " + r.changed + "건 · 채점기 대비 "
          + (d > 0 ? "+" : "") + d + "점";
        diff.removeAttribute("hidden");
      } else {
        diff.setAttribute("hidden", "");
        diff.textContent = "";
      }
    }
    var reset = byId("sg-reset");
    if (reset) {
      if (r.changed) { reset.removeAttribute("hidden"); }
      else { reset.setAttribute("hidden", ""); }
    }
    var cards = qsa("[data-sg-card]");
    for (i = 0; i < cards.length; i++) {
      el = cards[i];
      id = el.getAttribute("data-sg-card");
      lost = num(el.getAttribute("data-sg-lost"));
      on = state.cards[id] === true;
      el.setAttribute("data-sg-on", on ? "1" : "0");
      btn = qs(el, ".sg-toggle");
      if (btn) {
        btn.setAttribute("aria-pressed", on ? "true" : "false");
        btn.textContent = on ? "채점기 판정으로 되돌리기" : "내가 보기엔 맞음";
      }
      st = qs(el, ".sg-state");
      if (st) {
        st.textContent = on
          ? ("내 최종 판정: 맞음 — 감점 " + pt(lost) + "점을 되돌렸습니다")
          : ("채점기 1차 판정: 틀림" + (lost ? " (-" + pt(lost) + "점)" : ""));
      }
    }
    var holds = qsa("[data-sg-hold]");
    for (i = 0; i < holds.length; i++) {
      el = holds[i];
      id = el.getAttribute("data-sg-hold");
      lost = num(el.getAttribute("data-sg-points"));
      var v = state.holds[id] || "";
      el.setAttribute("data-sg-verdict", v);
      var opts = qsa(".sg-hold-btn", el);
      for (var j = 0; j < opts.length; j++) {
        var want = opts[j].getAttribute("data-sg-hold-val") || "";
        opts[j].setAttribute("aria-pressed", want === v ? "true" : "false");
      }
      st = qs(el, ".sg-hold-state");
      if (st) {
        st.textContent = v === "right"
          ? ("내 최종 판정: 맞음 — " + pt(lost) + "점을 득점으로 넣었습니다")
          : v === "wrong"
            ? ("내 최종 판정: 틀림 — " + pt(lost) + "점을 감점으로 넣었습니다")
            : "채점기가 판단하지 못한 항목 — 지금은 점수에서 뺀 상태입니다";
      }
    }
  }

  function setCard(id, on) {
    if (on) { state.cards[id] = true; } else { delete state.cards[id]; }
    render(); save();
  }
  function setHold(id, v) {
    if (v === "right" || v === "wrong") { state.holds[id] = v; }
    else { delete state.holds[id]; }
    render(); save();
  }
  function resetAll() {
    state.cards = {}; state.holds = {};
    try { window.localStorage.removeItem(KEY); } catch (e) { }
    render();
    flash("채점기의 1차 판정으로 모두 되돌렸습니다.");
  }

  function copyScore() {
    var r = selfScore(D, state);
    var txt = String(r.score);
    function ok() { flash("최종 점수 " + txt + "점을 복사했습니다."); }
    function fallback() {
      var inp = byId("sg-copy-src");
      if (inp) {
        inp.removeAttribute("hidden");
        inp.value = txt;
        try { inp.focus(); inp.select(); } catch (e) { }
        flash("자동 복사가 막혀 있습니다 — 칸에 있는 " + txt
              + " 을 Ctrl+C 로 복사하세요.");
      } else {
        flash("자동 복사가 막혀 있습니다 — 최종 점수는 " + txt + "점입니다.");
      }
    }
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(txt).then(ok, fallback);
      } else { fallback(); }
    } catch (e) { fallback(); }
  }

  function onClick(ev) {
    var t = ev.target;
    while (t && t !== document.body) {
      if (t.classList && t.classList.contains("sg-toggle")) {
        var card = t.closest ? t.closest("[data-sg-card]") : null;
        if (card) {
          var cid = card.getAttribute("data-sg-card");
          setCard(cid, state.cards[cid] !== true);
        }
        return;
      }
      if (t.classList && t.classList.contains("sg-hold-btn")) {
        var box = t.closest ? t.closest("[data-sg-hold]") : null;
        if (box) {
          setHold(box.getAttribute("data-sg-hold"),
                  t.getAttribute("data-sg-hold-val") || "");
        }
        return;
      }
      if (t.id === "sg-reset") { resetAll(); return; }
      if (t.id === "sg-copy") { copyScore(); return; }
      t = t.parentNode;
    }
  }

  load();
  render();
  try { document.addEventListener("click", onClick); } catch (e) { }

  window.kocoSelfGrade = {
    data: D, state: state, score: selfScore, roundHalfEven: roundHalfEven,
    setCard: setCard, setHold: setHold, reset: resetAll, render: render
  };
})();"""

def _wrong_card_html(card, idx=None, sg_id=None):
    cat = card.get("category")
    cat_chip = (f'<span class="chip">{esc(CAT_INFO.get(cat, {}).get("이름", cat))}'
                f'</span>') if cat else ""
    lost = card.get("lost") or 0
    lost_txt = f"-{lost:g}점" if lost else "감점 요인"
    head = (f'<div class="wc-top"><span class="wc-sheet">{esc(card["sheet"])}'
            f'</span><span class="wc-label">{esc(card["label"])}</span>'
            f'{cat_chip}<span class="wc-lost">{lost_txt}</span></div>')
    body = []
    cells = card.get("cells") or []
    chip_only = cells and all(c.get("expected") is None and c.get("got") is None
                              for c in cells)
    if chip_only:
        chips = ", ".join(f'<span class="cellchip">{esc(c["coord"])}</span>'
                          for c in cells)
        body.append(f'<div class="wc-chips">확인할 위치: {chips}'
                    + (f' 외 {card["more"]}곳' if card.get("more") else "")
                    + "</div>")
    else:
        for c in cells:
            got = c.get("got")
            gotf = c.get("got_formula")
            exp = c.get("expected")
            mine = esc(got if got is not None else "(비어 있음)")
            if gotf and str(gotf) != str(got):
                mine += f'<span class="sub">{esc(gotf)}</span>'
            ans = esc(exp if exp is not None else "(비어 있음)")
            coord = f'<span class="coord">{esc(c["coord"])}</span>' \
                if c.get("coord") else ""
            body.append(
                f'<div class="cmp">{coord}'
                f'<div class="box mine"><span class="tag">내 답</span>'
                f'<code>{mine}</code></div>'
                f'<div class="arrow">→</div>'
                f'<div class="box ans"><span class="tag">정답</span>'
                f'<code>{ans}</code></div></div>')
        if card.get("more"):
            body.append(f'<div class="wc-chips">외 {card["more"]}개 셀 동일 유형'
                        "</div>")
    if card.get("props"):
        rows = "".join(
            f'<div class="prop-row"><span class="prop-name">'
            f'{esc(p.get("name", "?"))}</span>'
            f'<span class="prop-exp">정답: {esc(p.get("expected", "?"))}</span>'
            f'<span class="prop-got">내 답: {esc(p.get("got", "?"))}</span>'
            "</div>"
            for p in card["props"][:8])
        body.append(f'<div class="props">{rows}</div>')
    if card.get("note"):
        body.append(f'<p class="wc-note">{esc(card["note"])}</p>')
    if card.get("diff_notes"):
        diffs = "".join(f'<div class="diff-item">{esc(d)}</div>'
                        for d in card["diff_notes"])
        body.append('<div class="diffbox"><div class="diff-title">무엇이 '
                    "다른가</div>" + diffs + "</div>")
    solve = []
    if card.get("formula"):
        solve.append(f'<code class="formula">{esc(card["formula"])}</code>')
    for step in card.get("explain") or []:
        solve.append(f'<div class="step">{esc(step)}</div>')
    if card.get("point"):
        solve.append(f'<p class="point">포인트: {esc(card["point"])}</p>')
    if solve:
        body.append('<div class="solve"><div class="solve-title">정확한 풀이'
                    "</div>" + "".join(solve) + "</div>")
    if card.get("hint"):
        body.append(f'<p class="hint">방법: {esc(card["hint"])}</p>')
    if sg_id and lost:
        body.append(
            '<div class="sg-row">'
            '<button type="button" class="sg-toggle" aria-pressed="false">'
            '내가 보기엔 맞음</button>'
            f'<span class="sg-state">채점기 1차 판정: 틀림'
            + (f' (-{lost:g}점)' if lost else "")
            + "</span></div>")
        return (f'<div class="wc" data-sg-card="{esc(sg_id)}" '
                f'data-sg-lost="{lost:g}" data-sg-on="0">'
                f'{head}{"".join(body)}</div>')
    if sg_id:
        body.append('<div class="sg-row"><span class="sg-state">'
                    "이 카드에는 되돌릴 감점이 없습니다 (설명용)</span></div>")
    return f'<div class="wc">{head}{"".join(body)}</div>'

def write_html(path, results, score100, global_notes, paths,
               set_name=None, history=None, partial=False):
    verdict = "합격권" if score100 >= PASS_LINE else "미달"
    vcls = "pass" if score100 >= PASS_LINE else "fail"
    set_name = set_name or derive_set_name(paths[0])
    _now = datetime.now()
    now_txt = _now.strftime("%Y-%m-%d %H:%M")
    rid = attempt_report_id(paths[2]) or _now.strftime("%Y%m%d%H%M%S")

    diff_total = sum(getattr(r, "diff_total", 0) or 0 for r in results)
    diff_matched = sum(getattr(r, "diff_matched", 0) or 0 for r in results)
    warn_html = ""
    if diff_total > 0 and diff_matched / diff_total < 0.10:
        warn_html = ('<div class="warn-banner">풀이가 거의 비어 있습니다 — '
                     'Excel에서 저장(Ctrl+S)했는지, 풀이 사본 파일이 맞는지 '
                     '확인하세요</div>')
    contam = [c for r in results for c in (getattr(r, "contaminated", None) or [])]
    if contam:
        warn_html += ('<div class="warn-banner">문제 파일 오염 의심 — 문제 파일의 '
                      '계산작업에 수식이 이미 들어 있습니다 ('
                      + esc(", ".join(contam[:4]))
                      + (" 외" if len(contam) > 4 else "")
                      + '). 이전 풀이로 덮어써진 파일이면 문제 수·배점과 다른 '
                      '시트의 채점 기준이 실제와 달라집니다. 원본 문제 파일을 '
                      '다시 받아 채점하세요</div>')

    all_cards = [c for r in results for c in r.wrong]
    cat_losses, top3, checklist = build_diagnosis(results)
    max_total = sum(r.alloc for r in results) if partial else 100
    graded_names = ", ".join(r.name for r in results)

    trend_html = ""
    if history and not partial:
        pts = _history_points(history, set_name, score100)
        if len(pts) >= 2:
            trend_html = ('<div class="card"><h2>성적 추이</h2>'
                          '<div class="scroll">' + _svg_trend(pts)
                          + "</div></div>")

    sg_card_ids = [self_grade_card_id(c, i) for i, c in enumerate(all_cards)]
    if all_cards:
        wrong_html = ('<div class="card"><h2>오답노트</h2>'
                      '<p class="lead">틀린 항목마다 내 답과 정답을 비교하고, '
                      '정답 수식의 풀이를 단계별로 설명합니다.</p>'
                      '<p class="lead sg-lead">아래는 <strong>채점기가 '
                      "'틀림'으로 본 1차 판정</strong>입니다. 문제지와 내 답을 "
                      '다시 보고 <strong>내가 보기엔 맞다</strong>면 항목마다 '
                      '[내가 보기엔 맞음]을 눌러 판정을 뒤집으세요 — 그 항목의 '
                      '감점이 취소되고 위의 최종 점수가 바로 다시 '
                      '계산됩니다.</p>'
                      + "".join(_wrong_card_html(c, i, sg_id=sg_card_ids[i])
                                for i, c in enumerate(all_cards))
                      + "</div>")
    elif any(getattr(r, "holds", None) for r in results):
        wrong_html = ('<div class="card"><h2>오답노트</h2>'
                      '<p class="perfect">틀린 항목이 없습니다 — 다만 아래 '
                      '[확인 필요] 항목은 판정하지 못해 채점에서 뺐으니 직접 '
                      '확인해 주세요.</p></div>')
    else:
        wrong_html = ('<div class="card"><h2>오답노트</h2>'
                      '<p class="perfect">틀린 항목이 없습니다. 모든 채점 '
                      '항목을 통과했습니다.</p></div>')

    body_cls = ""

    if cat_losses:
        tops = []
        for i, t in enumerate(top3, 1):
            items = "".join(f"<li>{esc(p)}</li>" for p in t["처방"])
            tops.append(
                f'<div class="weak"><div class="rank">{i}</div>'
                f'<div class="weak-body"><h4>{esc(t["이름"])} '
                f'<span class="weak-lost">-{t["lost"]:g}점</span></h4>'
                f'<p>{esc(t["진단"])}</p><ul>{items}</ul></div></div>')
        diag_html = ('<div class="card"><h2>취약점 진단</h2>'
                     '<div class="scroll">' + _svg_cat_bars(cat_losses)
                     + '</div><h3 class="sub-h">취약 TOP '
                     + str(len(top3)) + "</h3>" + "".join(tops) + "</div>")
    else:
        diag_html = ""

    if not checklist:
        checklist = ["내일 다른 회차 모의고사 1세트로 컨디션 유지",
                     "헷갈렸던 함수 1개를 골라 치트시트 훑어보기"]
    check_html = ('<div class="card"><h2>다음 액션</h2><ul class="checklist">'
                  + "".join(f'<li><label><input type="checkbox"> {esc(c)}'
                            "</label></li>" for c in checklist)
                  + "</ul></div>")

    holds = hold_items(results)
    held_total = hold_total(results)
    alloc_total = sum(r.alloc for r in results)
    gradable_total = max(0.0, alloc_total - held_total)
    review_html = ""
    score_note = ""
    sg_hold_ids = [self_grade_hold_id(h, i) for i, h in enumerate(holds)]
    if holds:
        rows = "".join(
            f'<div class="review-item" data-sg-hold="{esc(sg_hold_ids[i])}" '
            f'data-sg-points="{h["points"]:g}" data-sg-verdict="">'
            f'<div class="review-head">{esc(h["sheet"])} · '
            f'{esc(h["label"])}'
            f'<span class="review-pts">{h["points"]:g}점 분량 · 점수에서 '
            f'뺌</span></div>'
            f'<div class="review-why">{esc(h["why"])}</div>'
            '<div class="sg-hold-row">'
            '<span class="sg-hold-lead">직접 판정:</span>'
            '<button type="button" class="sg-hold-btn" '
            'data-sg-hold-val="right" aria-pressed="false">맞음</button>'
            '<button type="button" class="sg-hold-btn" '
            'data-sg-hold-val="wrong" aria-pressed="false">틀림</button>'
            '<button type="button" class="sg-hold-btn" data-sg-hold-val="" '
            'aria-pressed="true">채점 제외로 두기</button>'
            '<span class="sg-hold-state">채점기가 판단하지 못한 항목 — '
            '지금은 점수에서 뺀 상태입니다</span>'
            "</div></div>"
            for i, h in enumerate(holds))
        review_html = (
            '<div class="card review-card"><h2>확인 필요 '
            f'{len(holds)}건</h2>'
            '<p class="lead">아래 항목은 파일에서 판정할 근거를 읽어 내지 '
            '못했습니다. <strong>맞았는지 틀렸는지 모른다</strong>는 뜻이라 '
            '감점도 가점도 하지 않고 만점에서 뺐습니다 — 직접 확인해 '
            '주세요.</p>'
            '<p class="lead sg-lead">이 항목들은 오답노트의 <strong>자가 '
            '채점과는 다른 칸</strong>입니다. 오답노트의 자가 채점은 '
            "<strong>채점기가 '틀림'으로 판정한 것을 사용자가 뒤집는</strong> "
            '것이고, 여기 확인 필요는 <strong>채점기가 아예 판단하지 '
            '못한</strong> 항목입니다. 직접 [맞음]·[틀림]을 고르면 그 배점이 '
            '다시 채점 대상으로 들어가고, 그냥 두면 지금처럼 점수에서 '
            '빠진 채로 남습니다.</p>'
            + rows
            + '<p class="review-calc">점수 계산: 득점 '
            f'{sum(r.earned for r in results):g} ÷ 채점된 배점 '
            f'{gradable_total:g}(= 전체 {alloc_total:g} − 확인 필요 '
            f'{held_total:g}) × 100 = <strong>{score100}점</strong></p>'
            "</div>")
        score_note = (f'<div class="review-flag">확인 필요 {len(holds)}건 '
                      f'({held_total:g}점 분량은 채점에서 제외)</div>')

    disclaimer_html = (
        '<div class="disclaimer" role="note">'
        f"<strong>{esc(SELF_GRADE_DISCLAIMER)}</strong>"
        f"<span>{esc(SELF_GRADE_DISCLAIMER_SUB)}</span></div>")
    sg_data = self_grade_payload(results, score100, paths, partial=partial)
    final_html = (
        '<div class="sg-final" id="sg-final">'
        '<div class="sg-scores">'
        f'<div class="sg-line sg-line-auto">채점기 <strong id="sg-auto">'
        f'{score100}</strong>점<span class="sg-tag">1차</span></div>'
        f'<div class="sg-line sg-line-mine">내가 매긴 점수 '
        f'<strong id="sg-mine">{score100}</strong>점'
        '<span class="sg-tag sg-tag-final">최종</span></div>'
        '<div class="sg-diff" id="sg-diff" hidden></div>'
        '</div>'
        '<div class="sg-acts">'
        '<button type="button" class="sg-btn sg-btn-main" id="sg-copy">'
        '점수 복사</button>'
        '<button type="button" class="sg-btn" id="sg-reset" hidden>'
        '전체 되돌리기</button>'
        '<input type="text" id="sg-copy-src" readonly hidden>'
        '<div class="sg-status" id="sg-status"></div>'
        '</div></div>')
    sg_script = ('<script>window.KOCO_SELF_GRADE_DATA='
                 + js_json(sg_data) + ";\n" + _SELF_GRADE_JS + "</script>")

    gnotes = "".join(f"<li>{esc(n)}</li>" for n in global_notes)
    sheet_notes = []
    for r in results:
        for nt in r.notes:
            sheet_notes.append(f"<li>[{esc(r.name)}] {esc(nt)}</li>")
    notes_html = ""
    if gnotes or sheet_notes:
        notes_html = ('<div class="card"><h2>채점 참고</h2><ul class="notes">'
                      + "".join(sheet_notes) + gnotes + "</ul></div>")

    css = """
* { box-sizing:border-box; margin:0; padding:0; }
body { background:#F4FAF5; color:#1B3A26;
  font-family:'Malgun Gothic','맑은 고딕',Pretendard,sans-serif;
  line-height:1.65; padding:28px 14px; }
.wrap { max-width:820px; margin:0 auto; }
.card { background:#FFFFFF; border:1px solid #D7E3DA; border-radius:14px;
  padding:26px 30px; margin-bottom:18px; }
h1 { font-size:1.3rem; color:#107C41; }
h2 { font-size:1.06rem; color:#107C41; margin-bottom:12px;
  padding-bottom:8px; border-bottom:1px solid #E8EFEA; }
.sub-h { font-size:0.98rem; margin:16px 0 10px; color:#1B3A26; }
.meta { color:#57705F; font-size:0.84rem; margin-top:2px; }
.head-flex { display:flex; gap:26px; align-items:center; flex-wrap:wrap; }
.head-info { flex:1; min-width:230px; }
.badge { display:inline-block; margin-top:10px; padding:5px 16px;
  border-radius:999px; font-weight:700; font-size:0.92rem; color:#fff; }
.badge.pass { background:#107C41; }
.badge.fail { background:#B3372E; }
.files { color:#57705F; font-size:0.8rem; margin-top:12px;
  word-break:break-all; }
.file-solve { display:inline-block; margin-top:3px; padding:2px 10px;
  background:#E3F2E8; border:1px solid #BFDCCB; border-radius:6px;
  color:#0B5D31; font-size:0.84rem; }
.file-solve strong { font-weight:700; }
.warn-banner { margin-top:14px; padding:10px 14px; border-radius:10px;
  background:#FDECEA; border:1px solid #EFB9B2; color:#B3372E;
  font-weight:700; font-size:0.9rem; }
.scroll { overflow-x:auto; }
svg { font-family:'Malgun Gothic','맑은 고딕',sans-serif; display:block; }
.lead { color:#57705F; font-size:0.86rem; margin-bottom:14px; }
.perfect { color:#107C41; font-weight:700; }
.wc { border:1px solid #D7E3DA; border-left:4px solid #B45309;
  border-radius:10px; padding:14px 16px; margin-bottom:14px; }
.wc-top { display:flex; align-items:center; gap:8px; flex-wrap:wrap;
  margin-bottom:10px; }
.wc-sheet { color:#57705F; font-size:0.8rem; }
.wc-label { font-weight:700; font-size:0.95rem; }
.chip { background:#E3F2E8; color:#0B5D31; font-size:0.74rem;
  padding:2px 10px; border-radius:999px; }
.wc-lost { margin-left:auto; color:#B45309; font-weight:700;
  font-variant-numeric:tabular-nums; }
.cmp { display:flex; align-items:stretch; gap:8px; margin-bottom:8px;
  flex-wrap:wrap; }
.coord { align-self:center; min-width:44px; font-weight:700;
  font-size:0.85rem; color:#1B3A26; font-variant-numeric:tabular-nums; }
.box { flex:1; min-width:180px; border-radius:8px; padding:7px 10px;
  border:1px solid #D7E3DA; }
.box.mine { background:#FBF6EE; border-color:#E5D5B8; }
.box.ans { background:#F0F7F2; border-color:#BFDCCB; }
.box .tag { display:block; font-size:0.7rem; color:#57705F;
  margin-bottom:2px; }
.box code { font-family:Consolas,'Courier New',monospace; font-size:0.84rem;
  word-break:break-all; }
.box .sub { display:block; font-size:0.76rem; color:#8A7B5E;
  font-family:Consolas,monospace; word-break:break-all; }
.arrow { align-self:center; color:#57705F; }
.wc-chips { font-size:0.84rem; color:#57705F; margin-bottom:8px; }
.cellchip { display:inline-block; background:#ECF1ED; border-radius:6px;
  padding:1px 8px; margin:0 3px 3px 0; font-family:Consolas,monospace;
  font-size:0.8rem; }
.wc-note { font-size:0.85rem; color:#1B3A26; background:#F4FAF5;
  border-radius:8px; padding:8px 12px; margin-bottom:8px; }
.props { border:1px solid #D7E3DA; border-radius:8px; padding:8px 12px;
  margin-bottom:8px; }
.prop-row { display:flex; gap:12px; flex-wrap:wrap; font-size:0.86rem;
  padding:3px 0; border-bottom:1px dashed #E8EFEA; }
.prop-row:last-child { border-bottom:none; }
.prop-name { font-weight:700; min-width:88px; }
.prop-exp { color:#0B5D31; }
.prop-got { color:#B45309; }
.diffbox { background:#FBF6EE; border:1px solid #E5D5B8; border-radius:8px;
  padding:10px 14px; margin-bottom:8px; }
.diff-title { font-weight:700; font-size:0.84rem; color:#8A5A00;
  margin-bottom:6px; }
.diff-item { font-size:0.86rem; margin-bottom:4px; padding-left:4px; }
.solve { background:#F7FBF8; border:1px solid #E1EDE5; border-radius:8px;
  padding:10px 14px; margin-bottom:8px; }
.solve-title { font-weight:700; font-size:0.84rem; color:#0B5D31;
  margin-bottom:6px; }
.formula { display:block; font-family:Consolas,'Courier New',monospace;
  font-size:0.85rem; background:#FFFFFF; border:1px solid #D7E3DA;
  border-radius:6px; padding:6px 10px; margin-bottom:8px;
  word-break:break-all; }
.step { font-size:0.86rem; margin-bottom:4px; padding-left:4px; }
.point { font-size:0.82rem; color:#0B5D31; background:#E3F2E8;
  border-radius:6px; padding:6px 10px; margin-top:8px; }
.hint { font-size:0.82rem; color:#57705F; }
.weak { display:flex; gap:14px; border:1px solid #D7E3DA; border-radius:10px;
  padding:14px 16px; margin-bottom:10px; }
.rank { width:30px; height:30px; border-radius:50%; background:#B45309;
  color:#fff; font-weight:700; display:flex; align-items:center;
  justify-content:center; flex:none; font-variant-numeric:tabular-nums; }
.weak-body h4 { font-size:0.95rem; margin-bottom:4px; }
.weak-lost { color:#B45309; font-size:0.84rem;
  font-variant-numeric:tabular-nums; }
.weak-body p { font-size:0.86rem; color:#1B3A26; margin-bottom:6px; }
.weak-body ul { padding-left:18px; font-size:0.84rem; color:#57705F; }
.weak-body li { margin-bottom:2px; }
.checklist { list-style:none; }
.checklist li { margin-bottom:8px; font-size:0.9rem; }
.checklist input { accent-color:#107C41; margin-right:8px;
  transform:scale(1.15); }
.notes { padding-left:18px; font-size:0.82rem; color:#57705F; }
.notes li { margin-bottom:3px; }
footer { color:#57705F; font-size:0.78rem; text-align:center;
  margin-top:6px; }
[hidden] { display:none !important; }
.review-card { border-color:#E0B96A; background:#FFFBF2; }
.review-card h2 { color:#9A6B10; border-bottom-color:#F0DFBC; }
.review-item { border-left:3px solid #E0B96A; padding:8px 0 8px 12px;
  margin:10px 0; }
.review-head { font-weight:700; font-size:0.94rem; }
.review-pts { display:inline-block; margin-left:8px; font-weight:400;
  font-size:0.8rem; color:#9A6B10; }
.review-why { font-size:0.88rem; color:#4A5A50; margin-top:3px; }
.review-calc { margin-top:14px; font-size:0.86rem; color:#57705F;
  border-top:1px dashed #F0DFBC; padding-top:10px; }
.review-flag { display:inline-block; margin-top:8px; padding:3px 12px;
  border:1px solid #E0B96A; border-radius:999px; background:#FFFBF2;
  color:#9A6B10; font-size:0.82rem; }
/* --- v2.2.0 면책 문구 · 자가 채점 --- */
.donut-wrap { text-align:center; }
.donut-cap { font-size:0.78rem; color:#57705F; margin-top:2px; }
.disclaimer { margin-top:16px; padding:12px 16px; border-radius:10px;
  background:#FFF7E6; border:1px solid #E0B96A; color:#7A4E07; }
.disclaimer strong { display:block; font-size:0.96rem; color:#8A5A00; }
.disclaimer span { display:block; font-size:0.84rem; margin-top:3px;
  color:#6B5A3A; }
.sg-final { margin-top:14px; padding:14px 18px; border-radius:12px;
  border:1px solid #BFDCCB; background:#F4FAF5; display:flex; gap:18px;
  flex-wrap:wrap; align-items:center; }
.sg-scores { flex:1; min-width:230px; }
.sg-line { font-size:0.95rem; color:#57705F; }
.sg-line strong { font-variant-numeric:tabular-nums; }
.sg-line-auto strong { font-size:1.15rem; color:#1B3A26; }
.sg-line-mine { margin-top:2px; color:#0B5D31; }
.sg-line-mine strong { font-size:1.9rem; color:#107C41; }
.sg-tag { display:inline-block; margin-left:6px; padding:1px 9px;
  border-radius:999px; font-size:0.72rem; background:#ECF1ED;
  color:#57705F; vertical-align:middle; }
.sg-tag-final { background:#107C41; color:#fff; }
.sg-diff { margin-top:4px; font-size:0.82rem; color:#B45309;
  font-weight:700; }
.sg-acts { display:flex; flex-direction:column; gap:6px; align-items:stretch;
  min-width:150px; }
.sg-btn { font:inherit; font-size:0.86rem; padding:7px 14px;
  border-radius:8px; border:1px solid #BFDCCB; background:#FFFFFF;
  color:#0B5D31; cursor:pointer; }
.sg-btn:hover { background:#E3F2E8; }
.sg-btn-main { background:#107C41; border-color:#107C41; color:#fff;
  font-weight:700; }
.sg-btn-main:hover { background:#0B5D31; }
.sg-status { font-size:0.78rem; color:#57705F; min-height:1.2em; }
#sg-copy-src { font:inherit; font-size:0.9rem; width:100%; padding:4px 8px;
  border:1px solid #BFDCCB; border-radius:6px; text-align:center; }
.sg-lead { border-left:3px solid #BFDCCB; padding-left:10px; }
.sg-row { display:flex; gap:10px; align-items:center; flex-wrap:wrap;
  margin-top:10px; padding-top:9px; border-top:1px dashed #E8EFEA; }
.sg-toggle { font:inherit; font-size:0.82rem; padding:5px 13px;
  border-radius:999px; border:1px solid #BFDCCB; background:#FFFFFF;
  color:#0B5D31; cursor:pointer; flex:none; }
.sg-toggle:hover { background:#E3F2E8; }
.sg-state { font-size:0.8rem; color:#57705F; }
.wc[data-sg-on="1"] { border-left-color:#107C41; background:#F7FBF8; }
.wc[data-sg-on="1"] .wc-lost { color:#57705F; text-decoration:line-through; }
.wc[data-sg-on="1"] .sg-toggle { background:#107C41; border-color:#107C41;
  color:#fff; }
.wc[data-sg-on="1"] .sg-state { color:#0B5D31; font-weight:700; }
.sg-hold-row { display:flex; gap:8px; align-items:center; flex-wrap:wrap;
  margin-top:9px; }
.sg-hold-lead { font-size:0.8rem; color:#9A6B10; font-weight:700; }
.sg-hold-btn { font:inherit; font-size:0.8rem; padding:4px 12px;
  border-radius:999px; border:1px solid #E0B96A; background:#FFFFFF;
  color:#7A4E07; cursor:pointer; }
.sg-hold-btn:hover { background:#FFF3DC; }
.sg-hold-btn[aria-pressed="true"] { background:#9A6B10; border-color:#9A6B10;
  color:#fff; font-weight:700; }
.sg-hold-state { font-size:0.8rem; color:#7A4E07; }
.review-item[data-sg-verdict="right"] { border-left-color:#107C41; }
.review-item[data-sg-verdict="wrong"] { border-left-color:#B3372E; }
"""
    doc = (
        "<!doctype html>\n"
        '<html lang="ko"><head><meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1">\n'
        f"<title>학습 리포트 - {esc(set_name)}"
        + (f" ({score100}점 · 확인 필요 {len(holds)}건)" if holds else "")
        + "</title>\n"
        f"<style>{css}</style></head><body{body_cls}>"
        + '<div class="wrap">\n'
        '<div class="card"><div class="head-flex">'
        f'<div class="donut-wrap">'
        f'{_svg_donut(score100, max_total=max_total, show_pass=not partial)}'
        f'<div class="donut-cap">채점기 1차 판정</div></div>'
        '<div class="head-info">'
        + (f"<h1>{esc(set_name)} 부분 연습 리포트</h1>"
           f'<div class="meta">부분 채점: {esc(graded_names)} '
           f'({max_total:.0f}점 만점) · 채점 일시 {now_txt}</div>' + score_note
           if partial else
           f"<h1>{esc(set_name)} 학습 리포트</h1>"
           f'<div class="meta">채점 일시 {now_txt} · 합격선 {PASS_LINE}점</div>'
           f'<span class="badge {vcls}">{verdict}</span>' + score_note)
        + '<div class="files">'
        f"문제 {esc(os.path.basename(paths[0]))}<br>"
        f"정답 {esc(os.path.basename(paths[1]))}<br>"
        f'<span class="file-solve">풀이 '
        f"<strong>{esc(os.path.basename(paths[2]))}</strong></span></div>"
        "</div></div>" + warn_html + disclaimer_html + final_html + "</div>\n"
        '<div class="card"><h2>영역별 점수</h2><div class="scroll">'
        + _svg_sheet_bars(results) + "</div></div>\n"
        + trend_html + review_html + wrong_html + diag_html + check_html
        + notes_html
        + f"<footer>코코 채점 학습 리포트 · diff 기반 자동 채점 · {now_txt}"
          f"<br>{esc(SELF_GRADE_DISCLAIMER)} 최종 판정은 사용자가 합니다."
          "</footer>\n</div>"
        + sg_script + "</body></html>")
    with open(path, "w", encoding="utf-8") as f:
        f.write(doc)

def derive_set_name(problem_path):
    base = os.path.splitext(os.path.basename(str(problem_path)))[0]
    for suf in ("_문제", "-문제", " 문제"):
        if base.endswith(suf):
            return base[: -len(suf)]
    return base

def write_json(path, results, score100, global_notes, paths, partial=False):
    holds = hold_items(results)
    held = hold_total(results)
    data = {
        "total": score100,
        "pass_line": PASS_LINE,
        "passed": None if partial else score100 >= PASS_LINE,
        "mode": "partial" if partial else "full",
        "graded_sheets": [r.name for r in results],
        "max_total": round(sum(r.alloc for r in results), 1)
        if partial else 100,
        "review_count": len(holds),
        "review_points": held,
        "review_items": [
            {"sheet": h["sheet"], "label": h["label"], "why": h["why"],
             "points": h["points"], "what": h.get("what")}
            for h in holds
        ],
        "scoring": {
            "alloc_total": round(sum(r.alloc for r in results), 2),
            "held_total": held,
            "gradable_total": round(
                max(0.0, sum(r.alloc for r in results) - held), 2),
            "earned_total": round(sum(r.earned for r in results), 2),
            "method": ("확인 필요 항목의 배점은 만점에서 빼고 나머지로 "
                       "비율 환산합니다 (점수 = 득점 ÷ 채점된 배점 × 100). "
                       "확인 필요는 감점도 가점도 아닙니다."),
        },
        "self_grading": dict(
            self_grade_payload(results, score100, paths, partial=partial),
            keyPrefix=SELF_GRADE_KEY_PREFIX,
            disclaimer=SELF_GRADE_DISCLAIMER,
            method=("채점기의 판정은 1차 의견입니다. 리포트(HTML)에서 오답 "
                    "항목의 판정을 뒤집으면 그 감점이 취소되고, '확인 필요' "
                    "항목을 직접 판정하면 그 배점이 채점 대상으로 돌아옵니다. "
                    "최종 점수 = (득점 + 되돌린 감점 + 맞음으로 판정한 보류) "
                    "÷ (채점된 배점 + 직접 판정한 보류) × 100. 사용자의 선택은 "
                    "브라우저 localStorage 의 리포트별 키에 저장되며 이 JSON "
                    "에는 들어오지 않습니다 — 여기 값은 모두 1차 판정입니다."),
        ),
        "generated": datetime.now().isoformat(timespec="seconds"),
        "files": {
            "problem": os.path.abspath(paths[0]),
            "answer": os.path.abspath(paths[1]),
            "student": os.path.abspath(paths[2]),
        },
        "sheets": [
            {
                "name": r.name,
                "alloc": round(r.alloc, 2),
                "earned": round(r.earned, 2),
                "missing": r.missing,
                "details": r.details,
                "notes": r.notes,
                "held": round(getattr(r, "held", 0), 2),
                "gradable": round(r.alloc - getattr(r, "held", 0), 2),
                "review_items": list(getattr(r, "holds", [])),
            }
            for r in results
        ],
        "notes": global_notes,
        "wrong_items": [
            {
                "sheet": card["sheet"], "label": card["label"],
                "lost": card["lost"], "kind": card["kind"],
                "category": card.get("category"),
                "cells": card.get("cells") or [],
                "formula": card.get("formula"),
                "note": card.get("note"), "hint": card.get("hint"),
                "explain": card.get("explain") or [],
                "point": card.get("point"),
                "student_formula": card.get("student_formula"),
                "expected_formula": card.get("expected_formula"),
                "diff_notes": card.get("diff_notes") or [],
                "props": card.get("props") or [],
            }
            for r in results for card in r.wrong
        ],
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def main(argv=None):
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
            sys.stderr.reconfigure(encoding="utf-8")
        except Exception:
            pass
    ap = argparse.ArgumentParser(
        description="컴활 2급 실기 자동 채점 (문제↔정답 diff 기반)")
    ap.add_argument("--problem", required=True, help="문제 파일(.xlsx/.xlsm)")
    ap.add_argument("--answer", required=True, help="정답 파일(.xlsx/.xlsm)")
    ap.add_argument("--student", required=True, help="학생 풀이 파일(.xlsx/.xlsm)")
    ap.add_argument("--key", help="기대값/배점/그룹 JSON (선택)")
    ap.add_argument("--html", help="HTML 리포트 저장 경로 (선택)")
    ap.add_argument("--json", dest="json_out",
                    help="채점 결과(총점·시트별 점수) JSON 저장 경로 (선택)")
    ap.add_argument("--history",
                    help="응시 기록 JSON(기록.json) 경로 - 있으면 HTML "
                         "리포트에 같은 세트의 성적 추이를 표시 (선택)")
    ap.add_argument("--sheets",
                    help="부분 채점: 지정 시트만 채점 (쉼표 구분, 예: "
                         "--sheets 계산작업 또는 --sheets 기본작업-1,"
                         "기본작업-2). 총점은 해당 시트 배점 합 기준 (선택)")
    ap.add_argument("--no-recalc", dest="no_recalc", action="store_true",
                    help="계산값이 없는 수식을 LibreOffice로 다시 계산하지 "
                         "않습니다 (기본은 재계산해 결과값으로 비교)")
    args = ap.parse_args(argv)
    if getattr(args, "no_recalc", False):
        global RECALC_ENABLED
        RECALC_ENABLED = False

    key = {}
    if args.key:
        try:
            with open(args.key, encoding="utf-8") as f:
                key = json.load(f)
        except Exception as e:
            print(f"경고: --key 파일을 읽을 수 없어 무시합니다: {e}")
    key = normalize_key(key)

    sheets = None
    if args.sheets:
        sheets = [s.strip() for s in str(args.sheets).split(",") if s.strip()]
    try:
        results, score100, notes, _books = run_grading(
            args.problem, args.answer, args.student, key, sheets=sheets)
    except (FileNotFoundError, RuntimeError) as e:
        print(f"오류: {e}")
        return 2

    partial = bool(sheets)
    paths = (args.problem, args.answer, args.student)
    print_console(results, score100, notes, paths, partial=partial)
    set_name = derive_set_name(args.problem)
    html_path, json_path = args.html, args.json_out
    if not html_path or not json_path:
        out_dir = os.path.join(
            os.path.dirname(os.path.abspath(args.student)), "채점결과")
        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        try:
            os.makedirs(out_dir, exist_ok=True)
            base = os.path.join(out_dir, f"채점결과_{set_name}_{stamp}")
            if not html_path:
                html_path = base + ".html"
            if not json_path:
                json_path = base + ".json"
        except OSError as e:
            print(f"경고: 채점결과 폴더를 만들 수 없어 리포트 저장을 "
                  f"생략합니다: {e}")
    history = None
    if args.history:
        try:
            with open(args.history, encoding="utf-8") as f:
                loaded = json.load(f)
            history = loaded if isinstance(loaded, list) else None
        except Exception as e:
            print(f"경고: --history 파일을 읽을 수 없어 무시합니다: {e}")
    if html_path:
        write_html(html_path, results, score100, notes, paths,
                   set_name=set_name, history=history, partial=partial)
        print(f"HTML 리포트를 저장했습니다: {html_path}")
    if json_path:
        write_json(json_path, results, score100, notes, paths,
                   partial=partial)
        print(f"JSON 결과를 저장했습니다: {json_path}")
    return 0

if __name__ == "__main__":
    sys.exit(main())
