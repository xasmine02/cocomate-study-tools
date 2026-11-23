#!/usr/bin/env python3
# -*- coding: utf-8 -*-

__version__ = "3.2.1"

import argparse
import hashlib
import http.server
import json
import os
import platform
import queue
import re
import secrets
import shutil
import subprocess
import sys
import threading
import time
import urllib.parse
import webbrowser
from datetime import date, datetime, timedelta

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    HAS_TK = True
except ImportError:
    HAS_TK = False

APP_TITLE = "코코 시험장"
BRAND = "#107C41"
BRAND_DARK = "#0B5D31"
BRAND_SOFT = "#E3F2E8"
AMBER = "#B45309"
RED = "#B3372E"
BG = "#F4FAF5"
CARD = "#FFFFFF"
INK = "#1B3A26"
SUB = "#57705F"
LINE = "#CFE3D5"

PASS_LINE = 70
DEFAULT_MINUTES = 40
MIN_MINUTES, MAX_MINUTES = 10, 60

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RECORDS_DIR = os.path.join(BASE_DIR, "채점결과")
RECORDS_PATH = os.path.join(RECORDS_DIR, "기록.json")
_OLD_RECORDS_PATH = os.path.join(BASE_DIR, "기록.json")
ERROR_LOG_PATH = os.path.join(RECORDS_DIR, "시험장_오류.log")
STARTUP_LOG_PATH = os.path.join(RECORDS_DIR, "시험장_시작로그.txt")
STARTUP_LOG_KEEP = 200

def _ensure_records_home():
    try:
        os.makedirs(RECORDS_DIR, exist_ok=True)
        if os.path.isfile(_OLD_RECORDS_PATH) \
                and not os.path.isfile(RECORDS_PATH):
            shutil.move(_OLD_RECORDS_PATH, RECORDS_PATH)
    except OSError:
        pass

_LOG_LOCK = threading.RLock()

def startup_log(message, path=None, keep=STARTUP_LOG_KEEP):
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{stamp}] {__version__} " + str(message).replace("\n", " | ")
    p = path or STARTUP_LOG_PATH
    try:
        with _LOG_LOCK:
            os.makedirs(os.path.dirname(p), exist_ok=True)
            old = []
            if os.path.isfile(p):
                with open(p, encoding="utf-8", errors="replace") as f:
                    old = f.read().splitlines()
            old.append(line)
            with open(p, "w", encoding="utf-8") as f:
                f.write("\n".join(old[-keep:]) + "\n")
    except Exception:
        pass
    return line

def read_startup_log(path=None):
    p = path or STARTUP_LOG_PATH
    try:
        with open(p, encoding="utf-8", errors="replace") as f:
            body = f.read().strip()
    except OSError:
        body = ""
    head = (f"[코코 시험장 시작 로그] {__version__} · "
            f"{datetime.now().strftime('%Y-%m-%d %H:%M')}\n파일: {p}\n")
    return head + "\n" + (body or "(아직 기록이 없습니다)")

def log_error(context, exc=None, path=None):
    import traceback
    if isinstance(exc, tuple) and len(exc) == 3:
        detail = "".join(traceback.format_exception(*exc))
    elif isinstance(exc, BaseException):
        detail = "".join(traceback.format_exception(
            type(exc), exc, exc.__traceback__))
    else:
        detail = traceback.format_exc()
        if detail.strip() == "NoneType: None":
            detail = ""
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    text = f"[{stamp}] {__version__} {context}\n{detail}".rstrip() + "\n"
    p = path or ERROR_LOG_PATH
    try:
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, "a", encoding="utf-8") as f:
            f.write(text + "-" * 60 + "\n")
    except Exception:
        pass
    if path is None:
        last = detail.strip().splitlines()[-1] if detail.strip() else ""
        startup_log(f"예외: {context}" + (f" — {last}" if last else ""))
    return text

UI_FONT = ("Malgun Gothic", 10)
UI_FONT_BOLD = ("Malgun Gothic", 10, "bold")
DIGIT_FONT = ("Consolas", 44, "bold")
OVERLAY_DIGIT_FONT = ("Consolas", 26, "bold")

PROBLEM_RE = re.compile(r"^(?P<name>.+)_문제\.(?P<ext>xlsx|xlsm)$", re.IGNORECASE)

def default_scan_root():
    parent = os.path.dirname(BASE_DIR)
    return parent if parent else BASE_DIR

def find_grade_py(user_path=None):
    candidates = [
        os.path.join(BASE_DIR, "grade.py"),
        os.path.join(os.path.dirname(BASE_DIR), "채점", "grade.py"),
    ]
    if user_path:
        candidates.append(user_path)
    for c in candidates:
        if c and os.path.isfile(c):
            return os.path.abspath(c)
    return None

SET_CONFIG_PATH = os.path.join(BASE_DIR, "세트설정.json")
EXPECTED_DIR_NAME = "기대값"

SCAN_ROOT_SETTING = "스캔폴더"
HIDDEN_SETS_SETTING = "숨긴세트"

ROUTINE_API_VERSION = "3.0.0"
ROUTINE_HTML_NAME = "루틴.html"
ROUTINE_DATA_KEY = "시험장/루틴.html"
ROUTINE_PORT_DEFAULT = 8765
ROUTINE_PORT_SETTING = "루틴포트"
ROUTINE_AUTO_OPEN_SETTING = "루틴자동열기"
WEB_CHECKS_KEY = "_웹체크"

def expected_values_dir(base_dir=None):
    base_dir = base_dir or BASE_DIR
    root = os.path.dirname(base_dir)
    if os.path.basename(base_dir) == "시험장" \
            or os.path.isdir(os.path.join(root, "채점")):
        return os.path.join(root, EXPECTED_DIR_NAME)
    return os.path.join(base_dir, EXPECTED_DIR_NAME)

def key_link_label(s):
    key = s.get("key")
    if not key:
        return "없음"
    own = (os.path.dirname(os.path.abspath(key)) ==
           os.path.abspath(s.get("dir") or "")
           and norm_set_key(key) == (s.get("norm") or ""))
    return os.path.basename(key) + ("" if own else " (자동 연결)")

_ROLE_TOKENS_STRIP = ("문제지", "기대값", "정답지", "답안지", "정답", "답안",
                      "문제")

def norm_set_key(name):
    s = os.path.splitext(os.path.basename(str(name)))[0].lower()
    s = re.sub(r"[\s_\-()\[\]{}.,·~]+", "", s)
    for tok in _ROLE_TOKENS_STRIP:
        s = s.replace(tok, "")
    return s

def file_role(name):
    base = os.path.splitext(os.path.basename(str(name)))[0]
    if "정답" in base or "답안" in base:
        return "answer"
    if "문제" in base:
        return "problem"
    return None

def display_name(problem_path):
    stem = os.path.splitext(os.path.basename(str(problem_path)))[0]
    stem = re.sub(r"[\s_\-]*[(\[]?\s*문제(지)?\s*[)\]]?[\s_\-]*$", "",
                  stem).strip(" _-")
    return stem or os.path.splitext(os.path.basename(str(problem_path)))[0]

def set_tokens(text):
    s = str(text).lower()
    toks = set()
    for m in re.finditer(r"20\d{2}", s):
        toks.add(m.group())
    for m in re.finditer(r"(\d{1,2})\s*회", s):
        toks.add(f"{int(m.group(1))}회")
    for m in re.finditer(r"([ab가나])\s*형", s):
        toks.add(f"{m.group(1)}형")
    for m in re.finditer(r"([1-9])\s*급", s):
        toks.add(f"{m.group(1)}급")
    for m in re.finditer(r"(?<!\d)(2[0-9])(?!\d)", s):
        toks.add(m.group(1))
    for word in ("상시", "코코", "모의", "복원", "기출", "실기", "필기",
                 "드릴", "계산", "컴활"):
        if word in s:
            toks.add(word)
    return toks

_YEAR4_RE = re.compile(r"20\d{2}")
_YEAR2_RE = re.compile(r"2\d")
_ROUND_RE = re.compile(r"\d{1,2}회")
_FORM_RE = re.compile(r"[ab가나]형")

def _year_tokens(toks):
    out = set()
    for t in toks:
        if _YEAR4_RE.fullmatch(t):
            out.add(t)
        elif _YEAR2_RE.fullmatch(t):
            out.add("20" + t)
    return out

def _is_year_token(t):
    return bool(_YEAR4_RE.fullmatch(t) or _YEAR2_RE.fullmatch(t))

def _token_conflict(a, b):
    ya, yb = _year_tokens(a), _year_tokens(b)
    if ya and yb and not (ya & yb):
        return True
    for pat in (_ROUND_RE, _FORM_RE):
        ca = {t for t in a if pat.fullmatch(t)}
        cb = {t for t in b if pat.fullmatch(t)}
        if ca and cb and not (ca & cb):
            return True
    return False

def _shared_count(a, b):
    non_year = {t for t in a & b if not _is_year_token(t)}
    return len(non_year) + len(_year_tokens(a) & _year_tokens(b))

def _numeric_shared(a, b):
    n = len(_year_tokens(a) & _year_tokens(b))
    n += len({t for t in a
              if _ROUND_RE.fullmatch(t) or _FORM_RE.fullmatch(t)} & b)
    return n

def _pdf_score(toks, pt, name=""):
    if _token_conflict(toks, pt):
        return 0
    shared = _shared_count(toks, pt)
    if shared and not {t for t in toks & pt if not _is_year_token(t)} \
            and not any(w in str(name) for w in ("문제", "기대값")):
        return 0
    return shared

def match_pdf_for_set(toks, pdf_paths, others=()):
    scored = []
    for p in pdf_paths:
        name = os.path.basename(p)
        pt = set_tokens(name)
        shared = _pdf_score(toks, pt, name)
        if shared < 1:
            continue
        if any(_pdf_score(ot, pt, name) >= shared for ot in others):
            continue
        scored.append((shared, p))
    if not scored:
        return None
    best = max(s for s, _p in scored)
    matched = [p for s, p in scored if s == best]
    return matched[0] if len(matched) == 1 else None

def _file_link_score(st, path):
    name = os.path.basename(path)
    pt = set_tokens(name)
    shared = _pdf_score(st, pt, name)
    if shared < 1:
        return None
    contained = 1 if all(_has_token(st, t) for t in pt) else 0
    if shared < 2 and not contained:
        return None
    return (shared, _numeric_shared(st, pt), contained)

def link_files_globally(sets, paths, field, origin=None, tokens_of=None,
                        cands_of=None):
    tokens_of = tokens_of or _set_tokens_of
    used = {os.path.abspath(s[field]) for s in sets if s.get(field)}
    cands = [p for p in paths if os.path.abspath(p) not in used]
    if not cands:
        return
    toks = {id(s): tokens_of(s) for s in sets}
    best_any = {}
    for s in sets:
        for p in cands:
            sc = _file_link_score(toks[id(s)], p)
            if sc and sc[0] > best_any.get(p, 0):
                best_any[p] = sc[0]
    pairs = []
    for s in sets:
        if s.get(field):
            continue
        own = cands if cands_of is None else [
            p for p in cands_of(s) if os.path.abspath(p) not in used]
        for p in own:
            sc = _file_link_score(toks[id(s)], p)
            if sc and sc[0] >= best_any.get(p, 0):
                pairs.append((sc, s, p))
    pairs.sort(key=lambda t: (tuple(-x for x in t[0]), t[1]["name"],
                              os.path.basename(t[2])))
    taken_sets, taken_paths = set(), set()
    for i, (sc, s, p) in enumerate(pairs):
        if id(s) in taken_sets or p in taken_paths:
            continue
        rivals = [(s2, p2) for sc2, s2, p2 in pairs[i + 1:]
                  if sc2 == sc and (p2 == p or s2 is s)
                  and id(s2) not in taken_sets and p2 not in taken_paths]
        if rivals:
            names = sorted({s2["name"] for s2, _p in rivals if s2 is not s})
            files = sorted({os.path.basename(p2) for _s, p2 in rivals
                            if p2 != p})
            why = os.path.basename(p) + (
                f" ↔ 세트 {', '.join(names)}과(와) 동점" if names else
                f" ↔ {', '.join(files)} 동점")
            s[field + "_tie"] = why
            for s2, _p in rivals:
                if s2 is not s and not s2.get(field + "_tie"):
                    s2[field + "_tie"] = (os.path.basename(p) +
                                          f" ↔ 세트 {s['name']}과(와) 동점")
            taken_paths.add(p)
            taken_paths.update(p2 for _s, p2 in rivals)
            continue
        s[field] = p
        s.pop(field + "_tie", None)
        if origin:
            s[field + "_origin"] = origin
        taken_sets.add(id(s))
        taken_paths.add(p)

_CFG_DICT_KEYS = ("_슬롯매핑", "_진행", "_자동선택", "_설정", WEB_CHECKS_KEY)

def normalize_set_config(data):
    if not isinstance(data, dict):
        return {}
    out = {}
    for k, v in data.items():
        k = str(k)
        if k in _CFG_DICT_KEYS or not k.startswith("_"):
            if isinstance(v, dict):
                out[k] = v
        else:
            out[k] = v
    return out

def _cfg_section(cfg, key):
    v = (cfg or {}).get(key)
    return v if isinstance(v, dict) else {}

_FILE_LOCK = threading.RLock()

def load_set_config(path=SET_CONFIG_PATH):
    try:
        with _FILE_LOCK, open(path, encoding="utf-8-sig") as f:
            data = json.load(f)
    except Exception:
        return {}
    return normalize_set_config(data)

def save_set_config(config, path=SET_CONFIG_PATH):
    try:
        with _FILE_LOCK, open(path, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        routine_touch()
        return True
    except OSError:
        return False

_ROUTINE_GEN = {"value": None, "lock": threading.Lock()}

def routine_touch():
    with _ROUTINE_GEN["lock"]:
        now = datetime.now()
        prev = _ROUTINE_GEN["value"]
        if prev is not None:
            try:
                prev_dt = datetime.fromisoformat(prev)
                if now <= prev_dt:
                    now = prev_dt + timedelta(microseconds=1)
            except ValueError:
                pass
        _ROUTINE_GEN["value"] = now.isoformat(timespec="microseconds")
        return _ROUTINE_GEN["value"]

def routine_generated():
    with _ROUTINE_GEN["lock"]:
        v = _ROUTINE_GEN["value"]
    return v if v is not None else routine_touch()

def remember_set(s, path=SET_CONFIG_PATH):
    cfg = load_set_config(path)
    key = s.get("norm") or norm_set_key(s["problem"])
    ent = cfg.setdefault(key, {})
    ent.update({
        "name": s.get("name"),
        "problem": os.path.abspath(s["problem"]),
        "answer": os.path.abspath(s["answer"]),
        "key": os.path.abspath(s["key"]) if s.get("key") else None,
        "pdf": os.path.abspath(s["pdf"]) if s.get("pdf") else None,
    })
    if s.get("pdf") and s.get("pdf_확인됨"):
        ent["pdf_확인됨"] = True
    else:
        ent.pop("pdf_확인됨", None)
    return save_set_config(cfg, path)

PDF_MISMATCH_WARNING = ("⚠ 문제지 회차가 세트와 다릅니다 — [문제지 연결]로 "
                        "바로잡으세요")

def pdf_conflicts_with_set(s, pdf=None):
    pdf = pdf or s.get("pdf")
    if not pdf:
        return False
    try:
        return _token_conflict(_set_tokens_of(s),
                               set_tokens(os.path.basename(str(pdf))))
    except Exception:
        return False

def _attach_saved_pdf(s, pdf, confirmed, origin):
    if not (isinstance(pdf, str) and pdf and os.path.isfile(pdf)):
        return False
    if not confirmed and pdf_conflicts_with_set(s, pdf):
        s["pdf_warning"] = (f"저장된 문제지 연결({os.path.basename(pdf)})의 "
                            "회차·형·연도가 세트와 달라 무시했습니다")
        startup_log(f"문제지 연결 무시(회차 불일치, {origin}): "
                    f"세트 '{s.get('name')}' ← {pdf}")
        return False
    s["pdf"] = pdf
    if confirmed:
        s["pdf_확인됨"] = True
    return True

def apply_set_config(sets, config):
    by_key = {}
    for s in sets:
        for d in s.get("중복") or []:
            if d.get("norm"):
                by_key.setdefault(d["norm"], s)
        by_key[s["norm"]] = s
    for k, ent in (config or {}).items():
        if not isinstance(ent, dict) or str(k).startswith("_"):
            continue
        pdf, keyj = ent.get("pdf"), ent.get("key")
        if not isinstance(pdf, str):
            pdf = None
        if not isinstance(keyj, str):
            keyj = None
        confirmed = bool(ent.get("pdf_확인됨"))
        if k in by_key:
            s = by_key[k]
            if pdf:
                _attach_saved_pdf(s, pdf, confirmed, "세트설정")
            if keyj and os.path.isfile(keyj) and keyj != s.get("key"):
                s["key"] = keyj
                s["key_origin"] = "세트설정"
            continue
        prob, ans = ent.get("problem"), ent.get("answer")
        if isinstance(prob, str) and isinstance(ans, str) and prob and ans \
                and os.path.isfile(prob) and os.path.isfile(ans):
            s = {
                "name": str(ent.get("name") or display_name(prob)),
                "norm": k, "dir": os.path.dirname(prob),
                "problem": prob, "answer": ans,
                "key": keyj if keyj and os.path.isfile(keyj) else None,
                "pdf": None,
                "saved": True,
            }
            if pdf:
                _attach_saved_pdf(s, pdf, confirmed, "세트설정")
            sets.append(s)
    return sets

def _is_under(path, root):
    try:
        return os.path.commonpath([os.path.abspath(path),
                                   os.path.abspath(root)]) == \
            os.path.abspath(root)
    except ValueError:
        return False

def _dedupe_keys_by_name(paths, set_dir, key_dirs):
    by_name = {}
    for p in paths:
        by_name.setdefault(os.path.basename(p), []).append(p)
    out = []
    for name, group in by_name.items():
        if len(group) == 1:
            out.append(group[0])
        else:
            out.append(_pick_exact_key(group, set_dir, key_dirs)[0])
    return out

def _pick_exact_key(cands, set_dir, key_dirs):
    if not cands:
        return None, None
    own = [p for p in cands if os.path.dirname(p) == set_dir]
    if own:
        return sorted(own)[0], "세트 폴더"
    kd = [os.path.abspath(d) for d in key_dirs]
    in_kd = [p for p in cands if os.path.abspath(os.path.dirname(p)) in kd]
    if in_kd:
        return sorted(in_kd)[0], "기대값 폴더"
    return sorted(cands)[0], "다른 폴더"

def _dir_depth(path, root=None):
    try:
        rel = os.path.relpath(os.path.abspath(path), os.path.abspath(root)) \
            if root else os.path.abspath(path)
    except ValueError:
        rel = os.path.abspath(path)
    return len([p for p in rel.replace("\\", "/").split("/") if p not in ("", ".")])

def _dup_entry(s):
    return {k: s.get(k) for k in ("name", "norm", "dir", "problem", "answer",
                                  "pdf", "key")}

def merge_duplicate_sets(sets, root=None):
    by_size = {}
    for s in sets:
        try:
            size = os.path.getsize(s["problem"])
        except (OSError, TypeError):
            continue
        by_size.setdefault(size, []).append(s)
    drop = set()
    for group in by_size.values():
        if len(group) < 2:
            continue
        by_hash = {}
        for s in group:
            h = file_sha256(s["problem"])
            if h:
                by_hash.setdefault(h, []).append(s)
        for dups in by_hash.values():
            if len(dups) < 2:
                continue
            dups.sort(key=lambda s: (0 if (s.get("pdf") or s.get("key")) else 1,
                                     _dir_depth(s["dir"], root), s["name"]))
            rep = dups[0]
            rep.setdefault("중복", [])
            for d in dups[1:]:
                rep["중복"].append(_dup_entry(d))
                rep["중복"].extend(d.get("중복") or [])
                if not rep.get("pdf") and d.get("pdf"):
                    rep["pdf"] = d["pdf"]
                if not rep.get("key") and d.get("key"):
                    rep["key"], rep["key_origin"] = d["key"], d.get("key_origin")
                drop.add(id(d))
    return [s for s in sets if id(s) not in drop]

def similar_named_sets(sets):
    out = []
    for i, a in enumerate(sets):
        ta = _set_tokens_of(a)
        for b in sets[i + 1:]:
            tb = _set_tokens_of(b)
            if _token_conflict(ta, tb):
                continue
            common = ta & tb
            if _shared_count(ta, tb) >= 2 and any(
                    _ROUND_RE.fullmatch(t) or _FORM_RE.fullmatch(t)
                    for t in common):
                out.append((a, b, sorted(common)))
    return out

def scan_sets(root, config=None, key_dirs=None):
    sets = []
    all_pdfs = []
    all_keys = []
    groups = {}
    if key_dirs is None:
        key_dirs = [expected_values_dir()]

    def add_file(dirpath, fn):
        ext = os.path.splitext(fn)[1].lower()
        if ext not in (".xlsx", ".xlsm", ".pdf", ".json"):
            return
        if fn.startswith(("풀이_", "채점결과", "~$", ".")):
            return
        path = os.path.join(dirpath, fn)
        if ext == ".json" and "기대값" not in fn:
            return
        if ext == ".pdf":
            all_pdfs.append(path)
        elif ext == ".json":
            if path in all_keys:
                return
            all_keys.append(path)
        k = norm_set_key(fn)
        if not k:
            return
        g = groups.setdefault(k, {"excel": [], "pdf": [], "json": []})
        if ext in (".xlsx", ".xlsm"):
            g["excel"].append(path)
        elif ext == ".pdf":
            g["pdf"].append(path)
        else:
            g["json"].append(path)

    if root and os.path.isdir(root):
        for dirpath, dirnames, filenames in os.walk(root):
            dirnames[:] = [d for d in dirnames
                           if not d.startswith((".", "__"))
                           and d != "채점결과"]
            for fn in sorted(filenames):
                add_file(dirpath, fn)
    for kd in key_dirs:
        if not os.path.isdir(kd) or (root and _is_under(kd, root)):
            continue
        for fn in sorted(os.listdir(kd)):
            if fn.lower().endswith(".json"):
                add_file(kd, fn)
    for k, g in groups.items():
        answers = [p for p in g["excel"] if file_role(p) == "answer"]
        problems = [p for p in g["excel"] if file_role(p) == "problem"]
        answer = sorted(answers)[0] if answers else None
        problem = sorted(problems)[0] if problems else None
        if not problem and answer and len(g["excel"]) == 2:
            problem = next(p for p in g["excel"] if p != answer)
        if not (problem and answer) or problem == answer:
            continue
        set_dir = os.path.dirname(problem)
        key, origin = _pick_exact_key(g["json"], set_dir, key_dirs)
        pdfs = sorted(g["pdf"], key=lambda p: (os.path.dirname(p) != set_dir, p))
        sets.append({
            "name": display_name(problem), "norm": k,
            "dir": set_dir,
            "problem": problem, "answer": answer,
            "key": key, "key_origin": origin,
            "pdf": pdfs[0] if pdfs else None,
            "pdf_사본": pdfs[1:],
        })
    sets = merge_duplicate_sets(sets, root)
    used_pdf_names = {os.path.basename(s["pdf"]) for s in sets if s.get("pdf")}
    link_files_globally(
        sets, [p for p in all_pdfs if os.path.basename(p) not in used_pdf_names],
        "pdf")
    used_k = {os.path.abspath(s["key"]) for s in sets if s.get("key")}
    used_k_names = {os.path.basename(s["key"]) for s in sets if s.get("key")}
    free_keys = [p for p in all_keys if os.path.abspath(p) not in used_k
                 and os.path.basename(p) not in used_k_names]
    link_files_globally(
        sets, free_keys, "key", origin="토큰 일치",
        cands_of=lambda s: _dedupe_keys_by_name(free_keys, s["dir"], key_dirs))
    apply_set_config(sets, config if config is not None else load_set_config())
    sets.sort(key=lambda s: s["name"])
    return sets

def build_direct_set(problem, answer, pdf=None, key_dirs=None):
    d = os.path.dirname(os.path.abspath(problem))
    k = norm_set_key(problem)
    key = None
    origin = None
    auto_pdf = None
    key_cands = []
    if key_dirs is None:
        key_dirs = [expected_values_dir()]
    try:
        for fn in os.listdir(d):
            if fn.startswith("풀이_"):
                continue
            ext = os.path.splitext(fn)[1].lower()
            if ext == ".json" and "기대값" in fn and norm_set_key(fn) == k:
                key, origin = os.path.join(d, fn), "세트 폴더"
            elif ext == ".json" and "기대값" in fn:
                key_cands.append(os.path.join(d, fn))
            elif ext == ".pdf" and norm_set_key(fn) == k:
                auto_pdf = os.path.join(d, fn)
    except OSError:
        pass
    for kd in key_dirs:
        if key or not os.path.isdir(kd) or os.path.abspath(kd) == d:
            continue
        try:
            names = sorted(os.listdir(kd))
        except OSError:
            continue
        have = {os.path.basename(p) for p in key_cands}
        for fn in names:
            if not (fn.lower().endswith(".json") and "기대값" in fn):
                continue
            if norm_set_key(fn) == k and key is None:
                key, origin = os.path.join(kd, fn), "기대값 폴더"
            elif fn not in have:
                key_cands.append(os.path.join(kd, fn))
    if key is None and key_cands:
        toks = set_tokens(display_name(problem)) | set_tokens(os.path.basename(d))
        cand = match_pdf_for_set(toks, key_cands)
        if cand and _pdf_score(toks, set_tokens(os.path.basename(cand))) >= 2:
            key, origin = cand, "토큰 일치"
    return {
        "name": display_name(problem), "norm": k, "dir": d,
        "problem": os.path.abspath(problem),
        "answer": os.path.abspath(answer),
        "key": key, "key_origin": origin,
        "pdf": os.path.abspath(pdf) if pdf else auto_pdf,
    }

def _coerce_score(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, str):
        m = re.search(r"-?\d+(?:\.\d+)?", v)
        if m:
            f = float(m.group())
            return int(f) if f.is_integer() else f
    return None

def normalize_record(r):
    if not isinstance(r, dict):
        return None
    out = dict(r)
    name = r.get("세트명") or r.get("set") or r.get("세트")
    out["세트명"] = str(name) if name is not None else "?"
    when = r.get("일시") or r.get("date") or r.get("when")
    out["일시"] = str(when) if when is not None else "?"
    out["점수"] = _coerce_score(r.get("점수", r.get("score")))
    if "mode" in out and out["mode"] is not None \
            and not isinstance(out["mode"], str):
        out["mode"] = str(out["mode"])
    if out.get("만점") is not None:
        out["만점"] = _coerce_score(out["만점"])
    for k in ("소요시간", "리포트", "영역", "day", "루틴"):
        if out.get(k) is not None and not isinstance(out[k], str):
            out[k] = str(out[k])
    return out

def normalize_records(data):
    if isinstance(data, dict):
        inner = next((data[k] for k in ("기록", "records", "items")
                      if isinstance(data.get(k), list)), None)
        if inner is None:
            inner = []
            for k, v in data.items():
                if isinstance(v, dict):
                    v = dict(v)
                    v.setdefault("세트명", str(k))
                    inner.append(v)
                elif isinstance(v, list):
                    inner.extend(x for x in v if isinstance(x, dict))
        data = inner
    if not isinstance(data, list):
        return []
    out = []
    for r in data:
        n = normalize_record(r)
        if n is not None:
            out.append(n)
    return out

def load_records(path=RECORDS_PATH):
    if path == RECORDS_PATH:
        _ensure_records_home()
    try:
        with _FILE_LOCK, open(path, encoding="utf-8-sig") as f:
            data = json.load(f)
    except Exception:
        return []
    try:
        return normalize_records(data)
    except Exception as e:
        log_error("기록.json 정규화", e)
        return []

def _backup_corrupt_json(path):
    try:
        if os.path.isfile(path) and os.path.getsize(path) > 0:
            with open(path, encoding="utf-8-sig") as f:
                json.load(f)
    except Exception:
        try:
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            shutil.copy2(path, f"{path}.corrupt-{stamp}")
            startup_log(f"깨진 JSON 백업: {path}.corrupt-{stamp}")
        except OSError:
            pass

def append_record(record, path=RECORDS_PATH):
    if path == RECORDS_PATH:
        _ensure_records_home()
    with _FILE_LOCK:
        _backup_corrupt_json(path)
        records = load_records(path)
        records.append(record)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(records, f, ensure_ascii=False, indent=2)
    routine_touch()
    return records

def records_summary(records):
    by = {}
    for r in records:
        if r.get("mode") in ("부분연습", "오답재풀이"):
            continue
        by.setdefault(r.get("세트명", "?"), []).append(r)
    out = {}
    for name, rows in by.items():
        scores = [r.get("점수") for r in rows
                  if isinstance(r.get("점수"), (int, float))]
        out[name] = {
            "best": max(scores) if scores else None,
            "recent": [r.get("점수") for r in rows[-3:]],
        }
    return out

XLSX_MAIN_CT = ("application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet.main+xml")
XLSM_MAIN_CT = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"

CONTENT_TYPES_NAME = "[Content_Types].xml"
ZIP_FLAG_DATA_DESCRIPTOR = 0x08

def convert_xlsx_to_xlsm(src, dst):
    import zipfile
    with zipfile.ZipFile(src) as zin:
        items = [i for i in zin.infolist() if not i.is_dir()]
        names = [i.filename for i in items]
        if CONTENT_TYPES_NAME not in names:
            raise ValueError("[Content_Types].xml이 없어 Excel 파일이 아닙니다")
        order = ([i for i in items if i.filename == CONTENT_TYPES_NAME]
                 + [i for i in items if i.filename != CONTENT_TYPES_NAME])
        seen = set()
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in order:
                if item.filename in seen:
                    continue
                seen.add(item.filename)
                data = zin.read(item.filename)
                if item.filename == CONTENT_TYPES_NAME:
                    data = data.replace(XLSX_MAIN_CT.encode("utf-8"),
                                        XLSM_MAIN_CT.encode("utf-8"))
                    if XLSM_MAIN_CT.encode("utf-8") not in data:
                        raise ValueError("워크북 content-type을 찾지 못해 "
                                         "xlsm으로 바꿀 수 없습니다")
                dt = item.date_time
                if not dt or dt[0] < 1980:
                    dt = (1980, 1, 1, 0, 0, 0)
                info = zipfile.ZipInfo(item.filename, date_time=dt)
                info.compress_type = zipfile.ZIP_DEFLATED
                zout.writestr(info, data)
    return dst

def verify_xlsm_copy(path, load_openpyxl=True):
    import zipfile
    try:
        with zipfile.ZipFile(path) as zf:
            bad = zf.testzip()
            if bad:
                return False, f"손상된 항목: {bad}"
            infos = zf.infolist()
            if not infos or infos[0].filename != CONTENT_TYPES_NAME:
                return False, "[Content_Types].xml이 첫 항목이 아님"
            dd = [i.filename for i in infos
                  if i.flag_bits & ZIP_FLAG_DATA_DESCRIPTOR]
            if dd:
                return False, f"data descriptor 비트 항목: {dd[:3]}"
            ct = zf.read(CONTENT_TYPES_NAME)
            if XLSM_MAIN_CT.encode("utf-8") not in ct:
                return False, "매크로 content-type 없음"
    except Exception as e:
        return False, f"ZIP 검사 실패: {e}"
    note = f"ZIP 검사 통과({len(infos)}항목)"
    if load_openpyxl:
        try:
            import openpyxl
            openpyxl.load_workbook(path).close()
            note += ", openpyxl 로드 OK"
        except ImportError:
            note += ", openpyxl 없음(로드 생략)"
        except Exception as e:
            return False, f"openpyxl 로드 실패: {e}"
    return True, note

def _strip_motw(path):
    try:
        os.remove(path + ":Zone.Identifier")
    except OSError:
        pass

def _unique_stem(d, stem):
    cand = os.path.join(d, stem)
    n = 2
    while os.path.exists(cand + ".xlsm") or os.path.exists(cand + ".xlsx"):
        cand = os.path.join(d, f"{stem}_{n}")
        n += 1
    return cand

def copy_as_macro_enabled(source, dst_stem):
    ext = os.path.splitext(source)[1].lower()
    note = ""
    if ext == ".xlsx":
        dst = dst_stem + ".xlsm"
        try:
            convert_xlsx_to_xlsm(source, dst)
            ok, why = verify_xlsm_copy(dst)
            if not ok:
                raise RuntimeError(why)
            note = f"xlsx→xlsm 변환 성공 ({why})"
        except Exception as e:
            try:
                if os.path.isfile(dst):
                    os.remove(dst)
            except OSError:
                pass
            dst = dst_stem + ".xlsx"
            shutil.copy2(source, dst)
            note = f"xlsx→xlsm 변환 실패({e}) → .xlsx 그대로 복사"
    else:
        dst = dst_stem + (ext or ".xlsm")
        shutil.copy2(source, dst)
        note = f"{ext or '.xlsm'} 원본 그대로 복사"
    _strip_motw(dst)
    try:
        size = os.path.getsize(dst)
    except OSError:
        size = -1
    startup_log(f"사본 생성: {dst} ({size:,} bytes) · {note} · 원본={source}")
    return dst

def make_attempt_copy(problem, set_name, when=None):
    when = when or datetime.now()
    stamp = when.strftime("%Y%m%d_%H%M")
    d = os.path.dirname(os.path.abspath(problem))
    return copy_as_macro_enabled(
        problem, _unique_stem(d, f"풀이_{set_name}_{stamp}"))

def get_app_setting(name, default=None, path=None):
    cfg = load_set_config(path or SET_CONFIG_PATH)
    return _cfg_section(cfg, "_설정").get(name, default)

def set_app_setting(name, value, path=None):
    p = path or SET_CONFIG_PATH
    cfg = load_set_config(p)
    cfg.setdefault("_설정", {})[name] = value
    return save_set_config(cfg, p)

def configured_scan_root(path=None, cfg=None):
    if cfg is not None:
        v = _cfg_section(cfg, "_설정").get(SCAN_ROOT_SETTING)
    else:
        v = get_app_setting(SCAN_ROOT_SETTING, None, path=path)
    if isinstance(v, str) and v.strip() and os.path.isdir(v):
        return os.path.abspath(v)
    return default_scan_root()

def save_scan_root(folder, path=None):
    if folder is None or not str(folder).strip():
        set_app_setting(SCAN_ROOT_SETTING, None, path=path)
        return default_scan_root()
    set_app_setting(SCAN_ROOT_SETTING, os.path.abspath(str(folder)), path=path)
    return configured_scan_root(path=path)

def scan_root_is_default(path=None, cfg=None):
    return os.path.abspath(configured_scan_root(path, cfg)) == \
        os.path.abspath(default_scan_root())

def hidden_set_keys(cfg=None, path=None):
    sec = _cfg_section(cfg, "_설정") if cfg is not None \
        else _cfg_section(load_set_config(path or SET_CONFIG_PATH), "_설정")
    raw = sec.get(HIDDEN_SETS_SETTING)
    out = []
    for x in raw if isinstance(raw, (list, tuple)) else []:
        k = str(x or "").strip()
        if not k:
            continue
        k = norm_set_key(k) or k
        if k not in out:
            out.append(k)
    return sorted(out)

def set_hidden_keys(s):
    keys = []
    cands = [(s or {}).get("norm")] + [d.get("norm")
                                       for d in ((s or {}).get("중복") or [])]
    for k in cands:
        k = str(k or "").strip()
        k = (norm_set_key(k) or k) if k else ""
        if k and k not in keys:
            keys.append(k)
    if not keys and (s or {}).get("problem"):
        keys.append(norm_set_key(s["problem"]))
    return keys

def is_hidden_set(s, hidden=None, cfg=None, path=None):
    if hidden is None:
        hidden = hidden_set_keys(cfg=cfg, path=path)
    return bool(set(hidden or []) & set(set_hidden_keys(s)))

def split_hidden_sets(sets, hidden=None, cfg=None, path=None):
    if hidden is None:
        hidden = hidden_set_keys(cfg=cfg, path=path)
    hide = set(hidden or [])
    shown, gone = [], []
    for s in sets or []:
        if hide & set(set_hidden_keys(s)):
            s["숨김"] = True
            gone.append(s)
        else:
            s.pop("숨김", None)
            shown.append(s)
    return shown, gone

def set_set_hidden(key, hidden=True, path=None):
    k = str(key or "").strip()
    k = (norm_set_key(k) or k) if k else ""
    cur = hidden_set_keys(path=path)
    if not k:
        return cur
    if hidden and k not in cur:
        cur = sorted(cur + [k])
    elif not hidden and k in cur:
        cur = [x for x in cur if x != k]
    else:
        return cur
    set_app_setting(HIDDEN_SETS_SETTING, cur, path=path)
    return cur

def clear_hidden_sets(path=None):
    set_app_setting(HIDDEN_SETS_SETTING, [], path=path)
    return []

TIMER_OVERLAY_SETTING = "타이머오버레이"
TIMER_OVERLAY_POS_SETTING = "타이머오버레이위치"
MONITOR_PDF_SETTING = "문제지모니터"
MONITOR_EXCEL_SETTING = "엑셀모니터"
MONITOR_TIMER_SETTING = "타이머모니터"

OVERLAY_MARGIN = 16
OVERLAY_SIZE = (190, 96)
OVERLAY_MIN_VISIBLE = 40
EXCEL_PLACE_DELAY_MS = 2500
PDF_PLACE_DELAY_MS = 700
PLACE_RETRY_MS = 1500
EXCEL_PLACE_TRIES = 6
PDF_PLACE_TRIES = 8

_H_WORDS = {1: ("",), 2: ("왼쪽", "오른쪽"), 3: ("왼쪽", "가운데", "오른쪽")}
_V_WORDS = {1: ("",), 2: ("위", "아래"), 3: ("위", "가운데", "아래")}

def _rect_tuple(value, default=None):
    try:
        left, top, right, bottom = (int(v) for v in value)
    except Exception:
        return default
    if right <= left or bottom <= top:
        return default
    return (left, top, right, bottom)

def rect_size(rect):
    left, top, right, bottom = rect
    return (right - left, bottom - top)

def monitor_position_words(mons):
    rects = [m.get("rect") for m in mons or []]
    if not rects or any(r is None for r in rects):
        return ["" for _ in rects]
    lefts = sorted({r[0] for r in rects})
    tops = sorted({r[1] for r in rects})
    vertical = len(lefts) == 1 and len(tops) > 1
    keys = tops if vertical else lefts
    words = _V_WORDS if vertical else _H_WORDS
    tail = "위에서 %d번째" if vertical else "왼쪽에서 %d번째"
    out = []
    for r in rects:
        i = keys.index(r[1] if vertical else r[0])
        out.append(words[len(keys)][i] if len(keys) in words
                   else tail % (i + 1))
    return out

def monitor_label(m):
    bits = []
    if m.get("primary"):
        bits.append("주 모니터")
    bits.append(f"{m.get('width')}×{m.get('height')}")
    if m.get("position"):
        bits.append(m["position"])
    return f"모니터 {int(m.get('index') or 0) + 1} (" + ", ".join(bits) + ")"

def normalize_monitors(raw):
    mons = []
    for item in raw or []:
        if not isinstance(item, dict):
            continue
        rect = _rect_tuple(item.get("rect"))
        if rect is None:
            continue
        mons.append({"rect": rect,
                     "work": _rect_tuple(item.get("work"), rect),
                     "primary": bool(item.get("primary"))})
    mons.sort(key=lambda m: (m["rect"][0], m["rect"][1]))
    if mons and not any(m["primary"] for m in mons):
        mons[0]["primary"] = True
    words = monitor_position_words(mons)
    for i, m in enumerate(mons):
        m["index"] = i
        m["width"], m["height"] = rect_size(m["rect"])
        m["position"] = words[i]
        m["label"] = monitor_label(m)
    return mons

def primary_monitor(mons):
    for m in mons or []:
        if m.get("primary"):
            return m
    return (list(mons) or [None])[0]

def pick_monitor(mons, index):
    mons = list(mons or [])
    if not mons:
        return None
    try:
        i = int(index)
    except (TypeError, ValueError):
        return primary_monitor(mons)
    return mons[i] if 0 <= i < len(mons) else primary_monitor(mons)

def monitor_ui_visible(mons):
    return len(mons or []) >= 2

def default_monitor_choices(mons):
    mons = list(mons or [])
    if not mons:
        return {"pdf": 0, "excel": 0, "timer": 0}
    excel = int((primary_monitor(mons) or {}).get("index") or 0)
    other = next((int(m["index"]) for m in mons
                  if int(m["index"]) != excel), excel)
    return {"pdf": other, "excel": excel, "timer": excel}

def overlay_top_right(rect, size=OVERLAY_SIZE, margin=OVERLAY_MARGIN):
    left, top, right, _bottom = rect
    width, _height = size
    return (max(int(left), int(right) - int(width) - int(margin)),
            max(int(top), int(top) + int(margin)))

def point_on_any_monitor(x, y, mons, size=OVERLAY_SIZE,
                         need=OVERLAY_MIN_VISIBLE):
    width, height = size
    for m in mons or []:
        rect = m.get("rect") or m.get("work")
        if not rect:
            continue
        left, top, right, bottom = rect
        if (min(x + width, right) - max(x, left) >= need
                and min(y + height, bottom) - max(y, top) >= need):
            return True
    return False

def overlay_position(mons, index=None, size=OVERLAY_SIZE, saved=None,
                     margin=OVERLAY_MARGIN):
    mons = list(mons or [])
    if saved is not None:
        try:
            sx, sy = int(saved[0]), int(saved[1])
        except Exception:
            sx = sy = None
        if sx is not None and (not mons or point_on_any_monitor(
                sx, sy, mons, size=size)):
            return (sx, sy)
    m = pick_monitor(mons, index)
    if m is None:
        return None
    return overlay_top_right(m.get("work") or m.get("rect"), size, margin)

def overlay_position_for_screen(screen_w, screen_h, size=OVERLAY_SIZE,
                                saved=None, margin=OVERLAY_MARGIN):
    rect = (0, 0, int(screen_w), int(screen_h))
    return overlay_position([{"rect": rect, "work": rect, "primary": True,
                              "index": 0}], 0, size, saved, margin)

def clamp_overlay_position(x, y, mons, size=OVERLAY_SIZE,
                           need=OVERLAY_MIN_VISIBLE):
    x, y = int(x), int(y)
    mons = list(mons or [])
    if not mons or point_on_any_monitor(x, y, mons, size=size, need=need):
        return (x, y)
    width, height = size
    cx, cy = x + width / 2.0, y + height / 2.0

    def _dist(m):
        left, top, right, bottom = m["rect"]
        return ((left + right) / 2.0 - cx) ** 2 + \
               ((top + bottom) / 2.0 - cy) ** 2

    near = min(mons, key=_dist)
    left, top, right, bottom = near.get("work") or near["rect"]
    return (max(int(left), min(x, int(right) - int(width))),
            max(int(top), min(y, int(bottom) - int(height))))

def window_rect_for_monitor(m, margin=0):
    if not m:
        return None
    rect = m.get("work") or m.get("rect")
    if not rect:
        return None
    left, top, right, bottom = rect
    return (int(left) + margin, int(top) + margin,
            max(1, int(right) - int(left) - 2 * margin),
            max(1, int(bottom) - int(top) - 2 * margin))

def pick_new_windows(before, after):
    seen = set(before or ())
    return [h for h in (after or ()) if h not in seen]

def timer_overlay_enabled(cfg=None, path=None):
    if cfg is not None:
        v = _cfg_section(cfg, "_설정").get(TIMER_OVERLAY_SETTING, True)
    else:
        v = get_app_setting(TIMER_OVERLAY_SETTING, True, path=path)
    return True if v is None else bool(v)

def set_timer_overlay_enabled(on, path=None):
    on = bool(on)
    set_app_setting(TIMER_OVERLAY_SETTING, on, path=path)
    return on

def saved_overlay_pos(cfg=None, path=None):
    if cfg is not None:
        v = _cfg_section(cfg, "_설정").get(TIMER_OVERLAY_POS_SETTING)
    else:
        v = get_app_setting(TIMER_OVERLAY_POS_SETTING, None, path=path)
    if isinstance(v, dict):
        v = [v.get("x"), v.get("y")]
    try:
        x, y = (int(a) for a in v)
    except Exception:
        return None
    return (x, y)

def save_overlay_pos(x, y, path=None):
    if x is None or y is None:
        set_app_setting(TIMER_OVERLAY_POS_SETTING, None, path=path)
        return None
    try:
        pos = [int(x), int(y)]
    except (TypeError, ValueError):
        return None
    set_app_setting(TIMER_OVERLAY_POS_SETTING, pos, path=path)
    return (pos[0], pos[1])

_MONITOR_SETTINGS = (("pdf", MONITOR_PDF_SETTING),
                     ("excel", MONITOR_EXCEL_SETTING),
                     ("timer", MONITOR_TIMER_SETTING))

def monitor_assignment(cfg=None, path=None):
    sec = _cfg_section(cfg, "_설정") if cfg is not None else \
        _cfg_section(load_set_config(path or SET_CONFIG_PATH), "_설정")
    out = {}
    for key, name in _MONITOR_SETTINGS:
        try:
            out[key] = int(sec.get(name))
        except (TypeError, ValueError):
            out[key] = None
    return out

def save_monitor_assignment(pdf=None, excel=None, timer=None, path=None):
    p = path or SET_CONFIG_PATH
    cfg = load_set_config(p)
    sec = cfg.setdefault("_설정", {})
    for value, (_key, name) in zip((pdf, excel, timer), _MONITOR_SETTINGS):
        if value is None:
            continue
        try:
            sec[name] = int(value)
        except (TypeError, ValueError):
            continue
    save_set_config(cfg, p)
    return monitor_assignment(cfg=cfg)

def clear_monitor_assignment(path=None):
    p = path or SET_CONFIG_PATH
    cfg = load_set_config(p)
    sec = cfg.setdefault("_설정", {})
    for _key, name in _MONITOR_SETTINGS:
        sec.pop(name, None)
    save_set_config(cfg, p)
    return monitor_assignment(cfg=cfg)

def resolve_monitor_assignment(mons, cfg=None, path=None):
    mons = list(mons or [])
    out = dict(default_monitor_choices(mons))
    saved = monitor_assignment(cfg=cfg, path=path)
    for key in out:
        v = saved.get(key)
        if v is not None and (0 <= v < len(mons) if mons else v == 0):
            out[key] = v
    return out

SW_MAXIMIZE = 3
SW_RESTORE = 9
SWP_NOZORDER = 0x0004
SWP_NOACTIVATE = 0x0010
SWP_SHOWWINDOW = 0x0040
GW_OWNER = 4
MONITORINFOF_PRIMARY = 1
EXCEL_WINDOW_CLASS = "XLMAIN"

def monitors_supported():
    if sys.platform != "win32":
        return False
    try:
        import ctypes
        return hasattr(ctypes, "windll")
    except Exception:
        return False

def enum_monitors_raw():
    if not monitors_supported():
        return []
    try:
        import ctypes

        class _RECT(ctypes.Structure):
            _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                        ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

        class _MONITORINFO(ctypes.Structure):
            _fields_ = [("cbSize", ctypes.c_ulong), ("rcMonitor", _RECT),
                        ("rcWork", _RECT), ("dwFlags", ctypes.c_ulong)]

        user32 = ctypes.windll.user32
        found = []
        proc = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_void_p,
                                  ctypes.c_void_p, ctypes.POINTER(_RECT),
                                  ctypes.c_void_p)

        def _on_monitor(hmon, _hdc, _lprc, _data):
            try:
                info = _MONITORINFO()
                info.cbSize = ctypes.sizeof(_MONITORINFO)
                if user32.GetMonitorInfoW(hmon, ctypes.byref(info)):
                    found.append({
                        "rect": (info.rcMonitor.left, info.rcMonitor.top,
                                 info.rcMonitor.right, info.rcMonitor.bottom),
                        "work": (info.rcWork.left, info.rcWork.top,
                                 info.rcWork.right, info.rcWork.bottom),
                        "primary": bool(info.dwFlags & MONITORINFOF_PRIMARY)})
            except Exception:
                pass
            return 1

        user32.EnumDisplayMonitors(None, None, proc(_on_monitor), 0)
        return found
    except Exception as e:
        log_error("모니터 열거", e)
        return []

def list_monitors():
    try:
        return normalize_monitors(enum_monitors_raw())
    except Exception as e:
        log_error("모니터 목록", e)
        return []

def _window_class(user32, hwnd):
    import ctypes
    buf = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buf, 256)
    return buf.value

def _window_title(user32, hwnd):
    import ctypes
    n = int(user32.GetWindowTextLengthW(hwnd) or 0)
    if n <= 0:
        return ""
    buf = ctypes.create_unicode_buffer(n + 1)
    user32.GetWindowTextW(hwnd, buf, n + 1)
    return buf.value

def top_level_windows(visible_only=True):
    if not monitors_supported():
        return []
    try:
        import ctypes
        user32 = ctypes.windll.user32
        out = []
        proc = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_void_p,
                                  ctypes.c_void_p)

        def _on_window(hwnd, _data):
            try:
                if not visible_only or user32.IsWindowVisible(hwnd):
                    out.append(int(hwnd))
            except Exception:
                pass
            return 1

        user32.EnumWindows(proc(_on_window), 0)
        return out
    except Exception as e:
        log_error("창 열거", e)
        return []

def excel_windows(pid=None, strict=False):
    if not monitors_supported():
        return []
    try:
        import ctypes
        user32 = ctypes.windll.user32
        mine, others = [], []
        for hwnd in top_level_windows():
            try:
                if _window_class(user32, hwnd) != EXCEL_WINDOW_CLASS:
                    continue
                wpid = ctypes.c_ulong(0)
                user32.GetWindowThreadProcessId(hwnd, ctypes.byref(wpid))
                if pid is not None and int(wpid.value) == int(pid):
                    mine.append(hwnd)
                else:
                    others.append(hwnd)
            except Exception:
                continue
        return mine if strict else (mine or others)
    except Exception as e:
        log_error("Excel 창 찾기", e)
        return []

def move_window_to_monitor(hwnd, mon, maximize=True):
    rect = window_rect_for_monitor(mon)
    if not hwnd or rect is None or not monitors_supported():
        return False
    x, y, w, h = rect
    try:
        import ctypes
        user32 = ctypes.windll.user32
        user32.ShowWindow(hwnd, SW_RESTORE)
        user32.SetWindowPos(hwnd, None, int(x), int(y), int(w), int(h),
                            SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW)
        if maximize:
            user32.ShowWindow(hwnd, SW_MAXIMIZE)
        return True
    except Exception as e:
        log_error("창 배치", e)
        return False

def place_excel_on_monitor(mon, pid=None, maximize=True):
    if mon is None:
        return 0, "모니터 정보 없음"
    wins, note = (excel_windows(pid, strict=True) if pid is not None else [],
                  "프로세스 대조")
    if not wins:
        wins, note = excel_windows()[:1], "맨 앞 Excel 창"
    if not wins:
        return 0, f"Excel 창({EXCEL_WINDOW_CLASS})을 찾지 못함"
    moved = sum(1 for h in wins if move_window_to_monitor(h, mon, maximize))
    return moved, (f"Excel 창 {moved}개 배치({note})" if moved
                   else "창 배치 실패")

def place_new_window_on_monitor(mon, before, maximize=True,
                                skip_classes=(EXCEL_WINDOW_CLASS,)):
    if mon is None:
        return 0, "모니터 정보 없음"
    if not monitors_supported():
        return 0, "Windows 아님"
    try:
        import ctypes
        user32 = ctypes.windll.user32
        for hwnd in pick_new_windows(before, top_level_windows()):
            try:
                if _window_class(user32, hwnd) in (skip_classes or ()):
                    continue
                if not _window_title(user32, hwnd).strip():
                    continue
                if user32.GetWindow(hwnd, GW_OWNER):
                    continue
            except Exception:
                continue
            if move_window_to_monitor(hwnd, mon, maximize):
                return 1, "새 창 1개 배치"
        return 0, "새로 뜬 창을 찾지 못함"
    except Exception as e:
        log_error("문제지 창 배치", e)
        return 0, f"배치 실패({e})"

TRUST_MANUAL_GUIDE = (
    "수동 설정: Excel → 파일 → 옵션 → 보안 센터 → 보안 센터 설정 → "
    "신뢰할 수 있는 위치 → '새 위치 추가'에서 학습 폴더를 추가하고 "
    "'이 위치의 하위 폴더도 신뢰할 수 있음'을 체크하세요.")

def build_trusted_location_values(folder, version="16.0", slot="LocationCC"):
    path = (rf"Software\Microsoft\Office\{version}\Excel\Security"
            rf"\Trusted Locations\{slot}")
    folder = str(folder).rstrip("\\/") + "\\"
    return {"key": path,
            "values": {"Path": folder, "AllowSubfolders": 1,
                       "Description": "코코 시험장 학습 폴더"}}

def register_trusted_location(folder, version=None):
    try:
        import winreg
    except ImportError:
        return False, ("이 기능은 Windows에서만 동작합니다.\n\n"
                       + TRUST_MANUAL_GUIDE)
    versions = [version] if version else []
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Office") as k:
            i = 0
            while True:
                try:
                    sub = winreg.EnumKey(k, i)
                    i += 1
                except OSError:
                    break
                if re.match(r"^\d+\.\d+$", sub):
                    versions.append(sub)
    except OSError:
        pass
    if "16.0" not in versions:
        versions.append("16.0")
    seen = []
    for v in sorted(set(versions), key=lambda x: -float(x)):
        seen.append(v)
    last_err = None
    for v in seen:
        try:
            spec_reg = build_trusted_location_values(folder, v)
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER,
                                  spec_reg["key"]) as k:
                vals = spec_reg["values"]
                winreg.SetValueEx(k, "Path", 0, winreg.REG_SZ, vals["Path"])
                winreg.SetValueEx(k, "AllowSubfolders", 0, winreg.REG_DWORD,
                                  vals["AllowSubfolders"])
                winreg.SetValueEx(k, "Description", 0, winreg.REG_SZ,
                                  vals["Description"])
            return True, (f"Excel {v} 신뢰 위치로 등록했습니다:\n{folder}\n\n"
                          "이미 열려 있는 Excel은 닫았다가 다시 여세요.")
        except OSError as e:
            last_err = e
    return False, (f"레지스트리 등록에 실패했습니다 ({last_err}).\n\n"
                   + TRUST_MANUAL_GUIDE)

def build_grade_cmd(grade_py, problem, answer, student, key=None, html=None,
                    json_out=None, history=None, sheets=None):
    cmd = [sys.executable, grade_py,
           "--problem", problem, "--answer", answer, "--student", student]
    if key:
        cmd += ["--key", key]
    if html:
        cmd += ["--html", html]
    if json_out:
        cmd += ["--json", json_out]
    if history and os.path.isfile(history):
        cmd += ["--history", history]
    if sheets:
        cmd += ["--sheets", ",".join(str(s).strip() for s in sheets
                                     if str(s).strip())]
    return cmd

def run_grading(grade_py, problem, answer, student,
                key=None, html=None, json_out=None, history=None,
                sheets=None, timeout=600):
    cmd = build_grade_cmd(grade_py, problem, answer, student, key=key,
                          html=html, json_out=json_out, history=history,
                          sheets=sheets)
    try:
        proc = subprocess.run(
            cmd, capture_output=True, text=True,
            encoding="utf-8", errors="replace", timeout=timeout)
    except subprocess.TimeoutExpired:
        return None, "", "채점 시간이 초과되었습니다.", -1
    except OSError as e:
        return None, "", f"채점 프로세스를 실행할 수 없습니다: {e}", -1
    result = None
    if json_out and os.path.isfile(json_out):
        try:
            with open(json_out, encoding="utf-8") as f:
                result = json.load(f)
        except Exception:
            result = None
    if result is None and proc.returncode == 0:
        m = re.search(r"총점\s+100\s+(\d+)", proc.stdout)
        if m:
            result = {"total": int(m.group(1)), "pass_line": PASS_LINE,
                      "passed": int(m.group(1)) >= PASS_LINE, "sheets": []}
    return result, proc.stdout, proc.stderr, proc.returncode

def open_file(path):
    path = os.path.abspath(path)
    try:
        if sys.platform == "win32":
            os.startfile(path)
            return True, ""
        for opener in ("xdg-open", "open"):
            exe = shutil.which(opener)
            if exe:
                subprocess.Popen([exe, path], stdout=subprocess.DEVNULL,
                                 stderr=subprocess.DEVNULL)
                return True, ""
        webbrowser.open("file:///" + path.replace(os.sep, "/"))
        return True, ""
    except Exception as e:
        return False, str(e)

CREATE_NO_WINDOW = 0x08000000
DETACHED_PROCESS = 0x00000008
CREATE_NEW_PROCESS_GROUP = 0x00000200
EXCEL_OPEN_DELAY_MS = 1000
EXCEL_CHECK_DELAY_MS = 5000
_EXCEL_EXE_CACHE = [None, False]

def _exe_from_command(cmd):
    text = str(cmd or "").strip()
    if not text:
        return None
    if text.startswith('"'):
        end = text.find('"', 1)
        cand = text[1:end] if end > 0 else text[1:]
    else:
        m = re.match(r"(.+?\.exe)(?=\s|$)", text, re.IGNORECASE)
        cand = m.group(1) if m else text.split()[0]
    cand = os.path.expandvars(cand.strip())
    return cand or None

def _excel_registry_candidates():
    try:
        import winreg
    except ImportError:
        return []
    out = []

    def read(root, key, value=""):
        try:
            with winreg.OpenKey(root, key) as k:
                v, _t = winreg.QueryValueEx(k, value)
                return str(v)
        except OSError:
            return None

    for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        for key in (r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths"
                    r"\excel.exe",
                    r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion"
                    r"\App Paths\excel.exe"):
            v = read(root, key)
            if v:
                out.append(os.path.expandvars(v.strip('"')))
    cur = read(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CurVer")
    progids = [cur] if cur else []
    progids += ["Excel.Application.16", "Excel.Application.15",
                "Excel.Sheet.12", "Excel.SheetMacroEnabled.12", "Excel.Sheet.8"]
    for pid in progids:
        exe = _exe_from_command(
            read(winreg.HKEY_CLASSES_ROOT, rf"{pid}\shell\Open\command"))
        if exe:
            out.append(exe)
    for ver in ("16.0", "15.0", "14.0"):
        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            v = read(root, rf"SOFTWARE\Microsoft\Office\{ver}\Excel"
                     r"\InstallRoot", "Path")
            if v:
                out.append(os.path.join(v, "EXCEL.EXE"))
    return out

def _excel_common_paths():
    out = []
    subs = (r"Microsoft Office\root\Office16", r"Microsoft Office\Office16",
            r"Microsoft Office\root\Office15", r"Microsoft Office\Office15",
            r"Microsoft Office 15\root\Office15", r"Microsoft Office\Office14")
    for var in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432"):
        root = os.environ.get(var)
        if root:
            out.extend(os.path.join(root, sub, "EXCEL.EXE") for sub in subs)
    la = os.environ.get("LOCALAPPDATA")
    if la:
        out.append(os.path.join(la, r"Microsoft\WindowsApps\excel.exe"))
    return out

def find_excel_exe(refresh=False):
    if sys.platform != "win32":
        return None
    if _EXCEL_EXE_CACHE[1] and not refresh:
        return _EXCEL_EXE_CACHE[0]
    found = None
    try:
        for cand in _excel_registry_candidates() + _excel_common_paths():
            if cand and os.path.isfile(cand):
                found = os.path.abspath(cand)
                break
        if not found:
            w = shutil.which("excel") or shutil.which("EXCEL.EXE")
            if w:
                found = os.path.abspath(w)
    except Exception as e:
        log_error("Excel 경로 탐색", e)
    _EXCEL_EXE_CACHE[:] = [found, True]
    return found

def open_workbook(path, prefer_excel=True):
    path = os.path.abspath(path)
    err = ""
    if sys.platform == "win32" and prefer_excel:
        exe = find_excel_exe()
        if exe:
            try:
                proc = subprocess.Popen([exe, path], close_fds=True)
                return True, f"Excel 직접 실행 ({exe})", "", proc
            except Exception as e:
                err = f"Excel 직접 실행 실패({exe}): {e}"
        else:
            err = "EXCEL.EXE를 찾지 못함"
    ok, err2 = open_file(path)
    method = "기본 프로그램(os.startfile)" if sys.platform == "win32" \
        else "기본 프로그램"
    return ok, method, "; ".join(x for x in (err, err2) if x), None

def excel_running():
    if sys.platform != "win32":
        return None
    try:
        proc = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV",
             "/NH"], capture_output=True, text=True, encoding="utf-8",
            errors="replace", timeout=10, creationflags=CREATE_NO_WINDOW)
        return "excel.exe" in (proc.stdout or "").lower()
    except Exception:
        return None

def workbook_lock_file(path):
    d, base = os.path.split(os.path.abspath(path))
    try:
        for fn in os.listdir(d):
            if not fn.startswith("~$"):
                continue
            tail = fn[2:]
            if tail and base.endswith(tail) and len(tail) >= len(base) - 3:
                return os.path.join(d, fn)
    except OSError:
        pass
    return None

def check_workbook_open(path, proc=None):
    lock = workbook_lock_file(path)
    if lock:
        return "open", f"잠금 파일 확인({os.path.basename(lock)})"
    rc = None
    if proc is not None:
        try:
            rc = proc.poll()
        except Exception:
            rc = None
    running = excel_running()
    rc_note = f"직접 실행 프로세스 종료 코드 {rc}" if rc is not None else ""
    if running is True:
        return "open", "EXCEL.EXE 실행 중" + (f" ({rc_note})" if rc_note else "")
    if running is False:
        return "closed", "EXCEL.EXE 프로세스 없음" + (
            f", {rc_note}" if rc_note else "")
    if rc is not None and rc != 0:
        return "closed", rc_note
    return "unknown", "확인 불가 (Windows 외 환경)"

def restart_command(argv=None, executable=None, base_dir=None):
    argv = list(sys.argv if argv is None else argv)
    exe = executable or sys.executable
    base_dir = base_dir or BASE_DIR
    entry = os.path.join(base_dir, "시험장.py")
    a0 = os.path.abspath(argv[0]) if argv and argv[0] else ""
    if a0.lower().endswith(".pyw") and os.path.isfile(a0):
        entry = a0
    args = [exe, entry] + [a for a in argv[1:] if a != "--smoke"]
    kwargs = {"cwd": base_dir, "close_fds": True}
    if sys.platform == "win32":
        flags = CREATE_NEW_PROCESS_GROUP
        if os.path.basename(exe).lower().startswith("pythonw"):
            flags |= DETACHED_PROCESS
        else:
            flags |= CREATE_NO_WINDOW
        kwargs["creationflags"] = flags
    else:
        kwargs["start_new_session"] = True
    return args, kwargs

MANUAL_PIP = ("수동 설치: 명령 프롬프트에서  py -m pip install openpyxl\n"
              "(위 명령이 안 되면: python -m pip install openpyxl)")

def _pip_python():
    exe = sys.executable
    base = os.path.basename(exe).lower()
    if base.startswith("pythonw"):
        cand = os.path.join(os.path.dirname(exe),
                            base.replace("pythonw", "python", 1))
        if os.path.isfile(cand):
            return cand
    return exe

def install_openpyxl(timeout=600):
    kwargs = {}
    if sys.platform == "win32":
        kwargs["creationflags"] = 0x08000000
    try:
        proc = subprocess.run(
            [_pip_python(), "-m", "pip", "install", "openpyxl"],
            capture_output=True, text=True, encoding="utf-8",
            errors="replace", timeout=timeout, **kwargs)
    except Exception as e:
        return False, str(e)
    log = ((proc.stdout or "") + "\n" + (proc.stderr or "")).strip()
    if len(log) > 1200:
        log = log[-1200:]
    return proc.returncode == 0, log

def classify_grading_error(rc, text):
    t = text or ""
    if rc == 3 or "ModuleNotFoundError" in t or "ImportError" in t \
            or "openpyxl 라이브러리가 설치되어" in t:
        return "module"
    if "FileNotFoundError" in t or "BadZipFile" in t \
            or "찾을 수 없습니다" in t or "열 수 없습니다" in t:
        return "file"
    return "other"

EXAM_DATES_SETTING = "시험일목록"
EXAM_DATE_SETTING = "시험일"
EXAM_DATE2_SETTING = "시험일2"
ROUTINE_DAYS = 7

GOAL_TIERS = [70, 80, 90, 100]
DAILY_CAP = 1

PLAN_EXAM1, PLAN_EXAM2, PLAN_AFTER = 8, 9, 10
PLAN_ORDER = ([0] + list(range(1, ROUTINE_DAYS + 1))
              + [PLAN_EXAM1, PLAN_EXAM2, PLAN_AFTER])
AUTO = "자동선택"

def parse_iso_date(value):
    if isinstance(value, date):
        return value
    s = str(value or "").strip()[:10]
    try:
        return date(*(int(x) for x in s.split("-")))
    except (TypeError, ValueError):
        return None

_EXAM_CACHE = {"path": None, "mtime": None, "value": []}

def _norm_date_list(values):
    out = []
    for x in values or []:
        d = parse_iso_date(x)
        if d and d not in out:
            out.append(d)
    return sorted(out)

def _exam_dates_from_cfg(cfg):
    sec = _cfg_section(cfg, "_설정")
    raw = sec.get(EXAM_DATES_SETTING)
    if isinstance(raw, (list, tuple)):
        return _norm_date_list(raw)
    if isinstance(raw, str):
        return _norm_date_list([raw])
    return _norm_date_list([sec.get(EXAM_DATE_SETTING),
                            sec.get(EXAM_DATE2_SETTING)])

def exam_dates(cfg=None, path=None):
    if cfg is not None:
        return _exam_dates_from_cfg(cfg)
    p = path or SET_CONFIG_PATH
    try:
        m = os.path.getmtime(p)
    except OSError:
        m = None
    if _EXAM_CACHE["path"] == p and _EXAM_CACHE["mtime"] == m:
        return list(_EXAM_CACHE["value"])
    v = _exam_dates_from_cfg(load_set_config(p))
    _EXAM_CACHE.update(path=p, mtime=m, value=list(v))
    return list(v)

def _exams(exam=None):
    if exam is None:
        return exam_dates()
    if isinstance(exam, date):
        exam = [exam]
    return _norm_date_list(exam)

def effective_exams(exam=None, today=None):
    ex = _exams(exam)
    if not ex:
        return []
    today = today or date.today()
    return [d for d in ex if d >= today] or ex

def save_exam_date_list(dates, path=None):
    p = path or SET_CONFIG_PATH
    cfg = load_set_config(p)
    out = _norm_date_list(dates)
    sec = cfg.setdefault("_설정", {})
    sec[EXAM_DATES_SETTING] = [d.isoformat() for d in out]
    sec[EXAM_DATE_SETTING] = out[0].isoformat() if out else None
    sec[EXAM_DATE2_SETTING] = out[1].isoformat() if len(out) > 1 else None
    save_set_config(cfg, p)
    _EXAM_CACHE.update(path=None, mtime=None, value=[])
    return out

def save_exam_dates(first, second=None, path=None):
    d1, d2 = parse_iso_date(first), parse_iso_date(second)
    if d1 is None:
        d1, d2 = None, None
    if d2 is not None and d1 is not None and d2 <= d1:
        d2 = None
    return save_exam_date_list([d for d in (d1, d2) if d], path)

def add_exam_date(value, path=None):
    d = parse_iso_date(value)
    if d is None:
        return exam_dates(path=path), False
    cur = exam_dates(path=path)
    if d in cur:
        return cur, False
    return save_exam_date_list(cur + [d], path), True

def remove_exam_date(value, path=None):
    d = parse_iso_date(value)
    cur = exam_dates(path=path)
    if d is None or d not in cur:
        return cur, False
    return save_exam_date_list([x for x in cur if x != d], path), True

def clear_exam_dates(path=None):
    return save_exam_date_list([], path)

def exam_date_set(exam=None):
    return bool(_exams(exam))

def routine_start(exam=None, today=None):
    ex = effective_exams(exam, today)
    return ex[0] - timedelta(days=ROUTINE_DAYS) if ex else None

def routine_tag(exam=None, today=None):
    ex = effective_exams(exam, today)
    return ex[0].isoformat() if ex else ""

def progress_ns(exam=None, today=None):
    ex = effective_exams(exam, today)
    return "r" + (ex[0].strftime("%Y%m%d") if ex else "0")

EXAM_UNSET_TODO = (
    "시험일이 아직 설정되지 않았습니다. [시험일 설정]에서 시험 날짜를 "
    "입력하면 그 날짜 기준으로 D-7 ~ D-1 일주일 일정이 만들어집니다.")

ROUTINE_PLAN = {
    0: {"제목": "자유 연습", "종류": "안내",
        "할일": "아직 일주일 일정이 시작되기 전입니다. 지금은 자유 연습 "
               "구간이니 가진 세트를 편하게 풀어 보세요. 프로그램 실행·채점 "
               "흐름 확인, 문제·정답 파일 준비, 루틴 웹페이지 즐겨찾기까지 "
               "해 두면 좋습니다."},
    PLAN_EXAM1: {"제목": "시험일", "종류": "안내",
                 "할일": "오늘이 시험일입니다. 수험표·신분증 확인, 고사장 "
                        "30분 전 도착. 저장은 Ctrl+S 수시로, 계산작업은 한 "
                        "문제 3분 넘기면 다음으로 넘어가세요. 화이팅!"},
    PLAN_EXAM2: {"제목": "2차 시험일", "종류": "안내",
                 "할일": "앞선 시험에 이어 오늘도 시험일입니다. 아침에 실수 "
                        "노트와 지난 시험 복기 메모만 한 번 훑고 출발하세요. "
                        "그 항목만 지키면 됩니다 — 화이팅!"},
    PLAN_AFTER: {"제목": "루틴 완주", "종류": "안내",
                 "할일": "일주일 루틴과 시험을 완주했습니다. 수고 많았습니다! "
                        "결과와 관계없이 쌓은 실력은 남습니다. 다음 시험을 "
                        "준비한다면 [시험일 설정]에서 새 시험일을 넣으세요."},
}
for _no in range(1, ROUTINE_DAYS + 1):
    ROUTINE_PLAN[_no] = {
        "제목": f"Day {_no}", "종류": "모의", "세트": [AUTO], "목표": 70,
        **({"특별": ["실수노트"]} if _no == ROUTINE_DAYS else {}),
    }
del _no

PREP_STEPS = [
    {"이름": "시험일 설정", "형": "안내", "분": 2,
     "설명": "[시험일 설정]에서 시험 날짜를 넣으면 그 날 기준 D-7 ~ D-1 "
            "일주일 일정이 자동으로 만들어집니다. 나중에 언제든 바꿀 수 "
            "있습니다."},
    {"이름": "내 문제·정답 파일 넣기", "형": "안내", "분": 10,
     "설명": "학습 폴더에 '<이름>_문제.xlsx' 와 '<이름>_정답.xlsx' 를 같은 "
            "이름으로 넣으면 한 세트로 인식됩니다(문제지 PDF는 "
            "'<이름>_문제지.pdf'). 아래 목록에 보이지 않으면 "
            "[세트 인식 진단]에서 이유를 확인하세요."},
    {"이름": "프로그램 실행·채점 확인", "형": "안내", "분": 15,
     "설명": "아무 세트나 [시험 시작]으로 열고 바로 제출해 채점까지 한 번 "
            "돌려보세요. openpyxl 설치 안내가 뜨면 설치합니다."},
    {"이름": "루틴 웹페이지 즐겨찾기", "형": "안내", "분": 5,
     "설명": "웹 루틴 페이지에서 실수 노트·함수 사전 위치를 확인하고 "
            "즐겨찾기에 추가하세요. 하루 1세트, 완주 + 채점이면 클리어입니다."},
]

def set_intake_help(root=None):
    where = f"\n넣는 곳: {root}" if root else ""
    return (
        "가지고 있는 문제 파일로 일정이 만들어집니다. 학습 폴더(하위 폴더 "
        "포함)에 같은 이름의 두 파일을 넣어 주세요.\n"
        "  · <이름>_문제.xlsx  (풀 파일)\n"
        "  · <이름>_정답.xlsx  (채점 기준)\n"
        "문제지 PDF가 있으면 <이름>_문제지.pdf 로 같이 넣으면 됩니다. "
        "기대값 JSON은 없어도 되고, 없으면 기본 배점으로 채점합니다.\n"
        "예) 2026모의1회_문제.xlsx · 2026모의1회_정답.xlsx"
        + where +
        "\n파일을 넣고 [세트 다시 스캔]을 누르세요. 목록에 보이지 않으면 "
        "[세트 인식 진단]에 이유가 나옵니다.")

def routine_date_for(no, exam=None, today=None):
    ex = effective_exams(exam, today)
    if not ex:
        return None
    no = int(no)
    if 1 <= no <= ROUTINE_DAYS:
        return ex[0] - timedelta(days=ROUTINE_DAYS + 1 - no)
    if no == PLAN_EXAM1:
        return ex[0]
    if no == PLAN_EXAM2:
        return ex[1] if len(ex) > 1 else None
    return None

def routine_day_no(today=None, exam=None):
    today = today or date.today()
    all_ex = _exams(exam)
    if not all_ex:
        return 0
    if today in all_ex:
        return PLAN_EXAM1 if all_ex.index(today) == 0 else PLAN_EXAM2
    ex = effective_exams(all_ex, today)
    if today > max(ex):
        return PLAN_AFTER
    for no in PLAN_ORDER:
        d = routine_date_for(no, ex, today)
        if d is not None and d == today:
            return no
    return 0

def plan_day_tag(no):
    return f"d{int(no):02d}" if 0 <= int(no) <= ROUTINE_DAYS else None

def dday_num(today=None, exam=None):
    today = today or date.today()
    ex = effective_exams(exam, today)
    if not ex:
        return None
    return (ex[0] - today).days

def dday_label(n):
    if n is None:
        return ""
    return "D-day" if n == 0 else (f"D-{n}" if n > 0 else f"D+{-n}")

def dday_text(today=None, exam=None):
    today = today or date.today()
    ex = effective_exams(exam, today)
    if not ex:
        return "시험일 미설정"
    parts = [f"시험 {dday_label((ex[0] - today).days)}"]
    if len(ex) > 1:
        parts.append(f"2차 {dday_label((ex[1] - today).days)}")
    return " · ".join(parts)

def scored_exam_records(records):
    out = []
    for r in records or []:
        if r.get("mode") in ("부분연습", "오답재풀이"):
            continue
        if not isinstance(r.get("점수"), (int, float)):
            continue
        if _record_date(r) is None:
            continue
        out.append(r)
    return out

def best_total(records):
    scores = [r["점수"] for r in scored_exam_records(records)]
    return max(scores) if scores else None

def level_info(records):
    return level_from_best(best_total(records))

def goal_label(goal):
    if not goal:
        return ""
    return f"목표 {int(goal)}점(" \
        + ("합격선" if int(goal) == PASS_LINE else "승급") + ")"

def attempt_dates(records):
    return {_record_date(r) for r in scored_exam_records(records)}

def streak_info(records, today=None, exam=None):
    today = today or date.today()
    ex = _exams(exam)
    days = {d for d in attempt_dates(records)
            if d and d <= today and d not in ex}
    today_done = today in days
    cur = 0
    d = today if today_done else today - timedelta(days=1)
    while d in days:
        cur += 1
        d -= timedelta(days=1)
    best = run = 0
    d = min(days) if days else today
    while d <= today:
        run = run + 1 if d in days else 0
        best = max(best, run)
        d += timedelta(days=1)
    return {"current": cur, "best": best, "today_done": today_done}

SCORE_BANDS = [
    {"id": "none", "name": "미응시", "min": None, "max": None,
     "desc": "아직 응시 기록이 없는 세트 — 먼저 한 번 풀어 보세요"},
    {"id": "fail", "name": "미달", "min": 0, "max": 69,
     "desc": "최고점이 합격선 70점에 못 미친 세트 — 1순위로 다시 풀기"},
    {"id": "pass", "name": "합격", "min": 70, "max": 79,
     "desc": "합격선은 넘겼지만 아직 80점 미만"},
    {"id": "skilled", "name": "숙련", "min": 80, "max": 89,
     "desc": "80점대 — 90점까지 한 걸음"},
    {"id": "high", "name": "고득점", "min": 90, "max": 99,
     "desc": "90점대 — 만점까지 한 걸음"},
    {"id": "perfect", "name": "만점", "min": 100, "max": 100,
     "desc": "100점 — 더 올릴 점수가 없는 세트"},
]
BAND_IDS = [b["id"] for b in SCORE_BANDS]
BAND_LABELS = {b["id"]: b["name"] for b in SCORE_BANDS}

def score_band(best):
    if isinstance(best, bool) or not isinstance(best, (int, float)):
        return "none"
    if best >= 100:
        return "perfect"
    if best >= 90:
        return "high"
    if best >= 80:
        return "skilled"
    if best >= PASS_LINE:
        return "pass"
    return "fail"

def band_label(band_id):
    return BAND_LABELS.get(str(band_id or ""), "")

TIER_LABELS = {70: "합격", 80: "숙련", 90: "고득점", 100: "만점"}
FAST_CLEAR_SEC = 40 * 60
CALC_AREA = "계산작업"
BASE_AREA_POINTS = 60
CALC_AREA_POINTS = 40

ACHIEVEMENT_DEFS = [
    {"id": "first", "icon": "완주", "group": "시작", "name": "첫 완주",
     "desc": "모의고사 한 세트를 끝까지 풀고 채점했다"},
    {"id": "pass70", "icon": "70", "group": "점수", "name": "첫 합격",
     "desc": "합격선 70점을 처음 넘었다"},
    {"id": "tier80", "icon": "80", "group": "점수", "name": "80점 돌파",
     "desc": "한 세트에서 80점 이상을 받았다"},
    {"id": "tier90", "icon": "90", "group": "점수", "name": "90점 돌파",
     "desc": "한 세트에서 90점 이상을 받았다"},
    {"id": "tier100", "icon": "100", "group": "점수",
     "name": "만점 · 전 영역 무실점",
     "desc": "100점 — 모든 영역에서 한 점도 잃지 않았다"},
    {"id": "streak3", "icon": "3일", "group": "연속", "name": "3일 연속",
     "desc": "사흘 연속으로 응시했다"},
    {"id": "streak5", "icon": "5일", "group": "연속", "name": "5일 연속",
     "desc": "닷새 연속으로 응시했다"},
    {"id": "calc40", "icon": "계산", "group": "정확도", "name": "계산작업 만점",
     "desc": "계산작업 5문제를 모두 맞혀 40/40점을 받았다"},
    {"id": "base60", "icon": "무실점", "group": "정확도",
     "name": "기본·분석·기타 무실점",
     "desc": "계산작업을 뺀 세 영역에서 60점을 모두 지켰다"},
    {"id": "sets5", "icon": "5세트", "group": "완주량", "name": "세트 5개 완주",
     "desc": "서로 다른 세트를 5개 완주했다"},
    {"id": "sets9", "icon": "9세트", "group": "완주량", "name": "세트 9개 완주",
     "desc": "서로 다른 세트를 9개 완주했다"},
    {"id": "fast40", "icon": "40분", "group": "속도", "name": "40분 내 완주",
     "desc": "실제 시험 시간(40분) 안에 제출했다"},
]
ACHIEVEMENT_IDS = [a["id"] for a in ACHIEVEMENT_DEFS]

def tier_label(goal):
    try:
        return TIER_LABELS[int(goal)]
    except (TypeError, ValueError, KeyError):
        return ""

def level_from_best(best):
    achieved = [t for t in GOAL_TIERS if best is not None and best >= t]
    perfect = GOAL_TIERS[-1] in achieved
    goal = GOAL_TIERS[-1] if perfect \
        else next(t for t in GOAL_TIERS if t not in achieved)
    return {"tier": "perfect" if perfect else goal, "current_goal": goal,
            "best_total": best, "achieved": achieved,
            "next_goal": None if perfect else goal}

def parse_elapsed(value):
    if isinstance(value, bool):
        return 0
    if isinstance(value, (int, float)):
        return max(0, int(value))
    text = str(value or "").strip()
    if not text:
        return 0
    m = re.search(r"(?:(\d+)\s*시간)?\s*(\d+)\s*분\s*(?:(\d+)\s*초)?", text)
    if m and m.group(2):
        h = int(m.group(1) or 0)
        return h * 3600 + int(m.group(2)) * 60 + int(m.group(3) or 0)
    parts = text.split(":")
    if len(parts) in (2, 3) and all(p.strip().isdigit() for p in parts):
        nums = [int(p) for p in parts]
        if len(nums) == 2:
            return nums[0] * 60 + nums[1]
        return nums[0] * 3600 + nums[1] * 60 + nums[2]
    m2 = re.fullmatch(r"(\d+)\s*초?", text)
    return int(m2.group(1)) if m2 else 0

def _ach_records(records):
    out = []
    for r in records or []:
        if not isinstance(r, dict):
            continue
        mode = str(r.get("mode") or "시험")
        if mode not in ("시험", "수동"):
            continue
        total = r.get("total")
        if isinstance(total, bool) or not isinstance(total, (int, float)):
            continue
        d = str(r.get("date") or "")
        if len(d) != 10:
            continue
        st = r.get("set") if isinstance(r.get("set"), dict) else {}
        sheets = []
        for s in (r.get("sheets") or []):
            if isinstance(s, dict) \
                    and isinstance(s.get("alloc"), (int, float)) \
                    and isinstance(s.get("earned"), (int, float)) \
                    and not isinstance(s.get("alloc"), bool):
                sheets.append({"name": str(s.get("name") or ""),
                               "alloc": s["alloc"], "earned": s["earned"]})
        es = r.get("elapsed_sec")
        out.append({
            "id": str(r.get("id") or ""), "date": d,
            "time": str(r.get("time") or ""), "mode": mode, "total": total,
            "norm": str(st.get("norm") or "") or str(st.get("name") or ""),
            "sheets": sheets,
            "elapsed_sec": int(es) if isinstance(es, (int, float))
            and not isinstance(es, bool) and es > 0 else 0})
    out.sort(key=lambda r: (r["date"], r["time"], r["id"]))
    return out

def _area_totals(rec):
    out = {}
    for s in rec.get("sheets") or []:
        cat = _category_of_sheet(s.get("name"))
        cur = out.setdefault(cat, [0, 0])
        cur[0] += s["alloc"]
        cur[1] += s["earned"]
    return out

def _iso_next(d):
    try:
        return (date.fromisoformat(d) + timedelta(days=1)).isoformat()
    except (TypeError, ValueError):
        return None

def achievements(records, exam=None):
    recs = _ach_records(records)
    got = {}

    def mark(aid, when):
        if when and aid not in got:
            got[aid] = when

    seen, days = [], []
    for r in recs:
        d, total = r["date"], r["total"]
        mark("first", d)
        for aid, need in (("pass70", PASS_LINE), ("tier80", 80),
                          ("tier90", 90), ("tier100", 100)):
            if total >= need:
                mark(aid, d)
        areas = _area_totals(r)
        calc = areas.get(CALC_AREA)
        if calc and calc[0] >= CALC_AREA_POINTS and calc[1] >= calc[0]:
            mark("calc40", d)
        others = [v for k, v in areas.items() if k != CALC_AREA]
        if others and sum(v[0] for v in others) >= BASE_AREA_POINTS \
                and all(v[1] >= v[0] for v in others):
            mark("base60", d)
        if r["norm"] and r["norm"] not in seen:
            seen.append(r["norm"])
            if len(seen) >= 5:
                mark("sets5", d)
            if len(seen) >= 9:
                mark("sets9", d)
        if r["mode"] == "시험" and 0 < r["elapsed_sec"] <= FAST_CLEAR_SEC:
            mark("fast40", d)
        if d not in days:
            days.append(d)
    exam_iso = [x.isoformat() for x in _exams(exam)]
    run, prev = 0, None
    for d in sorted(days):
        if d in exam_iso:
            continue
        run = run + 1 if (prev is not None and _iso_next(prev) == d) else 1
        prev = d
        if run >= 3:
            mark("streak3", d)
        if run >= 5:
            mark("streak5", d)
    out = []
    for a in ACHIEVEMENT_DEFS:
        out.append(dict(a, done=a["id"] in got, when=got.get(a["id"])))
    return out

def achievement_summary(records):
    items = achievements(records)
    done = [a for a in items if a["done"]]
    latest = max(done, key=lambda a: a["when"]) if done else None
    return {"done": len(done), "total": len(items), "items": items,
            "latest": latest}

def new_achievements(before, after):
    had = {a["id"] for a in achievements(before) if a["done"]}
    return [a for a in achievements(after) if a["done"] and a["id"] not in had]

def serialized_best(records):
    scores = [r["total"] for r in _ach_records(records)]
    return max(scores) if scores else None

def gauge_fraction(best, goal):
    try:
        g = float(goal or 0)
    except (TypeError, ValueError):
        return 0.0
    if g <= 0:
        return 0.0
    b = float(best or 0)
    return max(0.0, min(1.0, b / g))

def progress_delta(before, after):
    prev_best = serialized_best(before)
    best = serialized_best(after)
    lv_b, lv_a = level_from_best(prev_best), level_from_best(best)
    summ = achievement_summary(after)
    return {
        "prev_best": prev_best, "best": best,
        "best_delta": None if prev_best is None else (best or 0) - prev_best,
        "level_before": lv_b, "level_after": lv_a,
        "tier_up": [t for t in lv_a["achieved"] if t not in lv_b["achieved"]],
        "new_badges": new_achievements(before, after),
        "badges_done": summ["done"], "badges_total": summ["total"],
    }

def set_progress(norm, records):
    key = str(norm or "")
    recs = [r for r in _ach_records(records) if r["norm"] == key] if key else []
    scores = [r["total"] for r in recs]
    best = max(scores) if scores else None
    prev = scores[:-1]
    prev_best = max(prev) if prev else None
    improved = bool(scores) and (prev_best is None or scores[-1] > prev_best)
    return {
        "attempts": len(scores),
        "best": best,
        "band": score_band(best),
        "scores": scores,
        "last_date": recs[-1]["date"] if recs else None,
        "prev_best": prev_best,
        "prev_band": score_band(prev_best) if prev_best is not None else "none",
        "best_delta": (best - prev_best) if (improved and prev_best is not None)
        else None,
        "improved": improved,
    }

def score_board(sets, records):
    rows = []
    for s in sets or []:
        norm = str((s or {}).get("norm") or "")
        row = {"name": str((s or {}).get("name") or ""), "norm": norm}
        row.update(set_progress(norm, records))
        rows.append(row)
    bands = []
    for i, b in enumerate(SCORE_BANDS):
        mine = [(k, r) for k, r in enumerate(rows) if r["band"] == b["id"]]
        if b["id"] == "none":
            order = mine
        else:
            order = sorted(mine, key=lambda kr: (kr[1]["best"], kr[1]["name"]))
        bands.append({"id": b["id"], "name": b["name"], "desc": b["desc"],
                      "count": len(order),
                      "sets": [r["norm"] for _k, r in order]})
    return {"sets": rows, "bands": bands}

def band_summary_text(bands):
    parts = [f"{b['name']} {b['count']}" for b in bands or [] if b["count"]]
    return " · ".join(parts) if parts else "세트 없음"

def best_update_text(row):
    if not row or not row.get("improved") or row.get("prev_best") is None:
        return ""
    return (f"최고 {row['prev_best']:g} → {row['best']:g} "
            f"(+{row['best_delta']:g})")

def mandatory_step_indexes(steps):
    return [i for i, s in enumerate(steps or []) if not s.get("선택")]

def plan_slot_names(plan):
    return [spec if spec != AUTO else "자동 선택"
            for spec in (plan.get("세트") or [])]

def build_day_steps(plan, slot_names=None):
    no = plan.get("no", 0)
    if plan.get("종류") != "모의":
        return [dict(s) for s in PREP_STEPS] if no == 0 else []
    slots = list(plan.get("세트") or [])
    names = list(slot_names or plan_slot_names(plan))
    goals = list(plan.get("목표들") or [])
    if len(goals) < len(slots):
        goals += [plan.get("목표")] * (len(slots) - len(goals))
    steps = []
    for k, spec in enumerate(slots):
        name = names[k] if k < len(names) else "자동 선택"
        goal = goals[k]
        g_txt = f" ({goal_label(goal)})" if goal else ""
        n_txt = f" ({k + 1}/{len(slots)})" if len(slots) > 1 else ""
        auto_txt = (" 자동 선택된 세트는 창 위쪽에 이유와 함께 표시되며 "
                    "[다른 세트로 바꾸기]로 바꿀 수 있습니다."
                    if spec == AUTO else "")
        steps.extend([
            {"이름": f"{name} 시험 모드 40분 완주{n_txt}", "형": "모의",
             "세트": spec, "슬롯": k, "목표": goal, "분": 40,
             "설명": f"{name} 40분 실전 완주{g_txt}. 제출하면 자동 채점되고 "
                    "이 단계와 다음 '채점·성적 확인' 단계가 자동으로 "
                    "체크됩니다. 응시 횟수 제한은 없어요 — 예습·복습으로 "
                    "몇 번이든 다시 풀 수 있습니다." + auto_txt},
            {"이름": f"채점·성적 확인{n_txt}", "형": "채점", "세트": spec,
             "슬롯": k, "분": 5,
             "설명": "채점이 끝나면 자동 체크되고, 결과는 루틴 페이지"
                    "([루틴 열기])에 자동으로 반영됩니다. 여기까지 하면 오늘은 "
                    "클리어입니다 — 점수는 클리어 조건이 아닙니다(낮으면 그 "
                    "세트가 이후 날짜에 자동으로 다시 배정됩니다)."},
            {"이름": f"(선택) 오답노트 확인{n_txt}", "형": "오답노트",
             "세트": spec, "슬롯": k, "분": 15, "선택": True,
             "설명": "(선택) 틀린 항목의 해설을 훑어보는 도구입니다. 채점 "
                    "리포트에서 이미 확인했다면 건너뛰어도 됩니다 — 클리어 "
                    "조건이 아닙니다."},
            {"이름": f"(선택) 오답 재풀이 15분{n_txt}", "형": "오답재풀이",
             "세트": spec, "슬롯": k, "분": 15, "선택": True,
             "설명": "(선택) 오답이 있던 시트만 새 사본(오답재풀이_*.xlsm)에서 "
                    "다시 풀어 보는 도구입니다. 시간이 남을 때만 하세요 — "
                    "클리어 조건이 아닙니다."},
        ])
    if "실수노트" in (plan.get("특별") or []):
        steps.append({"이름": "(선택) 실수 노트 정리", "형": "안내", "분": 10,
                      "선택": True,
                      "설명": "(선택) 오늘까지의 오답에서 반복된 실수를 웹 루틴 "
                             "실수 노트에 정리하세요. 내일 시험장에서 볼 "
                             "마지막 체크리스트입니다."})
    steps.append({"이름": "(선택) 함수 퀴즈", "형": "안내", "분": 10,
                  "선택": True, "웹탭": "quiz",
                  "설명": "(선택) 루틴 페이지 함수 퀴즈 10문제 — 오늘 틀린 "
                         "함수가 있으면 그 함수부터. [웹에서 퀴즈 풀기]로 "
                         "퀴즈 탭이 열리고, 페이지에서 ⑤ 체크하면 이 단계도 "
                         "자동 체크됩니다. 건너뛰어도 됩니다."})
    return steps

def _default_todo(plan):
    names = plan_slot_names(plan)
    txt = "모의고사 " + " + ".join(names) + " 40분 완주 → 채점"
    if plan.get("목표"):
        txt += " · " + goal_label(plan["목표"])
    txt += " (오답노트·오답 재풀이·함수 퀴즈는 선택)"
    if plan.get("메모"):
        txt += f"  [{plan['메모']}]"
    return txt

def plan_for_day(no, slot_names=None, exam=None):
    no = int(no)
    if no not in ROUTINE_PLAN:
        no = max(0, min(PLAN_AFTER, no))
        if no not in ROUTINE_PLAN:
            no = 0
    plan = dict(ROUTINE_PLAN[no])
    plan["no"] = no
    plan["날짜"] = routine_date_for(no, exam)
    plan["세트"] = list(plan.get("세트") or [])
    if no == 0 and not exam_date_set(exam):
        plan = dict(plan, 제목="시험일 설정", 할일=EXAM_UNSET_TODO)
    if not plan.get("할일"):
        plan["할일"] = _default_todo(plan)
    plan["스텝"] = build_day_steps(plan, slot_names)
    return plan

ROUTINE_STEPS = {no: plan_for_day(no)["스텝"] for no in PLAN_ORDER}

def plan_title(plan, today=None, set_names=None, exam=None):
    d = plan.get("날짜")
    d_txt = f"{d.month}/{d.day}({'월화수목금토일'[d.weekday()]})" if d else ""
    no = plan["no"]
    if 1 <= no <= ROUTINE_DAYS:
        head = f"Day {no}"
        if d is not None:
            head += f" (D-{ROUTINE_DAYS + 1 - no})"
        names = list(set_names or plan.get("세트표시") or plan_slot_names(plan))
        if plan.get("적응형"):
            body = f"{plan['제목']}: {' + '.join(names)}" if names \
                else plan["제목"]
            if plan.get("남은학습일") is not None and plan.get("종류") == "모의":
                body += f" · 남은 학습일 {plan['남은학습일']}일"
        else:
            body = f"{plan['제목']} — {' + '.join(names)}" if names \
                else plan["제목"]
    elif no in (PLAN_EXAM1, PLAN_EXAM2):
        head, body = "시험일", plan["제목"]
    elif no == PLAN_AFTER:
        head, body = "루틴", plan["제목"]
    else:
        s = routine_start(exam)
        head = "준비"
        body = plan["제목"]
        if s is not None:
            body += (f" — {s.month}/{s.day}"
                     f"({'월화수목금토일'[s.weekday()]})부터 Day 1")
    parts = [head]
    if d_txt:
        parts.append(d_txt)
    parts.append(body)
    if today:
        parts.append(dday_text(today, exam))
    return " · ".join(parts)

SLOT_IDENTITY = {
    "2024 상시 1회": {"핵심": [["2024", "24"], ["1회"]]},
    "2024 상시 2회": {"핵심": [["2024", "24"], ["2회"]]},
    "2024 A형": {"핵심": [["a형"]]},
    "2024 B형": {"핵심": [["b형"]]},
    "코코 1회": {"핵심": [["코코"], ["1회"]]},
    "코코 2회": {"핵심": [["코코"], ["2회"]]},
    "24 2급 상시": {"핵심": [["24", "2024"], ["2급"]],
                 "보조": [["모의", "실기", "상시"]]},
    "컴활 2급 상시": {"핵심": [["컴활"], ["2급"], ["상시"]], "연도제외": True},
    "2026 1회": {"핵심": [["2026"], ["1회"]]},
}
SLOT_TIER_LABEL = {3: "전체 토큰 일치", 2: "식별 토큰 일치", 1: "부분 일치"}

def slot_identity_spec(text):
    if text in SLOT_IDENTITY:
        return SLOT_IDENTITY[text]
    toks = set_tokens(text)
    rest = sorted(t for t in toks if not _YEAR4_RE.fullmatch(t)) or sorted(toks)
    return {"핵심": [[t] for t in rest]}

def slot_identity_tokens(text):
    spec = slot_identity_spec(text)
    return {"/".join(g) for g in spec["핵심"]} | \
        {"/".join(g) for g in spec.get("보조", [])}

def _set_tokens_of(s):
    toks = set_tokens(s["name"]) | set_tokens(os.path.basename(s["dir"] or ""))
    for d in s.get("중복") or []:
        toks |= set_tokens(d.get("name") or "")
        toks |= set_tokens(os.path.basename(d.get("dir") or ""))
    return toks

def _has_token(st, t):
    if _is_year_token(t):
        return bool(_year_tokens({t}) & _year_tokens(st))
    return t in st

def _group_hit(st, group):
    for t in group:
        if _has_token(st, t):
            return t
    return None

def slot_candidate_score(spec, s, st=None):
    toks = set_tokens(spec)
    st = st if st is not None else _set_tokens_of(s)
    if _token_conflict(toks, st):
        return None
    rule = slot_identity_spec(spec)
    if rule.get("연도제외") and _year_tokens(st):
        return None
    hits = []
    for g in rule["핵심"]:
        h = _group_hit(st, g)
        if h is None:
            return None
        hits.append(h)
    missing = []
    for g in rule.get("보조", []):
        h = _group_hit(st, g)
        if h is None:
            missing.append("/".join(g))
        else:
            hits.append(h)
    if all(_has_token(st, t) for t in toks):
        tier, why = 3, "전체 토큰 일치 (" + ", ".join(sorted(toks)) + ")"
    elif not missing:
        tier, why = 2, "식별 토큰 일치 (" + ", ".join(sorted(set(hits))) + ")"
    else:
        tier, why = 1, ("부분 일치 (" + ", ".join(sorted(set(hits))) +
                        " — " + "·".join(missing) + " 없음)")
    return (tier, _numeric_shared(toks, st), _shared_count(toks, st)), why

def all_slot_specs():
    specs = list(SLOT_IDENTITY)
    for no in sorted(ROUTINE_PLAN):
        for spec in ROUTINE_PLAN[no].get("세트") or []:
            if spec != AUTO and spec not in specs:
                specs.append(spec)
    return specs

def assign_slots(sets, specs=None, mapping=None):
    specs = list(dict.fromkeys(list(specs or []) + all_slot_specs()))
    result = {}
    taken = set()
    mapping = mapping or {}
    for spec in specs:
        ent = mapping.get(spec)
        if not isinstance(ent, dict):
            continue
        ms = set_from_mapping(ent)
        if ms is None:
            continue
        result[spec] = (ms, "직접 선택 저장됨 (" +
                        os.path.basename(ms["problem"]) + ")")
        taken.add(ms["norm"])
        for s in sets:
            if any(d.get("norm") == ms["norm"] for d in s.get("중복") or []):
                taken.add(s["norm"])
    toks = {id(s): _set_tokens_of(s) for s in sets}
    pairs = []
    for spec in specs:
        if spec in result:
            continue
        for s in sets:
            if s["norm"] in taken:
                continue
            r = slot_candidate_score(spec, s, toks[id(s)])
            if r:
                pairs.append((r[0], spec, s, r[1]))
    order = {spec: i for i, spec in enumerate(specs)}
    pairs.sort(key=lambda t: (tuple(-x for x in t[0]), order[t[1]], t[2]["name"]))
    tied = {}
    blocked = set()
    winner = {}
    for i, (sc, spec, s, why) in enumerate(pairs):
        if spec in result or spec in tied or s["norm"] in taken \
                or s["norm"] in blocked:
            continue
        rivals = [(sp2, s2) for sc2, sp2, s2, _w in pairs[i + 1:]
                  if sc2 == sc and (sp2 == spec or s2 is s)
                  and sp2 not in result and sp2 not in tied
                  and s2["norm"] not in taken and s2["norm"] not in blocked]
        if not rivals:
            result[spec] = (s, why)
            taken.add(s["norm"])
            winner[s["norm"]] = spec
            continue
        same_slot = sorted({s2["name"] for sp2, s2 in rivals if sp2 == spec})
        same_set = sorted({sp2 for sp2, s2 in rivals if s2 is s and sp2 != spec})
        if same_slot:
            tied[spec] = ("복수 후보 (동점): " +
                          ", ".join([s["name"]] + same_slot) +
                          " — [직접 선택]으로 지정하세요")
        if same_set:
            tied.setdefault(spec, f"세트 '{s['name']}'을(를) 두고 슬롯 " +
                            ", ".join(f"'{x}'" for x in same_set) +
                            "과(와) 동점 — [직접 선택]으로 지정하세요")
            for sp2 in same_set:
                tied.setdefault(sp2, f"세트 '{s['name']}'을(를) 두고 슬롯 "
                                f"'{spec}'과(와) 동점 — [직접 선택]으로 "
                                "지정하세요")
            blocked.add(s["norm"])
    for spec in specs:
        if spec in result:
            continue
        if spec in tied:
            result[spec] = (None, tied[spec])
            continue
        lost = [(s, winner[s["norm"]]) for sc, sp, s, _w in pairs
                if sp == spec and s["norm"] in winner]
        if lost:
            result[spec] = (None, "후보 없음 — " + ", ".join(
                f"'{s['name']}'은(는) 슬롯 '{w}'에 배정됨(더 잘 맞음)"
                for s, w in lost))
        else:
            idt = "·".join(sorted(slot_identity_tokens(spec)))
            result[spec] = (None, f"후보 없음 — 파일명에 {idt} 토큰을 가진 "
                                  "문제/정답 짝이 없음")
    return result

def match_slot(sets, text, mapping=None):
    return assign_slots(sets, [text], mapping)[text]

def find_set_for_tokens(sets, text):
    return match_slot(sets, text)[0]

def load_slot_mapping(path=None):
    cfg = load_set_config(path or SET_CONFIG_PATH)
    raw = _cfg_section(cfg, "_슬롯매핑")
    return {k: v for k, v in raw.items() if isinstance(v, dict)}

def save_slot_mapping(spec, s, path=None):
    p = path or SET_CONFIG_PATH
    cfg = load_set_config(p)
    cfg.setdefault("_슬롯매핑", {})[str(spec)] = {
        "name": s.get("name"),
        "problem": os.path.abspath(s["problem"]),
        "answer": os.path.abspath(s["answer"]),
        "pdf": os.path.abspath(s["pdf"]) if s.get("pdf") else None,
    }
    return save_set_config(cfg, p)

def set_from_mapping(ent):
    try:
        prob, ans = ent.get("problem"), ent.get("answer")
        if not (isinstance(prob, str) and isinstance(ans, str) and prob
                and ans and os.path.isfile(prob) and os.path.isfile(ans)):
            return None
        s = build_direct_set(prob, ans)
        if ent.get("name"):
            s["name"] = str(ent["name"])
        s["direct"] = True
        pdf = ent.get("pdf")
        if isinstance(pdf, str) and pdf and os.path.isfile(pdf):
            _attach_saved_pdf(s, pdf, bool(ent.get("pdf_확인됨")), "슬롯매핑")
        return s
    except Exception:
        return None

def scan_diagnosis_text(root, config=None, specs=None, sets=None,
                        hidden_sets=None):
    if config is None:
        config = load_set_config()
    hide_keys = hidden_set_keys(cfg=config)
    default_root = default_scan_root()
    is_default = os.path.abspath(str(root or "")) == os.path.abspath(default_root)
    lines = [f"[코코 시험장 세트 인식 진단] {__version__} · "
             f"{datetime.now().strftime('%Y-%m-%d %H:%M')}",
             f"스캔 폴더: {root}"
             + ("  (기본값 = 시험장.py 상위 폴더)" if is_default
                else f"  (직접 지정 · 기본값: {default_root})"),
             "  · 이 폴더와 하위 폴더에서만 세트를 찾습니다 — "
             "[설정]에서 바꾸거나 기본값으로 되돌릴 수 있습니다."]
    if sets is None:
        sets = scan_sets(root, config) if root and os.path.isdir(root) else []
        sets, auto_hidden = split_hidden_sets(sets, hide_keys)
        if hidden_sets is None:
            hidden_sets = auto_hidden
    hidden_sets = list(hidden_sets or [])
    if hide_keys:
        names = {}
        for s in hidden_sets:
            for k in set_hidden_keys(s):
                names[k] = s.get("name") or k
        lines.append(f"숨긴 세트 {len(hide_keys)}개 (목록·일정·점수대 보드에서 "
                     "빠짐 · 파일과 기록은 그대로): "
                     + ", ".join(f"{names.get(k, '(폴더에 없음)')} [{k}]"
                                 for k in hide_keys))
        lines.append("  · [설정] → 세트 숨기기 에서 다시 보이게 할 수 있습니다.")
    else:
        lines.append("숨긴 세트: 없음")
    lines.append("")
    owner = {}
    tie_files = {}
    for s in sets:
        for k in ("problem", "answer", "pdf", "key"):
            if s.get(k):
                owner[os.path.abspath(s[k])] = s["name"]
        for d in s.get("중복") or []:
            for k in ("problem", "answer", "pdf", "key"):
                if d.get(k):
                    owner.setdefault(os.path.abspath(d[k]),
                                     f"{s['name']} (중복·동일 내용)")
        for p in s.get("pdf_사본") or []:
            owner.setdefault(os.path.abspath(p), f"{s['name']} (같은 키 사본)")
        for k in ("pdf", "key"):
            if s.get(k + "_tie"):
                fn = s[k + "_tie"].split(" ↔ ")[0]
                tie_files.setdefault(fn, "미소속 (동점: " +
                                     s[k + "_tie"].split(" ↔ ", 1)[1] + ")")
    files = []
    if root and os.path.isdir(root):
        for dirpath, dirnames, filenames in os.walk(root):
            dirnames[:] = [d for d in dirnames
                           if not d.startswith((".", "__")) and d != "채점결과"]
            for fn in sorted(filenames):
                ext = os.path.splitext(fn)[1].lower()
                if ext not in (".xlsx", ".xlsm", ".pdf"):
                    continue
                if fn.startswith(("풀이_", "부분연습_", "오답연습_", "오답재풀이_",
                                  "채점결과", "~$", ".")):
                    continue
                files.append(os.path.join(dirpath, fn))
    unrelated = [p for p in files
                 if os.path.splitext(p)[1].lower() == ".pdf"
                 and not set_tokens(os.path.basename(p))
                 and os.path.abspath(p) not in owner]
    lines.append("== 파일 → 정규화 키 → 역할 → 소속 세트 ==")
    lines.append("파일명 | 정규화 키 | 역할 | 소속 세트 | 폴더")
    for p in files:
        if p in unrelated:
            continue
        fn = os.path.basename(p)
        ext = os.path.splitext(fn)[1].lower()
        if ext == ".pdf":
            role = "문제지(PDF)"
        else:
            role = {"problem": "문제", "answer": "정답"}.get(file_role(fn),
                                                        "미상(문제/정답 표기 없음)")
        own = owner.get(os.path.abspath(p)) or tie_files.get(fn, "미소속")
        rel = os.path.relpath(os.path.dirname(p), root) if root else ""
        lines.append(f"{fn} | {norm_set_key(fn)} | {role} | {own} | {rel}")
    if unrelated:
        lines.append(f"무관(세트 토큰 없음) PDF {len(unrelated)}개 — 연결 후보에서 "
                     "제외: " + ", ".join(os.path.basename(p) for p in unrelated))
    if not files:
        lines.append("(xlsx/xlsm/pdf 파일 없음)")
    lines.append("")
    lines.append("== 인식된 세트 ==")
    lines.append("세트명 | 토큰 | 문제 | 정답 | PDF | 기대값 | 비고")
    for s in sets:
        pdf_col = os.path.basename(s["pdf"]) if s.get("pdf") else "-"
        if s.get("pdf") and pdf_conflicts_with_set(s):
            pdf_col += " ⚠회차불일치"
        elif s.get("pdf_warning"):
            pdf_col += f" (⚠ {s['pdf_warning']})"
        elif s.get("pdf_tie"):
            pdf_col += f" (동점: {s['pdf_tie']})"
        key_col = "-"
        if s.get("key"):
            key_col = os.path.basename(s["key"])
            if s.get("key_origin") and s["key_origin"] != "세트 폴더":
                key_col += f" ({s['key_origin']})"
        elif s.get("key_tie"):
            key_col += f" (동점: {s['key_tie']})"
        note = ""
        if s.get("중복"):
            note = "중복(동일 내용): " + ", ".join(
                os.path.basename(d.get("problem") or d.get("name") or "?")
                for d in s["중복"])
        lines.append(f"{s['name']} | {','.join(sorted(_set_tokens_of(s)))} | "
                     f"{os.path.basename(s['problem'])} | "
                     f"{os.path.basename(s['answer'])} | "
                     f"{pdf_col} | {key_col} | {note or '-'}")
    if not sets:
        lines.append("(인식된 세트 없음 — 같은 폴더에 '…문제.xlsx'와 '…정답.xlsm' "
                     "짝이 있어야 합니다)")
    similar = similar_named_sets(sets)
    if similar:
        lines.append("")
        lines.append("== 이름이 비슷한 세트 (문제 파일 내용은 다름) ==")
        for a, b, common in similar:
            lines.append(f"{a['name']} ↔ {b['name']} (공통 토큰 {', '.join(common)})"
                         " — 같은 세트를 두 번 두었다면 한쪽을 지우세요")
    lines.append("")
    lines.append("== 일정 슬롯 매칭 (전역 유일 배정: 한 세트는 한 슬롯) ==")
    if specs is None:
        specs = all_slot_specs()
    mapping = {k: v for k, v in _cfg_section(config, "_슬롯매핑").items()
               if isinstance(v, dict)}
    assigned = assign_slots(sets, specs, mapping)
    for spec in specs:
        s, how = assigned.get(spec, (None, ""))
        line = f"슬롯 '{spec}' → " + (f"세트 '{s['name']}' ({how})" if s
                                     else f"미발견 ({how})")
        if s:
            line += (" · PDF " + (os.path.basename(s["pdf"]) if s.get("pdf")
                                  else "없음") +
                     " · 기대값 " + (os.path.basename(s["key"]) if s.get("key")
                                   else "없음"))
        m = mapping.get(spec)
        if m and not (s and s.get("direct")):
            line += (f" · 직접 선택 저장됨: {os.path.basename(m.get('problem') or '?')}"
                     " (파일 없음 — 무시)")
        lines.append(line)
    stray = [os.path.basename(p) for p in files
             if os.path.abspath(p) not in owner
             and os.path.splitext(p)[1].lower() != ".pdf"]
    if stray:
        lines.append("")
        lines.append("== 미소속 Excel 파일 (문제/정답 짝을 못 찾음) ==")
        lines.extend(stray)
        lines.append("→ 같은 폴더에 같은 이름으로 '…_문제.xlsx'와 '…_정답.xlsm'"
                     "(확장자 달라도 됨) 짝을 만들거나 [직접 선택]으로 지정하세요.")
    return "\n".join(lines)

def _record_date(r):
    try:
        return datetime.strptime(str(r.get("일시"))[:10], "%Y-%m-%d").date()
    except Exception:
        return None

def set_exam_records(set_name, records):
    return [r for r in set_records(set_name, records)
            if r.get("mode") not in ("부분연습", "오답재풀이")]

def pick_set_for_retry(sets, records, count=1, today=None, exclude=()):
    today = today or date.today()
    taken = set()
    for x in exclude:
        taken.add(x.get("norm") if isinstance(x, dict) else str(x))
    picks = []
    for _ in range(max(0, int(count))):
        cands = [s for s in sets
                 if s.get("norm") not in taken and s.get("name") not in taken]
        if not cands:
            break
        fresh = [s for s in cands
                 if not set_exam_records(s["name"], records)]
        if fresh:
            chosen, reason = fresh[0], "신규 세트라서 (응시 기록 없음)"
        else:
            rows = []
            for s in cands:
                recs = set_exam_records(s["name"], records)
                scores = [r["점수"] for r in recs
                          if isinstance(r.get("점수"), (int, float))]
                best = max(scores) if scores else None
                dates = [d for d in (_record_date(r) for r in recs) if d]
                last = max(dates) if dates else None
                recent = last is not None and (today - last).days < 2
                rows.append((s, best, len(recs), recent))
            pool = [r for r in rows if not r[3]] or rows
            fallback = pool is rows and any(r[3] for r in rows)
            s, best, n, _rec = min(
                pool, key=lambda r: (r[1] if r[1] is not None else -1, r[2],
                                     r[0]["name"]))
            reason = (f"최저점 {best:g}점이라서" if best is not None
                      else "채점 점수가 남아 있지 않아서")
            if fallback:
                reason += " (모든 세트를 최근 2일 내 응시 — 최저점으로 선택)"
            chosen = s
        picks.append((chosen, reason))
        taken.add(chosen.get("norm"))
        taken.add(chosen.get("name"))
    return picks

def load_auto_picks(day_tag, path=None):
    cfg = load_set_config(path or SET_CONFIG_PATH)
    raw = _cfg_section(cfg, "_자동선택").get(day_tag) or []
    if not isinstance(raw, (list, tuple)):
        raw = []
    out = []
    for x in raw:
        if isinstance(x, dict) and x.get("세트"):
            out.append({"세트": str(x["세트"]), "이유": str(x.get("이유") or "")})
        elif isinstance(x, str):
            out.append({"세트": x, "이유": ""})
        else:
            out.append(None)
    return out

def save_auto_picks(day_tag, picks, path=None):
    p = path or SET_CONFIG_PATH
    cfg = load_set_config(p)
    cfg.setdefault("_자동선택", {})[day_tag] = [
        {"세트": x["세트"], "이유": x.get("이유", "")} if x else None
        for x in picks]
    return save_set_config(cfg, p)

def resolve_day_sets(plan, sets, records=None, today=None, saved=None,
                     mapping=None):
    slots = list(plan.get("세트") or [])
    out = [None] * len(slots)
    reasons = [""] * len(slots)
    taken = set()
    mapping = mapping or {}
    objs = list(plan.get("세트객체") or [])
    whys = list(plan.get("슬롯사유") or [])
    assigned = None
    for k, spec in enumerate(slots):
        if spec != AUTO:
            s, how = None, ""
            if k < len(objs) and objs[k]:
                s = objs[k]
                how = "적응형 배정 · " + (whys[k] if k < len(whys) and whys[k]
                                       else "배정")
            if not s:
                if assigned is None:
                    assigned = assign_slots(
                        sets, [x for x in slots if x != AUTO], mapping)
                s, how = assigned.get(spec, (None, ""))
            if s and how.startswith("적응형 배정"):
                reasons[k] = how
            else:
                reasons[k] = (f"일정 지정 세트 · {how}" if s else
                              f"세트를 찾지 못함 · {how}")
            out[k] = s
            if s:
                taken.add(s["norm"])
    saved = list(saved or [])
    for k, spec in enumerate(slots):
        if spec == AUTO and k < len(saved) and saved[k]:
            s = next((x for x in sets if x["name"] == saved[k]["세트"]
                      and x["norm"] not in taken), None)
            if s:
                out[k] = s
                reasons[k] = (saved[k].get("이유") or "이전 선택") + " (유지)"
                taken.add(s["norm"])
    for k, spec in enumerate(slots):
        if spec == AUTO and out[k] is None:
            picks = pick_set_for_retry(
                [s for s in sets if s["norm"] not in taken],
                records or [], 1, today)
            if picks:
                out[k], reasons[k] = picks[0]
                taken.add(out[k]["norm"])
            else:
                reasons[k] = "선택할 세트가 없습니다 (세트를 폴더에 넣어 주세요)"
    return list(zip(out, reasons))

def wrong_sheets_from_items(items):
    out = []
    for it in items or []:
        sh = str((it or {}).get("sheet") or "").strip()
        if sh and sh not in out:
            out.append(sh)
    return out

def make_retry_copy(problem, set_name, when=None):
    when = when or datetime.now()
    stamp = when.strftime("%Y%m%d_%H%M")
    d = os.path.dirname(os.path.abspath(problem))
    return copy_as_macro_enabled(
        problem, _unique_stem(d, f"오답재풀이_{set_name}_{stamp}"))

def retry_payload_for_set(s, minutes=15, json_path=None):
    jp = json_path if json_path and os.path.isfile(json_path) \
        else find_latest_result_json(s, full_only=True)
    if not jp:
        return "missing", {"이유": f"'{s['name']}'의 채점 기록(채점결과 JSON)"
                                 "이 없습니다. 먼저 시험 모드로 응시해 "
                                 "채점을 받으세요."}
    data, items = load_wrong_items(jp)
    if data is None:
        return "missing", {"이유": f"채점결과 파일을 읽을 수 없습니다:\n{jp}"}
    sheets = wrong_sheets_from_items(items)
    if not sheets:
        return "info", {"메시지": f"최근 채점({data.get('total', '?')}점)에 "
                                "오답 시트가 없습니다. 재풀이할 내용이 없어 "
                                "이 단계를 완료 처리합니다.",
                        "자동완료": True}
    return "retry", {"set": s, "sheets": sheets, "label": "오답재풀이",
                     "minutes": int(minutes), "mode": "오답재풀이",
                     "json": jp, "점수": data.get("total")}

def study_segments(exam=None, today=None):
    ex = effective_exams(exam, today)
    if not ex:
        return []
    return [{"seg": 1, "seg_start": ex[0] - timedelta(days=ROUTINE_DAYS),
             "end": ex[0] - timedelta(days=1), "exam": ex[0]}]

def next_segment_after(seg_no, exam=None, today=None):
    return next((sg for sg in study_segments(exam, today)
                 if sg["seg"] == seg_no + 1), None)

def segment_for(today, exam=None):
    all_ex = _exams(exam)
    if not all_ex:
        return {"kind": "unset"}
    if today in all_ex:
        return {"kind": "exam", "exam_no": all_ex.index(today) + 1}
    if today > max(all_ex):
        return {"kind": "after"}
    sg = study_segments(all_ex, today)[0]
    if today < sg["seg_start"]:
        return {"kind": "free", "seg_start": sg["seg_start"], "end": sg["end"]}
    return {"kind": "week", "seg": sg["seg"], "seg_start": sg["seg_start"],
            "start": max(today, sg["seg_start"]), "end": sg["end"]}

def base_capacity(d, seg_end):
    return DAILY_CAP

def exam_records_by_set(sets, records):
    out = {}
    for s in sets:
        recs = [r for r in set_exam_records(s["name"], records)
                if _record_date(r)]
        recs.sort(key=lambda r: (_record_date(r), str(r.get("일시"))))
        out[s["norm"]] = recs
    return out

def prioritized_sets(sets):
    ordered = sorted(sets or [], key=lambda s: (str(s.get("name") or ""),
                                                str(s.get("norm") or "")))
    return [(s, f"{i + 1}/{len(ordered)}") for i, s in enumerate(ordered)]

def _best(recs):
    scores = [r["점수"] for r in recs
              if isinstance(r.get("점수"), (int, float))]
    return max(scores) if scores else None

def _slot(kind, s, goal=None, why="", spec=None):
    return {"kind": kind, "set": s, "name": s["name"] if s else "자동 선택",
            "goal": goal, "auto": s is None, "why": why,
            "spec": spec or (s["name"] if s else AUTO)}

NO_SETS_REASON = "세트 없음 — 문제·정답 파일을 넣어 주세요"

def build_adaptive_plan(today, records, sets, exam=None):
    ex = _exams(exam)
    seg = segment_for(today, ex)
    lv = level_info(records)
    plan = {"today": today, "kind": seg["kind"], "days": {}, "free": {},
            "boost_days": [], "demoted": [], "missed_days": [], "reason": "",
            "remaining": 0, "warning": None, "retry_target": 0,
            "seg": seg.get("seg"), "level": lv,
            "streak": streak_info(records, today, ex), "attempts_left": 0,
            "today_locked": False, "passed_sets": 0,
            "total_sets": len(sets or []), "all_clear": False,
            "no_sets": not sets, "exam_dates": list(ex),
            "retry_queue": []}
    if seg["kind"] not in ("week", "free"):
        plan.update(seg)
        if seg["kind"] == "unset":
            plan["reason"] = "시험일 미설정 — 시험 날짜를 입력하면 일정이 " \
                             "만들어집니다"
        elif seg["kind"] == "after":
            plan["reason"] = "시험이 끝났습니다 — 새 시험을 준비한다면 " \
                             "시험일을 다시 설정하세요"
        elif seg["kind"] == "between":
            plan["reason"] = "2차 시험 대비 — 오답 복습과 실수 노트만 가볍게"
        return plan
    seg_start, end = seg["seg_start"], seg["end"]
    start = seg.get("start", seg_start)
    plan.update({"start": start, "end": end, "seg_start": seg_start})
    by_set = exam_records_by_set(sets, records)
    recs_all = scored_exam_records(records)
    by_date = {}
    for r in recs_all:
        by_date.setdefault(_record_date(r), []).append(r)
    plan["_recs_by_date"] = by_date
    plan["today_locked"] = False

    days = []
    d = start
    while d <= end:
        if d not in ex:
            days.append(d)
        d += timedelta(days=1)
    cap = {dd: DAILY_CAP for dd in days}
    plan["attempts_left"] = len(days)

    goal = lv["current_goal"]
    ordered = prioritized_sets(sets)
    first_pool = [(s, lab) for s, lab in ordered if not by_set.get(s["norm"])]
    attempted = [s for s, _lab in ordered if by_set.get(s["norm"])]
    bests = {s["norm"]: _best(by_set[s["norm"]]) for s in attempted}

    def _b(s):
        b = bests.get(s["norm"])
        return b if b is not None else -1

    def _key(s):
        return (_b(s), s["name"])

    plan["passed_sets"] = len([s for s in attempted if _b(s) >= PASS_LINE])
    below_pass = sorted((s for s in attempted if _b(s) < PASS_LINE), key=_key)
    below_goal = sorted((s for s in attempted
                         if PASS_LINE <= _b(s) < goal), key=_key)
    retry_pool = [("retry", s) for s in below_pass] \
        + [("goal", s) for s in below_goal]
    plan["all_clear"] = bool(sets) and not first_pool and not retry_pool
    plan["retry_queue"] = [{"kind": k, "name": s["name"],
                            "norm": s.get("norm") or "",
                            "best": bests.get(s["norm"])}
                           for k, s in retry_pool]
    round_pool = [] if plan["all_clear"] else [s for s, _lab in ordered]

    queue = [_slot("first", s, goal, f"첫 응시 · 보유 세트 {lab}",
                   spec=s["name"]) for s, lab in first_pool]
    per_day = {dd: [] for dd in days}
    cursor = rcursor = 0
    for dd in days:
        while len(per_day[dd]) < cap[dd]:
            if queue:
                per_day[dd].append(dict(queue.pop(0), date=dd))
                continue
            pick = None
            for _try in range(len(retry_pool)):
                cand = retry_pool[cursor % len(retry_pool)]
                cursor += 1
                if all(x["set"]["norm"] != cand[1]["norm"]
                       for x in per_day[dd]):
                    pick = cand
                    break
            if pick is None and round_pool:
                for _try in range(len(round_pool)):
                    cand = round_pool[rcursor % len(round_pool)]
                    rcursor += 1
                    if all(x["set"]["norm"] != cand["norm"]
                           for x in per_day[dd]):
                        pick = ("goal", cand)
                        break
            if pick is None:
                break
            kind, s = pick
            b = bests.get(s["norm"])
            b_txt = f"{b:g}점" if b is not None else "점수 없음"
            why = (f"재응시 · 최고점 {b_txt} (합격선 {PASS_LINE} 미달)"
                   if kind == "retry"
                   else f"오답 복습 후 재응시 · 목표 {goal}점 · 최고점 {b_txt}")
            per_day[dd].append(dict(_slot(kind, s, goal, why), date=dd))
        plan["free"][dd] = max(0, cap[dd] - len(per_day[dd]))
    plan["days"] = {dd: per_day[dd] for dd in days}
    plan["remaining"] = sum(len(v) for v in per_day.values())

    missed = []
    dd = seg_start
    while dd < min(today, end + timedelta(days=1)):
        if dd not in ex and dd not in by_date:
            missed.append(dd)
        dd += timedelta(days=1)
    plan["missed_days"] = missed
    left = plan["attempts_left"]
    if not sets:
        plan["reason"] = NO_SETS_REASON
    elif seg["kind"] == "free":
        plan["reason"] = (f"자유 연습 · {seg_start.isoformat()}(D-{ROUTINE_DAYS})"
                          f"부터 일주일 일정 시작 · 현재 목표 {goal}점")
    else:
        plan["reason"] = (f"만점 달성 · 남은 학습일 {left}일"
                          if lv["tier"] == "perfect"
                          else f"현재 목표 {goal}점 · 남은 학습일 {left}일")
    return plan

def adaptive_day_plan(adaptive, no, exam=None):
    base = plan_for_day(no, exam=exam)
    if not adaptive or adaptive.get("kind") not in ("week", "free"):
        return base
    d = base.get("날짜")
    if d is None or not (1 <= base["no"] <= ROUTINE_DAYS):
        return base
    today = adaptive["today"]
    lv = adaptive.get("level") or level_info([])
    plan = dict(base)
    plan["적응형"] = True
    plan["특별"] = []
    plan["레벨"] = lv
    plan["남은학습일"] = adaptive.get("attempts_left", 0)
    if d < adaptive["start"]:
        recs = adaptive.get("_recs_by_date", {}).get(d, [])
        plan.update({"종류": "지난", "세트": [], "세트객체": [], "스텝": [],
                     "목표들": []})
        if recs:
            plan["제목"] = "완료"
            plan["세트표시"] = [f"{r.get('세트명', '?')}"
                            + (f"({r['점수']}점)" if r.get("점수") is not None
                               else "") for r in recs]
            plan["할일"] = "응시하고 채점까지 끝낸 날입니다 (클리어)."
        else:
            plan["제목"] = "응시 없음"
            plan["세트표시"] = []
            plan["할일"] = ("이 날은 응시 기록이 없습니다 — 지난 날은 "
                          "되돌리지 않습니다 (경고 아님). 밀린 세트는 점수대 "
                          "보드에서 골라 언제든 다시 풀 수 있어요.")
        return plan
    if d not in adaptive["days"]:
        plan.update({"종류": "안내", "세트": [], "세트객체": [], "스텝": [],
                     "제목": "범위 밖", "세트표시": [],
                     "할일": "기록을 반영해 다시 계산됩니다."})
        return plan
    slots = adaptive["days"][d]
    plan["세트"] = [sl["spec"] if sl["set"] else AUTO for sl in slots]
    plan["세트객체"] = [sl["set"] for sl in slots]
    plan["슬롯사유"] = [sl["why"] for sl in slots]
    plan["목표들"] = [sl["goal"] for sl in slots]
    plan["목표"] = lv["current_goal"]
    plan["종류"] = "모의"
    plan["제목"] = "오늘" if d == today else "예정"
    plan["남은세트"] = adaptive["remaining"]
    if d == adaptive["end"]:
        plan["특별"].append("실수노트")
    if d == today:
        plan["사유"] = adaptive.get("reason") or ""
        plan["경고"] = None
        recs = adaptive.get("_recs_by_date", {}).get(d, [])
        plan["오늘응시"] = len([r for r in recs
                            if isinstance(r.get("점수"), (int, float))])
        low = [r for r in recs
               if isinstance(r.get("점수"), (int, float))
               and r["점수"] < PASS_LINE]
        if low:
            plan["안내"] = (f"오늘 점수가 합격선 {PASS_LINE}점에 못 미쳤어요 — "
                          "리포트에서 틀린 부분만 확인해 두면 됩니다 (체크는 "
                          "필수 아님). 같은 세트를 지금 바로 다시 풀어 점수를 "
                          "갱신해도 됩니다.")
    names = [sl["name"] for sl in slots]
    if not slots:
        if adaptive.get("all_clear"):
            plan["제목"] = "전 세트 목표 달성"
            plan["할일"] = (f"전 세트가 현재 목표 {lv['current_goal']}점을 "
                          "넘었습니다 — 자유 복습. 필수 과제 없음 — 남는 "
                          "시간엔 실수 노트·함수 퀴즈.")
        elif adaptive.get("no_sets"):
            plan["제목"] = "세트 없음"
            plan["할일"] = set_intake_help()
        else:
            plan["제목"] = "자유 복습"
            plan["할일"] = ("필수 과제 없음 — 남는 시간엔 실수 노트·함수 "
                          "퀴즈. 점수를 올리고 싶은 세트는 점수대 보드에서 "
                          "골라 바로 응시하세요.")
        plan["스텝"] = []
        return plan
    plan["할일"] = ("오늘 추천 세트 " + " + ".join(names) + " 40분 완주 → 채점 · "
                  + goal_label(plan["목표"])
                  + " (오답노트·오답 재풀이·함수 퀴즈는 선택)")
    if plan.get("오늘응시"):
        plan["할일"] += (f"\n오늘 이미 {plan['오늘응시']}회 응시했지만 "
                         "횟수 제한은 없습니다 — 예습·복습으로 더 풀어도 "
                         "됩니다.")
    plan["스텝"] = build_day_steps(plan, names)
    return plan

STEP_DONE_MESSAGE = ("오늘 클리어! 완주 + 채점까지 끝났습니다 — 오답노트·오답 "
                     "재풀이·함수 퀴즈는 선택입니다. 채점 결과는 루틴 페이지"
                     "([루틴 열기])에 자동 반영되어 있습니다.")

def sheet_names_of(path):
    try:
        import zipfile
        with zipfile.ZipFile(path) as zf:
            xml = zf.read("xl/workbook.xml").decode("utf-8", "replace")
        return [m.group(1).replace("&amp;", "&").replace("&lt;", "<")
                .replace("&gt;", ">").replace("&quot;", '"')
                .replace("&apos;", "'")
                for m in re.finditer(r'<sheet[^>]*\sname="([^"]*)"', xml)]
    except Exception:
        return []

def drill_sheet_name(path, number):
    names = sheet_names_of(path)
    if not names:
        return None
    num = str(int(number))
    cands = [n for n in names if re.search(rf"(?<!\d){num}(?!\d)", n)]
    if len(cands) >= 1:
        return cands[0]
    if 1 <= int(number) <= len(names):
        return names[int(number) - 1]
    return None

def _progress_key(day_tag, exam=None):
    return f"{progress_ns(exam)}:{day_tag}"

def load_step_progress(day_tag, path=None):
    cfg = load_set_config(path or SET_CONFIG_PATH)
    raw = _cfg_section(cfg, "_진행").get(_progress_key(day_tag)) or []
    if isinstance(raw, dict):
        raw = list(raw.keys())
    elif not isinstance(raw, (list, tuple, set)):
        raw = []
    out = set()
    for x in raw:
        try:
            out.add(int(x))
        except (TypeError, ValueError):
            pass
    return out

def save_step_progress(day_tag, done, path=None):
    p = path or SET_CONFIG_PATH
    cfg = load_set_config(p)
    cfg.setdefault("_진행", {})[_progress_key(day_tag)] = \
        sorted(int(i) for i in done)
    return save_set_config(cfg, p)

def resolve_step_action(step, sets, slot_sets=None):
    kind = step.get("형", "안내")

    def slot_set():
        k = step.get("슬롯")
        if slot_sets is not None and k is not None and k < len(slot_sets):
            return slot_sets[k][0], slot_sets[k][1]
        spec = step.get("세트")
        if spec and spec != AUTO:
            return find_set_for_tokens(sets, spec), "일정 지정 세트"
        return None, ""

    if kind == "드릴":
        s = find_set_for_tokens(sets, "계산 드릴")
        if not s:
            return "missing", {"이유": "계산드릴 세트(계산드릴_문제/정답)를 "
                                     "찾지 못했습니다. 드릴 파일을 스캔 "
                                     "폴더에 넣거나 [직접 선택]으로 "
                                     "지정하세요."}
        name = drill_sheet_name(s["problem"], step.get("번호", 1))
        if not name:
            return "missing", {"이유": f"드릴 워크북에서 {step.get('번호')}번 "
                                     "시트를 찾지 못했습니다."}
        return "practice", {"set": s, "sheets": [name],
                            "label": f"드릴{step.get('번호', 1)}",
                            "minutes": int(step.get("분", 15))}
    if kind == "부분연습":
        s = find_set_for_tokens(sets, step["세트"]) if step.get("세트") \
            else None
        sheets = list(step.get("시트") or
                      sheets_for_areas(step.get("영역") or []))
        label = step.get("라벨") or "+".join(step.get("영역") or []) \
            or "+".join(sheets)
        minutes = int(step.get("분") or
                      practice_minutes(step.get("영역") or []))
        return "practice", {"set": s, "sheets": sheets, "label": label,
                            "minutes": minutes,
                            "세트문구": step.get("세트")}
    if kind == "모의":
        s, reason = slot_set()
        if not s:
            spec = step.get("세트")
            why = ("응시할 세트를 자동으로 고르지 못했습니다. 세트 파일을 "
                   "스캔 폴더에 넣거나 [다른 세트로 바꾸기]로 지정하세요."
                   if spec == AUTO else
                   f"'{spec}' 세트를 찾지 못했습니다. 파일을 스캔 폴더에 "
                   "넣거나 [다른 세트로 바꾸기]/[직접 선택]으로 지정하세요.")
            return "missing", {"이유": why}
        return "exam", {"set": s, "minutes": int(step.get("분", 40)),
                        "목표": step.get("목표"), "이유": reason}
    if kind == "채점":
        return "info", {"메시지": step.get("설명") or "채점이 끝나면 자동으로 "
                                                  "체크됩니다."}
    if kind == "오답노트":
        s, _reason = slot_set()
        return "review", {"set": s}
    if kind == "오답재풀이":
        s, _reason = slot_set()
        if not s:
            return "missing", {"이유": "재풀이할 세트를 찾지 못했습니다. 먼저 "
                                     "같은 슬롯의 시험 모드 단계를 진행하세요."}
        return retry_payload_for_set(s, int(step.get("분", 15)))
    payload = {"메시지": step.get("설명") or step.get("이름") or ""}
    if step.get("웹탭"):
        payload["웹탭"] = str(step["웹탭"])
    return "info", payload

PRACTICE_AREAS = [
    ("기본작업", ["기본작업-1", "기본작업-2", "기본작업-3"], 8),
    ("계산작업", ["계산작업"], 15),
    ("분석작업", ["분석작업-1", "분석작업-2"], 7),
    ("기타작업", ["매크로작업", "차트작업"], 8),
]
PRACTICE_PRESETS = [
    ("기본작업", ["기본작업"]),
    ("계산작업", ["계산작업"]),
    ("분석작업", ["분석작업"]),
    ("기타작업", ["기타작업"]),
]

def sheets_for_areas(area_names):
    out = []
    for name, sheet_list, _min in PRACTICE_AREAS:
        if name in area_names:
            out.extend(sheet_list)
    return out

def practice_minutes(area_names):
    total = sum(m for name, _s, m in PRACTICE_AREAS if name in area_names)
    return max(MIN_MINUTES, min(MAX_MINUTES, total)) if total else \
        DEFAULT_MINUTES

def make_practice_copy(problem, set_name, area_label, when=None):
    when = when or datetime.now()
    stamp = when.strftime("%Y%m%d_%H%M")
    safe_area = re.sub(r"[\\/:*?\"<>|,\s]+", "", str(area_label))[:20]
    d = os.path.dirname(os.path.abspath(problem))
    return copy_as_macro_enabled(
        problem, _unique_stem(d, f"부분연습_{set_name}_{safe_area}_{stamp}"))

def set_records(set_name, records=None):
    records = records if records is not None else load_records()
    return [r for r in records
            if str(r.get("세트명", "")).startswith(str(set_name))]

def problem_has_formula_traces(problem_path, sheet_name="계산작업"):
    try:
        import zipfile
        zf = zipfile.ZipFile(problem_path)
        wb = zf.read("xl/workbook.xml").decode("utf-8", "replace")
        rels = dict(re.findall(
            r'Id="(rId\d+)"[^>]*Target="([^"]+)"',
            zf.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")))
        part = None
        for m in re.finditer(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"',
                             wb):
            if re.sub(r"\s+", "", m.group(1)) == sheet_name:
                t = rels.get(m.group(2), "")
                part = t.lstrip("/") if t.startswith("/") else "xl/" + t
                part = part.replace("xl/xl/", "xl/")
                break
        if not part:
            return 0
        xml = zf.read(part).decode("utf-8", "replace")
        return len(re.findall(r"<f[ >]", xml))
    except Exception:
        return 0

def find_latest_result_json(set_info, full_only=False):
    d = os.path.join(set_info["dir"], "채점결과")
    best = None
    if os.path.isdir(d):
        for fn in os.listdir(d):
            if not (fn.startswith("채점결과_") and fn.endswith(".json")):
                continue
            path = os.path.join(d, fn)
            if set_info["name"] not in fn or full_only:
                try:
                    with open(path, encoding="utf-8") as f:
                        j = json.load(f)
                    if set_info["name"] not in fn:
                        pb = os.path.basename(
                            (j.get("files") or {}).get("problem") or "")
                        if pb != os.path.basename(set_info["problem"]):
                            continue
                    if full_only and j.get("mode") == "partial":
                        continue
                except Exception:
                    continue
            mt = os.path.getmtime(path)
            if best is None or mt > best[0]:
                best = (mt, path)
    return best[1] if best else None

def load_wrong_items(json_path):
    try:
        with open(json_path, encoding="utf-8") as f:
            data = json.load(f)
        items = [i for i in (data.get("wrong_items") or [])
                 if isinstance(i, dict)]
        return data, items
    except Exception:
        return None, []

def find_latest_attempt(set_info):
    best = None
    try:
        for fn in os.listdir(set_info["dir"]):
            if fn.startswith("풀이_") and \
                    fn.lower().endswith((".xlsx", ".xlsm")):
                p = os.path.join(set_info["dir"], fn)
                mt = os.path.getmtime(p)
                if best is None or mt > best[0]:
                    best = (mt, p)
    except OSError:
        pass
    return best[1] if best else None

def make_review_copy(source, set_name, when=None):
    when = when or datetime.now()
    stamp = when.strftime("%Y%m%d_%H%M")
    d = os.path.dirname(os.path.abspath(source))
    return copy_as_macro_enabled(
        source, _unique_stem(d, f"오답연습_{set_name}_{stamp}"))

def review_item_text(item):
    lines = [f"[{item.get('sheet', '?')}] {item.get('label', '?')}"
             + (f"  (-{item.get('lost', 0):g}점)" if item.get("lost") else "")]
    for c in item.get("cells") or []:
        if c.get("coord"):
            lines.append(f"\n위치: {c['coord']}")
        got = c.get("got")
        lines.append("  내 답: "
                     + (str(got) if got is not None else "(비어 있음)"))
        if c.get("got_formula") and str(c["got_formula"]) != str(got):
            lines.append(f"          {c['got_formula']}")
        exp = c.get("expected")
        lines.append("  정답  : "
                     + (str(exp) if exp is not None else "(비어 있음)"))
        if c.get("formula"):
            lines.append(f"          {c['formula']}")
    if item.get("diff_notes"):
        lines.append("\n[무엇이 다른가]")
        lines.extend(f"  - {n}" for n in item["diff_notes"])
    if item.get("explain"):
        lines.append("\n[정확한 풀이]")
        lines.extend(f"  {s}" for s in item["explain"])
    if item.get("point"):
        lines.append(f"\n포인트: {item['point']}")
    if item.get("note"):
        lines.append(f"\n참고: {item['note']}")
    if item.get("hint"):
        lines.append(f"\n방법: {item['hint']}")
    return "\n".join(lines)

def load_review_state(norm_key, json_name, path=SET_CONFIG_PATH):
    cfg = load_set_config(path)
    rv = (cfg.get(norm_key) or {}).get("오답연습") or {}
    try:
        return {int(i) for i in (rv.get("이해체크") or {}).get(json_name, [])}
    except Exception:
        return set()

def record_review_state(norm_key, json_name, checks,
                        path=SET_CONFIG_PATH, count_up=False, total=None,
                        done=None):
    cfg = load_set_config(path)
    ent = cfg.setdefault(norm_key, {})
    if not isinstance(ent.get("오답연습"), dict):
        ent["오답연습"] = {}
    rv = ent["오답연습"]
    if count_up:
        try:
            rv["횟수"] = int(rv.get("횟수") or 0) + 1
        except (TypeError, ValueError):
            rv["횟수"] = 1
    if not isinstance(rv.get("이해체크"), dict):
        rv["이해체크"] = {}
    rv["이해체크"][str(json_name)] = sorted(int(i) for i in checks)
    if not isinstance(rv.get("완료"), dict):
        rv["완료"] = {}
    if done is None and total is not None:
        done = int(total) > 0 and len(set(int(i) for i in checks)) >= int(total)
    if done is True:
        rv["완료"].setdefault(str(json_name),
                            datetime.now().isoformat(timespec="minutes"))
    elif done is False:
        rv["완료"].pop(str(json_name), None)
    return save_set_config(cfg, path)

def review_done_state(norm_key, json_name, total=None, cfg=None,
                      path=SET_CONFIG_PATH):
    if cfg is None:
        cfg = load_set_config(path)
    rv = (cfg.get(norm_key) or {}).get("오답연습") or {}
    when = (rv.get("완료") or {}).get(str(json_name)) if isinstance(
        rv.get("완료"), dict) else None
    if when:
        return True, str(when)
    try:
        checks = {int(i) for i in (rv.get("이해체크") or {}).get(str(json_name), [])}
    except Exception:
        checks = set()
    if total and len(checks & set(range(int(total)))) >= int(total):
        return True, None
    return False, None

UPDATE_BASE_URL = ("https://raw.githubusercontent.com/"
                   "xasmine02/cocomate-study-tools/main/")
SET_FILES_KEY = "set_files"
_BINARY_MAGIC = {".xlsx": b"PK\x03\x04", ".xlsm": b"PK\x03\x04", ".pdf": b"%PDF"}
UPDATE_NOTICE_PATH = os.path.join(RECORDS_DIR, ".update_applied.json")
AUTO_UPDATE_SETTING = "자동업데이트"
_VERSION_MARK_RE = re.compile(r'^__version__\s*=\s*["\']([^"\']+)["\']',
                              re.MULTILINE)

def _version_tuple(v):
    nums = re.findall(r"\d+", str(v or ""))
    return tuple(int(x) for x in nums[:3]) if nums else (0,)

def auto_update_enabled(path=None):
    return bool(get_app_setting(AUTO_UPDATE_SETTING, True, path=path))

FIRST_RUN_NOTICE_SETTING = "자동업데이트고지"

def _repo_url(base_url=None):
    m = re.match(r"https://raw\.githubusercontent\.com/([^/]+)/([^/]+)/",
                 str(base_url or UPDATE_BASE_URL))
    return (f"https://github.com/{m.group(1)}/{m.group(2)}" if m
            else str(base_url or UPDATE_BASE_URL))

UPDATE_REPO_URL = _repo_url()

def first_run_notice_text(repo_url=UPDATE_REPO_URL):
    return (
        "이 프로그램은 실행할 때마다 아래 저장소에서 최신 코드와 자료를 "
        "내려받아 적용할 수 있습니다. 자동 업데이트를 쓸지 지금 골라 "
        "주세요.\n\n"
        f"저장소: {repo_url}\n\n"
        "· [동의] — 실행할 때마다 새 버전을 자동으로 받아 적용합니다. 받은 "
        "파일은 sha256 검증과 문법 검사를 거쳐 백업 후 교체됩니다.\n"
        "· [거절] — 자동 업데이트를 끕니다. 프로그램은 그대로 쓸 수 있고, "
        "필요할 때 [업데이트 확인]으로 직접 받으면 됩니다.\n"
        "· 이 창을 그냥 닫으면 **거절**로 봅니다.\n"
        "· 인터넷이 없거나 서버가 응답하지 않으면 그냥 지금 버전으로 "
        "실행됩니다.\n\n"
        f"고른 값은 세트설정.json 의 \"_설정\" → \"{AUTO_UPDATE_SETTING}\" 에 "
        "저장되고, 시작 화면 [설정]에서 언제든 바꿀 수 있습니다.")

def first_run_notice_pending(path=None):
    return not bool(get_app_setting(FIRST_RUN_NOTICE_SETTING, False, path=path))

def mark_first_run_notice(path=None):
    return set_app_setting(FIRST_RUN_NOTICE_SETTING, True, path=path)

def apply_update_consent(agree, path=None):
    agree = bool(agree)
    set_app_setting(AUTO_UPDATE_SETTING, agree, path=path)
    mark_first_run_notice(path=path)
    return agree

UPDATE_CONSENT_YES = "동의 — 자동으로 최신 버전 받기"
UPDATE_CONSENT_NO = "거절 — 자동 업데이트 끄기"

def _http_get(url, timeout):
    import urllib.request
    req = urllib.request.Request(url, headers={"User-Agent": "cocomate-updater"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()

def fetch_update_info(base_url=UPDATE_BASE_URL, timeout=3):
    try:
        info = json.loads(_http_get(base_url + "version.json", timeout)
                          .decode("utf-8"))
        if isinstance(info, dict) and info.get("version"):
            for k in ("files", "data_files", SET_FILES_KEY, "sha256"):
                if not isinstance(info.get(k), dict):
                    info[k] = {}
            return info
    except Exception:
        pass
    return None

def update_available(info, current=None):
    return bool(info) and _version_tuple(info.get("version")) > \
        _version_tuple(current if current is not None else __version__)

def sha256_hex(data):
    return hashlib.sha256(data).hexdigest()

def file_sha256(path):
    try:
        with open(path, "rb") as f:
            return sha256_hex(f.read())
    except OSError:
        return None

def version_marker(text):
    m = _VERSION_MARK_RE.search(text or "")
    return m.group(1) if m else None

def _safe_rel_parts(rel_key):
    raw = str(rel_key).replace(chr(92), "/")
    parts = []
    for part in raw.split("/"):
        part = part.strip()
        if part in ("", ".", ".."):
            continue
        if ":" in part or chr(0) in part:
            raise RuntimeError(f"잘못된 경로 키: {rel_key!r}")
        parts.append(part)
    if not parts:
        raise RuntimeError(f"잘못된 경로 키: {rel_key!r}")
    return parts

def _within(root, path):
    r = os.path.abspath(root)
    a = os.path.abspath(path)
    if a != r and not a.startswith(r + os.sep):
        raise RuntimeError(f"설치 폴더 밖으로 나가는 경로: {path!r}")
    return a

def _update_target_path(rel_key, base_dir=None):
    base_dir = base_dir or BASE_DIR
    root = os.path.dirname(base_dir)
    parts = _safe_rel_parts(rel_key)
    cand = _within(root, os.path.join(root, *parts))
    if os.path.isfile(cand):
        return cand
    flat = _within(root, os.path.join(base_dir, parts[-1]))
    if os.path.isfile(flat):
        return flat
    return cand if os.path.isdir(os.path.dirname(cand)) else flat

def _data_target_path(rel_key, base_dir=None):
    parts = _safe_rel_parts(rel_key)
    if len(parts) == 1:
        parts = [EXPECTED_DIR_NAME] + parts
    root = os.path.dirname(expected_values_dir(base_dir))
    return _within(root, os.path.join(root, *parts))

def _verify_download(repo_rel, data, kind, info, expect_version=None):
    expected = (info.get("sha256") or {}).get(repo_rel)
    if expected:
        got = sha256_hex(data)
        if got.lower() != str(expected).lower():
            raise RuntimeError(f"{repo_rel}: sha256 불일치 "
                               f"(기대 {str(expected)[:12]}…, 실제 {got[:12]}…)")
    ext = os.path.splitext(str(repo_rel))[1].lower()
    if kind != "code" and ext not in (".json", ".html", ".htm"):
        if not expected:
            raise RuntimeError(f"{repo_rel}: sha256 항목이 없어 바이너리 파일을 "
                               "검증할 수 없습니다")
        magic = _BINARY_MAGIC.get(ext)
        if magic and not data.startswith(magic):
            raise RuntimeError(f"{repo_rel}: {ext} 파일 형식이 아닙니다")
        return None
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError as e:
        raise RuntimeError(f"{repo_rel}: UTF-8 텍스트가 아닙니다 ({e})")
    if kind != "code" and ext in (".html", ".htm"):
        if not expected:
            raise RuntimeError(f"{repo_rel}: sha256 항목이 없어 페이지 파일을 "
                               "검증할 수 없습니다")
        return text
    if kind == "code":
        ver = version_marker(text)
        if ver is None:
            raise RuntimeError(f"{repo_rel}: __version__ 표식이 없어 "
                               "배포 파일이 아닌 것으로 판단")
        if expect_version and _version_tuple(ver) != _version_tuple(expect_version):
            raise RuntimeError(f"{repo_rel}: 파일 버전 {ver}이(가) version.json의 "
                               f"{expect_version}와 다릅니다 (재시작 반복 방지)")
    else:
        try:
            json.loads(text)
        except ValueError as e:
            raise RuntimeError(f"{repo_rel}: JSON 형식 오류 ({e})")
    return text

def _cleanup_staged(staged):
    for it in staged:
        for leftover in (it["tmp"], it["tmp"] + "c"):
            try:
                if os.path.isfile(leftover):
                    os.remove(leftover)
            except OSError:
                pass

def _stage_items(items, info, base_url, timeout):
    import py_compile
    import urllib.parse
    staged = []
    try:
        for rel_key, repo_rel, target, kind in items:
            data = _http_get(base_url + urllib.parse.quote(str(repo_rel)),
                             timeout)
            expect = None
            if kind == "code" and os.path.basename(target) == "시험장.py":
                expect = info.get("version")
            _verify_download(repo_rel, data, kind, info, expect_version=expect)
            tmp = target + ".new"
            os.makedirs(os.path.dirname(target) or ".", exist_ok=True)
            with open(tmp, "wb") as f:
                f.write(data)
            staged.append({"target": target, "tmp": tmp, "kind": kind,
                           "rel": repo_rel})
            if kind == "code":
                py_compile.compile(tmp, cfile=tmp + "c", doraise=True)
                try:
                    os.remove(tmp + "c")
                except OSError:
                    pass
        return staged
    except Exception:
        _cleanup_staged(staged)
        raise

def stage_update(info, base_url=UPDATE_BASE_URL, base_dir=None, timeout=15):
    info = info or {}
    files = info.get("files") or {}
    data_files = _all_data_files(info)
    if not files and not data_files:
        return False, "업데이트 파일 목록이 비어 있습니다.", []
    items = [(k, v, _update_target_path(k, base_dir), "code")
             for k, v in files.items()]
    try:
        items += [(k, v, _data_target_path(k, base_dir), "data")
                  for k, v in data_files.items()]
        staged = _stage_items(items, info, base_url, timeout)
    except Exception as e:
        return False, f"업데이트 실패(기존 버전 유지): {e}", []
    return True, f"{len(staged)}개 파일 검증 완료 (v{info.get('version')})", staged

def commit_staged(staged, version=None):
    replaced = []
    try:
        for it in staged:
            target, tmp = it["target"], it["tmp"]
            existed = os.path.isfile(target)
            if existed:
                shutil.copy2(target, target + ".bak")
            os.replace(tmp, target)
            replaced.append((target, existed))
        n_code = sum(1 for it in staged if it["kind"] == "code")
        n_web = sum(1 for it in staged if it["kind"] != "code"
                    and str(it["rel"]).lower().endswith((".html", ".htm")))
        n_set = sum(1 for it in staged if it["kind"] != "code"
                    and not str(it["rel"]).lower().endswith(
                        (".json", ".html", ".htm")))
        n_data = len(staged) - n_code - n_set - n_web
        what = " + ".join(x for x in (
            f"프로그램 {n_code}개" if n_code else "",
            f"기대값 {n_data}개" if n_data else "",
            f"세트 파일 {n_set}개" if n_set else "",
            f"루틴 페이지 {n_web}개" if n_web else "") if x)
        return True, (f"{what} 파일을 v{version}(으)로 업데이트했습니다."
                      if version else f"{what} 파일을 갱신했습니다.")
    except Exception as e:
        _cleanup_staged(staged)
        for target, existed in replaced:
            try:
                if existed and os.path.isfile(target + ".bak"):
                    shutil.copy2(target + ".bak", target)
                elif not existed and os.path.isfile(target):
                    os.remove(target)
            except OSError:
                pass
        return False, f"업데이트 실패(기존 버전 유지): {e}"

def apply_update(info, base_url=UPDATE_BASE_URL, base_dir=None, timeout=15):
    ok, msg, staged = stage_update(info, base_url, base_dir, timeout)
    if not ok:
        return False, msg
    return commit_staged(staged, (info or {}).get("version"))

def _all_data_files(info):
    out = {}
    for k in ("data_files", SET_FILES_KEY):
        v = (info or {}).get(k)
        if isinstance(v, dict):
            out.update(v)
    return out

def data_files_to_sync(info, base_dir=None):
    out = []
    sha = (info or {}).get("sha256") or {}
    for rel_key, repo_rel in _all_data_files(info).items():
        try:
            target = _data_target_path(rel_key, base_dir)
        except RuntimeError:
            continue
        expected = sha.get(repo_rel)
        if not os.path.isfile(target):
            out.append((rel_key, repo_rel, target))
        elif expected and (file_sha256(target) or "").lower() != \
                str(expected).lower():
            out.append((rel_key, repo_rel, target))
    return out

def sync_data_files(info, base_url=UPDATE_BASE_URL, base_dir=None, timeout=15):
    todo = data_files_to_sync(info, base_dir)
    if not todo:
        return True, 0, "기대값 파일이 모두 최신입니다."
    try:
        staged = _stage_items([(k, v, t, "data") for k, v, t in todo],
                              info or {}, base_url, timeout)
    except Exception as e:
        return False, 0, f"기대값·세트 파일 갱신 실패(기존 파일 유지): {e}"
    ok, msg = commit_staged(staged)
    return ok, (len(staged) if ok else 0), msg

def write_update_notice(prev, new, notes="", path=None):
    p = path or UPDATE_NOTICE_PATH
    try:
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, "w", encoding="utf-8") as f:
            json.dump({"from": prev, "to": new, "notes": notes or "",
                       "at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                       "shown": False}, f, ensure_ascii=False, indent=2)
        return True
    except OSError:
        return False

def read_update_notice(path=None):
    try:
        with open(path or UPDATE_NOTICE_PATH, encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else None
    except Exception:
        return None

def consume_update_notice(current=None, path=None):
    p = path or UPDATE_NOTICE_PATH
    d = read_update_notice(p)
    if not d or d.get("shown"):
        return None
    cur = current if current is not None else __version__
    if _version_tuple(d.get("to")) != _version_tuple(cur):
        return None
    d["shown"] = True
    try:
        with open(p, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except OSError:
        pass
    return d

class UpdateCoordinator:

    def __init__(self, base_url=UPDATE_BASE_URL, base_dir=None, enabled=None,
                 current=None, log=None, notice_path=None):
        self.base_url = base_url
        self.base_dir = base_dir or BASE_DIR
        self._enabled = enabled if enabled is not None else auto_update_enabled
        self.current = current or __version__
        self.log = log or startup_log
        self.notice_path = notice_path or UPDATE_NOTICE_PATH
        self.checked = False
        self.checking = False
        self.info = None
        self.pending = None
        self.state = None
        self.detail = ""
        self.data_synced = 0
        self.applied_version = None

    def enabled(self):
        try:
            return bool(self._enabled())
        except Exception:
            return True

    def check(self, force=False, timeout=3, dl_timeout=15):
        if not force:
            if not self.enabled():
                self.log("자동 업데이트 꺼짐(세트설정 _설정.자동업데이트=false) — "
                         "확인 건너뜀")
                return self._done("disabled", "자동 업데이트가 꺼져 있습니다.")
            if self.checked:
                return self._done("already", "이번 실행에서 이미 확인했습니다.")
        self.checked = True
        self.checking = True
        self.data_synced = 0
        try:
            info = fetch_update_info(self.base_url, timeout=timeout)
            if info is None:
                self.log("업데이트 확인: 서버 응답 없음(오프라인/차단) — 건너뜀")
                return self._done("offline", "업데이트 서버에 연결할 수 없습니다.")
            self.info = info
            remote = str(info.get("version"))
            if update_available(info, self.current):
                notice = read_update_notice(self.notice_path)
                if notice and not force and _version_tuple(notice.get("to")) == \
                        _version_tuple(remote):
                    self.log(f"v{remote}을(를) 이미 적용했는데 실행 버전이 "
                             f"v{self.current}에 머물러 자동 적용을 중단합니다 "
                             "([업데이트 확인]으로 수동 재시도 가능)")
                    return self._done("stuck", f"v{remote} 적용 후에도 버전이 "
                                      "오르지 않았습니다.")
                n_files = len(info.get("files") or {})
                n_data = len(info.get("data_files") or {})
                n_set = len(info.get(SET_FILES_KEY) or {})
                self.log(f"새 버전 v{remote} 발견 (현재 v{self.current}) — "
                         f"내려받는 중: 프로그램 {n_files}개 + 기대값 {n_data}개"
                         + (f" + 세트 파일 {n_set}개" if n_set else ""))
                ok, msg, staged = stage_update(info, self.base_url,
                                               self.base_dir, timeout=dl_timeout)
                if not ok:
                    self.log(msg)
                    return self._done("failed", msg)
                self.pending = staged
                self.log(f"검증 통과: {msg} — 적용 대기")
                return self._done("staged", msg)
            ok, n, msg = sync_data_files(info, self.base_url, self.base_dir,
                                         timeout=dl_timeout)
            if not ok:
                self.log(msg)
                return self._done("failed", msg)
            if n:
                self.data_synced = n
                self.log(f"기대값 자동 갱신: {msg}")
                return self._done("data", msg)
            self.log(f"업데이트 확인: 최신 버전 (v{self.current}) · {msg}")
            return self._done("latest", f"현재 최신 버전입니다 (v{self.current}).")
        except Exception as e:
            self.log(f"업데이트 확인 중 예외: {e}")
            return self._done("failed", f"업데이트 확인 중 오류: {e}")
        finally:
            self.checking = False

    def _done(self, state, detail):
        self.state, self.detail = state, detail
        return state

    def apply_pending(self, exam_running=False):
        if not self.pending:
            return "none"
        if exam_running:
            self.log("시험 진행 중 — 업데이트 적용을 시험 종료 후로 미룹니다")
            return "deferred"
        staged, self.pending = self.pending, None
        new_ver = (self.info or {}).get("version")
        ok, msg = commit_staged(staged, new_ver)
        self.detail = msg
        if not ok:
            self.log(msg)
            return "failed"
        self.applied_version = new_ver
        write_update_notice(self.current, new_ver,
                            (self.info or {}).get("notes") or "",
                            path=self.notice_path)
        self.log(f"업데이트 적용: {msg}")
        return "applied"

CORRECTIONS_DIR_NAME = "정정"
CORRECTIONS_FILE_NAME = "점수정정.json"
CORRECTIONS_DATA_KEY = "정정/점수정정.json"
CORRECTION_HISTORY_FIELD = "정정이력"
CORRECTION_KEY_FIELD = "키"
CORRECTION_BONUS_FIELD = "정정가산키"
REGRADE_EVIDENCE_FIELDS = ("채점기버전", "문제해시", "정답해시", "기대값해시",
                           "풀이경로")
REGRADE_SUFFIX = "_재채점"
REGRADE_SKIP_MODES = ("부분연습", "오답재풀이", "수동")
REGRADE_MAX = 60
_GRADER_VERSION_CACHE = {}
_SHA_CACHE = {}

def corrections_path(base_dir=None):
    root = os.path.dirname(expected_values_dir(base_dir))
    return os.path.join(root, CORRECTIONS_DIR_NAME, CORRECTIONS_FILE_NAME)

def file_sha256_cached(path):
    if not path:
        return None
    try:
        st = os.stat(path)
    except OSError:
        return None
    key = (st.st_mtime, st.st_size)
    hit = _SHA_CACHE.get(path)
    if hit and hit[0] == key:
        return hit[1]
    h = file_sha256(path)
    _SHA_CACHE[path] = (key, h)
    return h

def grader_version(grade_py=None):
    p = grade_py or find_grade_py()
    if not p or not os.path.isfile(p):
        return None
    try:
        st = os.stat(p)
        key = (st.st_mtime, st.st_size)
    except OSError:
        return None
    hit = _GRADER_VERSION_CACHE.get(p)
    if hit and hit[0] == key:
        return hit[1]
    try:
        with open(p, encoding="utf-8", errors="replace") as f:
            ver = version_marker(f.read())
    except OSError:
        return None
    _GRADER_VERSION_CACHE[p] = (key, ver)
    return ver

def grading_evidence(problem=None, answer=None, student=None, key=None,
                     grade_py=None):
    out = {}
    ver = grader_version(grade_py)
    if ver:
        out["채점기버전"] = ver
    for field, path in (("문제해시", problem), ("정답해시", answer),
                        ("기대값해시", key)):
        h = file_sha256_cached(path) if path and os.path.isfile(path) else None
        if h:
            out[field] = h
    if student:
        out["풀이경로"] = os.path.abspath(student)
    return out

def record_evidence_changes(rec, s, grade_py=None, student=None):
    s = s or {}
    fresh = grading_evidence(problem=s.get("problem"), answer=s.get("answer"),
                             key=s.get("key"), student=student,
                             grade_py=grade_py)
    return {k: v for k, v in fresh.items() if (rec or {}).get(k) != v}

def regrade_reason(rec, s, cur_version=None):
    if not isinstance(rec, dict) or not isinstance(s, dict):
        return ""
    if str(rec.get("mode") or "시험") in REGRADE_SKIP_MODES:
        return ""
    score = rec.get("점수")
    if isinstance(score, bool) or not isinstance(score, (int, float)):
        return ""
    reasons = []
    old = rec.get("채점기버전")
    if cur_version:
        try:
            older = (not old) or _version_tuple(str(old)) \
                < _version_tuple(str(cur_version))
        except Exception:
            older = True
        if older:
            reasons.append(f"채점기 {cur_version}"
                           + (f" (기록 {old})" if old else " (버전 기록 없음)"))
    for field, label, path in (("정답해시", "정답 파일", s.get("answer")),
                               ("문제해시", "문제 파일", s.get("problem")),
                               ("기대값해시", "기대값 파일", s.get("key"))):
        want = rec.get(field)
        now = file_sha256_cached(path) if path and os.path.isfile(path) \
            else None
        if want and now and str(want).lower() != str(now).lower():
            reasons.append(f"{label} 정정")
    return " · ".join(reasons)

def attempt_file_for(rec, s=None):
    rec = rec or {}
    p = rec.get("풀이경로")
    if isinstance(p, str) and p and os.path.isfile(p):
        return p
    m = _RECORD_STAMP_RE.search(str(rec.get("일시") or ""))
    if not m:
        return None
    y, mo, d, h, mi, _sec = m.groups()
    stamp = f"{y}{mo}{d}_{h}{mi}"
    dirs = []
    if s and s.get("dir"):
        dirs.append(s["dir"])
    rep = rec.get("리포트")
    if isinstance(rep, str) and rep:
        dirs.append(os.path.dirname(os.path.dirname(os.path.abspath(rep))))
    seen = set()
    for d0 in dirs:
        if not d0 or d0 in seen or not os.path.isdir(d0):
            continue
        seen.add(d0)
        try:
            names = sorted(os.listdir(d0))
        except OSError:
            continue
        for fn in names:
            if fn.startswith("풀이_") and stamp in fn \
                    and fn.lower().endswith((".xlsx", ".xlsm")):
                return os.path.join(d0, fn)
    return None

def regrade_report_base(rec, s=None, when=None):
    when = when or datetime.now()
    stamp = when.strftime("%Y%m%d_%H%M%S")
    name = record_set_name(rec)
    out_dir = None
    rep = (rec or {}).get("리포트")
    if isinstance(rep, str) and rep:
        cand = os.path.dirname(os.path.abspath(rep))
        if os.path.isdir(cand):
            out_dir = cand
    if out_dir is None:
        student = attempt_file_for(rec, s)
        base = os.path.dirname(os.path.abspath(student)) if student \
            else ((s or {}).get("dir") or RECORDS_DIR)
        out_dir = os.path.join(base, "채점결과")
    try:
        os.makedirs(out_dir, exist_ok=True)
    except OSError:
        out_dir = RECORDS_DIR
        _ensure_records_home()
    return os.path.join(out_dir, f"채점결과_{name}_{stamp}{REGRADE_SUFFIX}")

def regrade_record(rec, s, grade_py=None, student=None, when=None,
                   timeout=600):
    grade_py = grade_py or find_grade_py()
    if not grade_py or not os.path.isfile(grade_py):
        return None, None, None, "grade.py 를 찾을 수 없습니다"
    student = student or attempt_file_for(rec, s)
    if not student or not os.path.isfile(student):
        return None, None, None, "풀이 파일이 없습니다"
    if not s or not s.get("problem") or not s.get("answer"):
        return None, None, None, "문제·정답 파일을 찾을 수 없습니다"
    base = regrade_report_base(rec, s, when)
    html_path, json_path = base + ".html", base + ".json"
    result, out, err, rc = run_grading(
        grade_py, s["problem"], s["answer"], student, key=s.get("key"),
        html=html_path, json_out=json_path, timeout=timeout)
    if result is None:
        detail = ((err or "") + "\n" + (out or "")).strip()
        return None, None, None, (detail[-300:] or f"채점 실패(rc={rc})")
    return result, html_path, json_path, ""

def correction_entry(before, after, reason, when=None, source="재채점",
                     key=""):
    return {"이전총점": before, "새총점": after, "사유": str(reason or ""),
            "일시": (when or datetime.now()).strftime("%Y-%m-%d %H:%M:%S"),
            "출처": source, CORRECTION_KEY_FIELD: str(key or "")}

def correction_applied(rec, key):
    if not key:
        return False
    hist = (rec or {}).get(CORRECTION_HISTORY_FIELD)
    if not isinstance(hist, list):
        return False
    return any(isinstance(h, dict)
               and str(h.get(CORRECTION_KEY_FIELD) or "") == str(key)
               for h in hist)

def apply_correction(rec, new_total, reason, when=None, source="재채점",
                     key="", report=None, result=None, evidence=None):
    if not isinstance(rec, dict):
        return None
    if key and correction_applied(rec, key):
        return None
    before = _coerce_score(rec.get("점수"))
    after = _coerce_score(new_total)
    if after is None:
        return None
    if report:
        rec["리포트"] = report
    if isinstance(result, dict):
        mx = result.get("max_total")
        if isinstance(mx, (int, float)) and not isinstance(mx, bool) \
                and mx != 100:
            rec["만점"] = _coerce_score(mx)
        gs = result.get("graded_sheets")
        if isinstance(gs, list) and gs:
            rec["채점시트"] = [str(x) for x in gs]
    if isinstance(evidence, dict):
        rec.update(evidence)
    if before == after:
        return None
    rec["점수"] = after
    hist = rec.get(CORRECTION_HISTORY_FIELD)
    hist = list(hist) if isinstance(hist, list) else []
    entry = correction_entry(before, after, reason, when=when, source=source,
                             key=key)
    hist.append(entry)
    rec[CORRECTION_HISTORY_FIELD] = hist
    return entry

def load_corrections(path=None):
    p = path or corrections_path()
    try:
        with open(p, encoding="utf-8-sig") as f:
            data = json.load(f)
    except Exception:
        return []
    items = data.get("corrections") if isinstance(data, dict) else data
    if not isinstance(items, list):
        return []
    out = []
    for it in items:
        if not isinstance(it, dict):
            continue
        rid = it.get("report_id")
        total = _coerce_score(it.get("new_total"))
        delta = _coerce_score(it.get("delta"))
        if not isinstance(rid, str) or not rid.strip():
            continue
        if total is None and not delta:
            continue
        out.append({"report_id": rid.strip(), "set": str(it.get("set") or ""),
                    "new_total": total, "delta": delta or 0,
                    "reason": str(it.get("reason") or "이의제기 인정")})
    return out

def correction_matches(it, rec):
    if not it.get("set"):
        return True
    name = record_set_name(rec)
    return it["set"] in (name, norm_set_key(name))

def correction_deltas(items, rid, rec):
    return [(it["delta"], it["reason"]) for it in items
            if it["report_id"] == rid and it.get("delta")
            and correction_matches(it, rec)]

def resolve_correction_target(by_id, wanted):
    rec = by_id.get(wanted)
    if rec is not None:
        return wanted, rec
    if not (isinstance(wanted, str) and len(wanted) == 14
            and wanted.isdigit()):
        return wanted, None
    same = [(rid, r) for rid, r in by_id.items()
            if len(rid) == 14 and rid[:12] == wanted[:12]]
    return same[0] if len(same) == 1 else (wanted, None)

def correction_delta_key(rid, delta):
    return f"data:{rid}:{delta:+g}"

def correction_bonus_applied(rec, key):
    keys = rec.get(CORRECTION_BONUS_FIELD)
    return isinstance(keys, list) and key in keys

def correction_change_text(ch):
    before = ch.get("before")
    b = f"{before:g}" if isinstance(before, (int, float)) else "—"
    return f"{ch.get('name') or '?'} {b} → {ch.get('after'):g}점"

def correction_summary_text(changes, limit=3):
    changes = [c for c in (changes or []) if isinstance(c, dict)]
    if not changes:
        return ""
    head = "이의제기 반영: "
    if len(changes) == 1:
        return head + correction_change_text(changes[0])
    shown = [correction_change_text(c) for c in changes[:limit]]
    more = len(changes) - len(shown)
    return (f"{head}{len(changes)}건 — " + " / ".join(shown)
            + (f" 외 {more}건" if more > 0 else "")
            + " (클릭하면 자세히)")

class RegradeCoordinator:

    def __init__(self, sets=None, grade_py=None, records_path=None,
                 corrections_file=None, progress=None, log=None, when=None,
                 timeout=600, limit=REGRADE_MAX):
        self.sets = list(sets or [])
        self.grade_py = grade_py or find_grade_py()
        self.records_path = records_path or RECORDS_PATH
        self.corrections_file = corrections_file
        self._progress = progress
        self._log = log or startup_log
        self.when = when
        self.timeout = timeout
        self.limit = limit

    def log(self, msg):
        try:
            self._log(msg)
        except Exception:
            pass

    def progress(self, i, n, name):
        if self._progress is None:
            return
        try:
            self._progress(i, n, name)
        except Exception:
            pass

    def _set_for(self, rec):
        name = record_set_name(rec)
        for s in self.sets:
            if s.get("name") == name:
                return s
        norm = norm_set_key(name) if name and name != "?" else ""
        if norm:
            for s in self.sets:
                if s.get("norm") == norm:
                    return s
        return None

    def _commit(self, patches):
        if not patches:
            return False
        with _FILE_LOCK:
            records = load_records(self.records_path)
            hit = 0
            for rid, r in records_with_ids(records):
                patch = patches.get(rid)
                if patch:
                    r.update(patch)
                    hit += 1
            if not hit:
                return False
            _backup_corrupt_json(self.records_path)
            tmp = self.records_path + ".tmp"
            try:
                with open(tmp, "w", encoding="utf-8") as f:
                    json.dump(records, f, ensure_ascii=False, indent=2)
                os.replace(tmp, self.records_path)
            except OSError as e:
                log_error("자동 재채점 기록 저장", e)
                try:
                    if os.path.isfile(tmp):
                        os.remove(tmp)
                except OSError:
                    pass
                return False
        routine_touch()
        return True

    def plan(self, records=None):
        records = load_records(self.records_path) if records is None \
            else records
        cur = grader_version(self.grade_py)
        out = []
        for rid, rec in records_with_ids(records):
            s = self._set_for(rec)
            if s is None:
                continue
            why = regrade_reason(rec, s, cur)
            if not why:
                continue
            out.append((rid, rec, s, why, attempt_file_for(rec, s)))
        return out

    def run(self):
        summary = {"checked": 0, "targets": 0, "regraded": 0, "changes": [],
                   "errors": [], "data_applied": 0, "skipped": 0}
        try:
            records = load_records(self.records_path)
        except Exception as e:
            log_error("자동 재채점 기록 읽기", e)
            return summary
        summary["checked"] = len(records)
        try:
            targets = self.plan(records)
        except Exception as e:
            log_error("자동 재채점 대상 선별", e)
            targets = []
        summary["targets"] = len(targets)
        try:
            data_items = load_corrections(self.corrections_file)
        except Exception as e:
            log_error("점수 정정 데이터 읽기", e)
            data_items = []
        patches = {}
        todo = [t for t in targets if t[4]][:self.limit]
        for i, (rid, rec, s, why, student) in enumerate(todo, 1):
            self.progress(i, len(todo), record_set_name(rec))
            try:
                result, html_p, json_p, err = regrade_record(
                    rec, s, grade_py=self.grade_py, student=student,
                    when=self.when, timeout=self.timeout)
            except Exception as e:
                log_error("자동 재채점 실행", e)
                result, html_p, json_p, err = None, None, None, str(e)
            if result is None:
                summary["errors"].append(f"{record_set_name(rec)}: {err}")
                continue
            summary["regraded"] += 1
            patch = dict(rec)
            ev = record_evidence_changes(patch, s, grade_py=self.grade_py,
                                         student=student)
            total = result.get("total")
            bonus = correction_deltas(data_items, rid, rec)
            reason = why
            if bonus and isinstance(total, (int, float)):
                total = total + sum(d for d, _ in bonus)
                reason = why + " + " + ", ".join(
                    f"{d:+g}점({r})" for d, r in bonus)
            patch[CORRECTION_BONUS_FIELD] = [
                correction_delta_key(rid, d) for d, _ in bonus]
            entry = apply_correction(
                patch, total, reason, when=self.when,
                source="재채점", key=f"regrade:{rid}:{reason}",
                report=html_p if os.path.isfile(html_p) else None,
                result=result, evidence=ev)
            patches[rid] = {k: v for k, v in patch.items()
                            if rec.get(k) != v or k == "점수"}
            if entry:
                summary["changes"].append({
                    "id": rid, "name": record_set_name(rec),
                    "date": str(rec.get("일시") or "")[:10],
                    "before": entry["이전총점"], "after": entry["새총점"],
                    "reason": entry["사유"], "source": "재채점",
                    "report": html_p})
                self.log(f"자동 재채점: {record_set_name(rec)} "
                         f"{entry['이전총점']} → {entry['새총점']}점 ({why})")
        summary["skipped"] = len([t for t in targets if not t[4]])
        try:
            summary["data_applied"] = self._apply_data(records, patches,
                                                       summary)
        except Exception as e:
            log_error("점수 정정 데이터 반영", e)
        self._commit(patches)
        return summary

    def _apply_data(self, records, patches, summary):
        items = load_corrections(self.corrections_file)
        if not items:
            return 0
        by_id = {rid: rec for rid, rec in records_with_ids(records)}
        n = 0
        for it in items:
            rid, rec = resolve_correction_target(by_id, it["report_id"])
            if rec is None or rid in patches:
                continue
            if not correction_matches(it, rec):
                continue
            bonus_key = None
            if it.get("delta"):
                base = _coerce_score(rec.get("점수"))
                bonus_key = correction_delta_key(rid, it["delta"])
                if base is None or correction_bonus_applied(rec, bonus_key):
                    continue
                target, key = base + it["delta"], bonus_key
            else:
                target = it["new_total"]
                key = f"data:{rid}:{target:g}"
            patch = dict(rec)
            entry = apply_correction(
                patch, target, it["reason"], when=self.when,
                source="정정데이터", key=key)
            if entry is not None and bonus_key:
                prev = rec.get(CORRECTION_BONUS_FIELD)
                patch[CORRECTION_BONUS_FIELD] = (
                    list(prev) if isinstance(prev, list) else []) + [bonus_key]
            if entry is None:
                continue
            patches[rid] = {k: v for k, v in patch.items()
                            if rec.get(k) != v or k == "점수"}
            n += 1
            summary["changes"].append({
                "id": rid, "name": record_set_name(rec),
                "date": str(rec.get("일시") or "")[:10],
                "before": entry["이전총점"], "after": entry["새총점"],
                "reason": entry["사유"], "source": "정정데이터",
                "report": rec.get("리포트")})
            self.log(f"점수 정정 데이터 반영: {record_set_name(rec)} "
                     f"{entry['이전총점']} → {entry['새총점']}점 "
                     f"({it['reason']})")
        return n

ROUTINE_PAGE_HEAD = ('<!doctype html><html lang="ko"><head><meta charset="utf-8">'
                     '<meta name="viewport" content="width=device-width, '
                     'initial-scale=1"></head><body>')
ROUTINE_PAGE_TAIL = ('<script>(function(){try{var h=(location.hash||"")'
                     '.replace(/^#/,"");if(!h){return;}var t=document.'
                     'getElementById("tab-"+h);if(t&&t.getAttribute("role")'
                     '==="tab"){t.click();}}catch(e){}})();</script>'
                     '</body></html>')
ROUTINE_ACTIONS = ("start_exam", "open_review", "start_retry", "open_report",
                   "open_pdf", "show")
ROUTINE_BODY_LIMIT = 1024 * 1024
ROUTINE_UI_TIMEOUT = 8.0
ROUTINE_ACTION_TIMEOUT = 1.5
_RECORD_STAMP_RE = re.compile(
    r"(\d{4})-?(\d{2})-?(\d{2})[ T]?(\d{2}):?(\d{2})(?::?(\d{2}))?")

def routine_html_path(base_dir=None):
    base_dir = base_dir or BASE_DIR
    cands = [os.path.join(base_dir, ROUTINE_HTML_NAME)]
    try:
        alt = _data_target_path(ROUTINE_DATA_KEY, base_dir)
        if os.path.abspath(alt) != os.path.abspath(cands[0]):
            cands.append(alt)
    except Exception:
        pass
    for c in cands:
        if os.path.isfile(c):
            return c
    return cands[0]

def render_routine_page(text):
    if text.lstrip()[:9].lower().startswith("<!doctype"):
        return text
    return ROUTINE_PAGE_HEAD + text + ROUTINE_PAGE_TAIL

def routine_missing_page(path):
    import html as _html
    return (ROUTINE_PAGE_HEAD
            + '<meta http-equiv="refresh" content="15">'
            '<div style="font-family:sans-serif;max-width:640px;margin:60px '
            'auto;line-height:1.7"><h1>루틴 페이지 파일이 아직 없습니다</h1>'
            '<p>시험장 폴더에 <code>루틴.html</code> 이 없습니다. 자동 업데이트가 '
            '켜져 있으면 곧 내려받고, 시험장 화면의 [업데이트 확인]을 누르면 '
            '바로 받습니다. 이 페이지는 15초마다 다시 확인합니다.</p>'
            f'<p>찾는 위치: <code>{_html.escape(str(path))}</code></p></div>'
            '</body></html>')

def record_stamp_id(when):
    m = _RECORD_STAMP_RE.search(str(when or ""))
    if not m:
        return None
    y, mo, d, h, mi, sec = m.groups()
    return f"{y}{mo}{d}{h}{mi}{sec or '00'}"

def records_with_ids(records):
    used = set()
    out = []
    for i, r in enumerate(records or []):
        base = record_stamp_id((r or {}).get("일시")) or f"00000000{i:06d}"
        rid, n = base, 0
        while rid in used:
            n += 1
            rid = f"{base[:12]}{n:02d}" if len(base) == 14 else f"{base}-{n}"
        used.add(rid)
        out.append((rid, r))
    return out

def _record_datetime(r):
    try:
        return datetime.strptime(str(r.get("일시"))[:16], "%Y-%m-%d %H:%M")
    except Exception:
        return None

def record_set_name(r):
    name = str((r or {}).get("세트명") or "?")
    return name[:-5] if name.endswith(" (연습)") else name

def record_report_paths(r):
    html_p = (r or {}).get("리포트")
    if not html_p or not isinstance(html_p, str):
        return None, None
    json_p = os.path.splitext(html_p)[0] + ".json"
    return (html_p if os.path.isfile(html_p) else None,
            json_p if os.path.isfile(json_p) else None)

def _num(v, default=0):
    if isinstance(v, bool) or not isinstance(v, (int, float)):
        return default
    return int(v) if float(v).is_integer() else v

def _category_of_sheet(sheet):
    head = str(sheet or "").split("-")[0].strip()
    if head in ("매크로작업", "차트작업"):
        return "기타작업"
    return head or "기타작업"

_RESULT_CACHE = {}

def load_result_summary(json_path):
    try:
        st = os.stat(json_path)
    except OSError:
        return None
    key = (st.st_mtime, st.st_size)
    hit = _RESULT_CACHE.get(json_path)
    if hit and hit[0] == key:
        return hit[1]
    data, items = load_wrong_items(json_path)
    if data is None:
        return None
    sheets = [{"name": str(sh.get("name", "?")), "alloc": _num(sh.get("alloc")),
               "earned": _num(sh.get("earned"))}
              for sh in (data.get("sheets") or []) if isinstance(sh, dict)]
    wrong = [{"label": str(it.get("label") or "?"),
              "category": str(it.get("category") or
                              _category_of_sheet(it.get("sheet"))),
              "lost": _num(it.get("lost")), "sheet": str(it.get("sheet") or "")}
             for it in items]
    summ = {"sheets": sheets, "wrong_items": wrong, "count": len(items),
            "total": data.get("total")}
    _RESULT_CACHE[json_path] = (key, summ)
    return summ

def serialize_corrections(records, sets=None, limit=20):
    by_name = {s.get("name"): s for s in sets or []}
    out = []
    for rid, r in records_with_ids(records):
        hist = (r or {}).get(CORRECTION_HISTORY_FIELD)
        if not isinstance(hist, list):
            continue
        name = record_set_name(r)
        s = by_name.get(name)
        norm = s["norm"] if s else (norm_set_key(name) if name != "?" else "")
        dt = _record_datetime(r)
        for h in hist:
            if not isinstance(h, dict):
                continue
            after = _coerce_score(h.get("새총점"))
            if after is None:
                continue
            before = _coerce_score(h.get("이전총점"))
            out.append({
                "record_id": rid,
                "set": {"name": name, "norm": norm},
                "date": dt.strftime("%Y-%m-%d") if dt
                        else str(r.get("일시"))[:10],
                "before": before, "after": after,
                "delta": (after - before) if before is not None else None,
                "reason": str(h.get("사유") or ""),
                "source": str(h.get("출처") or "재채점"),
                "when": str(h.get("일시") or ""),
            })
    out.sort(key=lambda c: (c["when"], c["record_id"]), reverse=True)
    return out[:max(0, int(limit))]

def serialize_records(records, sets, cfg=None):
    by_name = {s.get("name"): s for s in sets or []}
    out, review, index = [], {}, {}
    for rid, r in records_with_ids(records):
        name = record_set_name(r)
        s = by_name.get(name)
        norm = s["norm"] if s else (norm_set_key(name) if name != "?" else "")
        dt = _record_datetime(r)
        html_p, json_p = record_report_paths(r)
        summ = load_result_summary(json_p) if json_p else None
        if summ is None:
            json_p = None
        mode = str(r.get("mode") or "시험")
        total = r.get("점수")
        if isinstance(total, bool) or not isinstance(total, (int, float)):
            total = None
        rec = {
            "id": rid,
            "date": dt.strftime("%Y-%m-%d") if dt else str(r.get("일시"))[:10],
            "time": dt.strftime("%H:%M") if dt else "",
            "set": {"name": name, "norm": norm},
            "mode": mode,
            "total": _num(total, None) if total is not None else None,
            "pass_line": PASS_LINE,
            "passed": total is not None and total >= PASS_LINE,
            "elapsed": str(r.get("소요시간") or "") or None,
            "elapsed_sec": parse_elapsed(r.get("소요초")
                                         if r.get("소요초") is not None
                                         else r.get("소요시간")),
            "report_html": os.path.basename(html_p) if html_p else None,
            "report_json": os.path.basename(json_p) if json_p else None,
            "sheets": summ["sheets"] if summ else [],
            "wrong_items": summ["wrong_items"] if summ else [],
        }
        out.append(rec)
        index[rid] = {"record": r, "set": s, "norm": norm, "name": name,
                      "html": html_p, "json": json_p,
                      "count": summ["count"] if summ else 0,
                      "date": rec["date"], "time": rec["time"], "mode": mode}
        if json_p:
            done, when = review_done_state(norm, os.path.basename(json_p),
                                           total=index[rid]["count"], cfg=cfg)
            review[rid] = {"done": done, "when": when}
    return out, review, index

def serialize_sets(sets, records, index, review, board=None):
    by_norm = {r["norm"]: r for r in (board or [])}
    out = []
    for s in sets or []:
        recs = set_exam_records(s["name"], records or [])
        mine = [(k, v) for k, v in index.items() if v["norm"] == s.get("norm")]
        exams = [(k, v) for k, v in mine if v["mode"] not in ("부분연습",
                                                              "오답재풀이")]
        review_done = retry_done = False
        if exams:
            lk, lv = max(exams, key=lambda kv: (kv[1]["date"], kv[1]["time"],
                                               kv[0]))
            review_done = bool(review.get(lk, {}).get("done")) or (
                lv["json"] is not None and lv["count"] == 0)
            retry_done = any(
                v["mode"] == "오답재풀이"
                and (v["date"], v["time"], k) > (lv["date"], lv["time"], lk)
                for k, v in mine)
        row = {
            "name": s["name"], "norm": s.get("norm") or "",
            "pdf": bool(s.get("pdf")) and os.path.isfile(str(s.get("pdf"))),
            "key": bool(s.get("key")),
            "attempts": len(recs), "best": _best(recs),
            "review_done": review_done, "retry_done": retry_done,
        }
        bd = by_norm.get(row["norm"])
        if bd:
            row.update({"band": bd["band"], "scores": list(bd["scores"]),
                        "last_date": bd["last_date"],
                        "prev_best": bd["prev_best"],
                        "prev_band": bd["prev_band"],
                        "best_delta": bd["best_delta"],
                        "improved": bd["improved"]})
        else:
            row.update({"band": score_band(row["best"]), "scores": [],
                        "last_date": None, "prev_best": None,
                        "prev_band": "none", "best_delta": None,
                        "improved": False})
        out.append(row)
    return out

def _iso(d):
    return d.isoformat() if isinstance(d, date) else None

def _seg_end_for(d, exam=None, today=None):
    for sg in study_segments(exam, today):
        if sg["seg_start"] <= d <= sg["end"]:
            return sg["end"]
    return d

def serialize_slot(sl, by_set):
    s = sl.get("set")
    kind = sl.get("kind") or "first"
    if kind not in ("first", "redo", "retry", "goal"):
        kind = "retry" if s is None else "first"
    return {
        "kind": kind,
        "set": ({"name": s["name"], "norm": s.get("norm") or "",
                 "pdf": bool(s.get("pdf")) and os.path.isfile(str(s.get("pdf")))}
                if s else None),
        "goal": sl.get("goal"),
        "why": str(sl.get("why") or ""),
        "best": _best(by_set.get(s["norm"], [])) if s else None,
        "counts_as_retry": bool(sl.get("counts_as_retry")),
        "done": False, "record_id": None,
    }

def _day_title(d, today, kind, slots, promoted, deadline, recs_by_date,
               exam=None):
    if kind == "exam":
        ex = _exams(exam)
        i = ex.index(d) if d in ex else 0
        return "시험일" if i == 0 else f"{i + 1}차 시험일"
    if kind == "next":
        return "일정 밖"
    if kind == "past":
        recs = recs_by_date.get(d) or []
        return "완료" if recs else "응시 없음"
    parts = []
    if slots:
        parts.append(f"추천 {len(slots)}세트")
    elif d == today and recs_by_date.get(d):
        parts.append("오늘 응시함")
    else:
        parts.append("자유 복습")
    if deadline:
        parts.append("실수 노트")
    return " + ".join(parts)

def serialize_days(adaptive, today, recs_by_date, by_set, dates=None,
                   exam=None):
    all_ex = _exams(exam)
    ex = effective_exams(all_ex, today)
    week = bool(adaptive) and adaptive.get("kind") in ("week", "free")
    end = adaptive.get("end") if week else None
    boost = set(adaptive.get("boost_days") or []) if week else set()
    days_map = (adaptive.get("days") or {}) if week else {}
    free_map = (adaptive.get("free") or {}) if week else {}
    deadlines = {sg["end"] for sg in study_segments(ex, today)}
    if dates is None:
        if not ex:
            return []
        dates = []
        d = ex[0] - timedelta(days=ROUTINE_DAYS)
        while d <= max(ex):
            dates.append(d)
            d += timedelta(days=1)
    out = []
    for d in dates:
        slots = []
        if d in all_ex:
            kind, capacity = "exam", 0
        elif week and d > end:
            kind, capacity = "next", base_capacity(d, _seg_end_for(d, ex, today))
        elif d < today:
            kind, capacity = "past", base_capacity(d, _seg_end_for(d, ex, today))
        elif not week:
            kind, capacity = "next", base_capacity(d, _seg_end_for(d, ex, today))
        else:
            slots = [serialize_slot(sl, by_set) for sl in days_map.get(d, [])]
            kind = "mock" if slots else "rest"
            capacity = base_capacity(d, end) + (1 if d in boost else 0)
        promoted = d in boost
        out.append({
            "date": d.isoformat(), "no": routine_day_no(d, ex), "kind": kind,
            "title": _day_title(d, today, kind, slots, promoted, d in deadlines,
                                recs_by_date, all_ex),
            "capacity": capacity, "promoted": promoted,
            "deadline": d in deadlines, "review_day": False,
            "free": int(free_map.get(d, 0)) if kind in ("mock", "rest") else 0,
            "slots": slots,
        })
    return out

def _recs_by_date(records):
    out = {}
    for r in records or []:
        if r.get("mode") in ("부분연습", "오답재풀이"):
            continue
        d = _record_date(r)
        if d:
            out.setdefault(d, []).append(r)
    return out

def serialize_plan(adaptive, today, records, sets, preview=None, exam=None):
    ex = _exams(exam if exam is not None else adaptive.get("exam_dates"))
    week = adaptive.get("kind") in ("week", "free")
    by_set = exam_records_by_set(sets or [], records or [])
    recs_by_date = adaptive.get("_recs_by_date") if week else None
    if not recs_by_date:
        recs_by_date = _recs_by_date(records)
    lv = adaptive.get("level") or level_info(records or [])
    stk = adaptive.get("streak") or streak_info(records or [], today, ex)
    plan = {
        "kind": adaptive.get("kind"),
        "level": lv,
        "streak": stk,
        "attempts_left": int(adaptive.get("attempts_left") or 0),
        "today_locked": False,
        "passed_sets": int(adaptive.get("passed_sets") or 0),
        "total_sets": int(adaptive.get("total_sets") or len(sets or [])),
        "all_clear": bool(adaptive.get("all_clear")),
        "retry_queue": [{"kind": str(q.get("kind") or "retry"),
                         "set": str(q.get("norm") or ""),
                         "name": str(q.get("name") or ""),
                         "best": q.get("best")}
                        for q in adaptive.get("retry_queue") or []],
        "seg": adaptive.get("seg") if week else None,
        "today": today.isoformat(),
        "exam_dates": [d.isoformat() for d in ex],
        "exam_set": bool(ex),
        "dday": dday_num(today, ex),
        "no_sets": bool(adaptive.get("no_sets")),
        "routine_days": ROUTINE_DAYS,
        "start": _iso(adaptive.get("start")) if week else None,
        "end": _iso(adaptive.get("end")) if week else None,
        "seg_start": _iso(adaptive.get("seg_start")) if week else None,
        "reason": str(adaptive.get("reason") or ""),
        "warning": adaptive.get("warning"),
        "remaining": int(adaptive.get("remaining") or 0),
        "retry_target": int(adaptive.get("retry_target") or 0),
        "missed_days": [_iso(x) for x in adaptive.get("missed_days") or []],
        "boost_days": [_iso(x) for x in adaptive.get("boost_days") or []],
        "demoted": [str(x) for x in adaptive.get("demoted") or []],
        "days": serialize_days(adaptive, today, recs_by_date, by_set, exam=ex),
        "preview": None,
    }
    if week and preview and preview.get("kind") in ("week", "free"):
        p_today = preview.get("today") or preview["start"]
        dates = []
        d = preview["start"]
        while d <= preview["end"]:
            dates.append(d)
            d += timedelta(days=1)
        plan["preview"] = {
            "days": serialize_days(preview, p_today, recs_by_date, by_set,
                                   dates=dates, exam=ex),
            "boost_days": [_iso(x) for x in preview.get("boost_days") or []],
        }
    return plan

def build_state(sets, records=None, cfg=None, today=None, exam_running=False,
                generated=None, exam=None):
    today = today or date.today()
    if records is None:
        records = load_records(RECORDS_PATH)
    if cfg is None:
        cfg = load_set_config(SET_CONFIG_PATH)
    ex = _exams(exam) if exam is not None else exam_dates(cfg=cfg)
    sets = list(sets or [])
    try:
        adaptive = build_adaptive_plan(today, records, sets, ex)
    except Exception as e:
        log_error("루틴 연동 일정 계산", e)
        adaptive = {"kind": segment_for(today, ex)["kind"], "today": today,
                    "days": {}, "reason": "", "warning": f"일정 계산 오류: {e}",
                    "exam_dates": list(ex)}
    preview = None
    nxt = next_segment_after(adaptive.get("seg") or 0, ex, today) \
        if adaptive.get("kind") in ("week", "free") else None
    if nxt:
        try:
            preview = build_adaptive_plan(nxt["seg_start"], records, sets, ex)
        except Exception as e:
            log_error("루틴 연동 다음 구간 미리보기", e)
    ex_act = effective_exams(ex, today)
    recs, review, index = serialize_records(records, sets, cfg)
    checks = {str(k): bool(v) for k, v in
              _cfg_section(cfg, WEB_CHECKS_KEY).items()}
    plan_json = serialize_plan(adaptive, today, records, sets, preview, ex)
    board = score_board(sets, recs)
    return {
        "version": ROUTINE_API_VERSION,
        "app_version": __version__,
        "today": today.isoformat(),
        "generated": generated or routine_generated(),
        "exam_running": bool(exam_running),
        "level": plan_json["level"],
        "streak": plan_json["streak"],
        "attempts_left": plan_json["attempts_left"],
        "today_locked": plan_json["today_locked"],
        "plan": plan_json,
        "bands": board["bands"],
        "sets": serialize_sets(sets, records, index, review, board["sets"]),
        "records": recs,
        "achievements": achievements(recs),
        "corrections": serialize_corrections(records, sets),
        "review": review,
        "checks": checks,
        "exam": {"dates": [d.isoformat() for d in ex],
                 "active": [d.isoformat() for d in ex_act],
                 "past": [d.isoformat() for d in ex if d < today],
                 "next": (ex_act[0].isoformat() if ex_act else None),
                 "set": bool(ex),
                 "dday": dday_num(today, ex),
                 "routine_days": ROUTINE_DAYS,
                 "start": _iso(routine_start(ex, today))},
        "settings": {"auto_open_routine": bool(
            _cfg_section(cfg, "_설정").get(ROUTINE_AUTO_OPEN_SETTING, False)),
                     "auto_update": bool(_cfg_section(cfg, "_설정").get(
                         AUTO_UPDATE_SETTING, True)),
                     "scan_root": str(_cfg_section(cfg, "_설정").get(
                         SCAN_ROOT_SETTING) or ""),
                     "hidden_sets": hidden_set_keys(cfg),
                     "exam_dates": [d.isoformat() for d in ex],
                     "exam_date": (ex[0].isoformat() if ex else None),
                     "exam_date2": (ex[1].isoformat() if len(ex) > 1 else None)},
    }

def record_index(sets, records=None, cfg=None):
    if records is None:
        records = load_records(RECORDS_PATH)
    if cfg is None:
        cfg = load_set_config(SET_CONFIG_PATH)
    return serialize_records(records, sets, cfg)[2]

class RoutineRequestError(Exception):

    def __init__(self, status, message):
        super().__init__(message)
        self.status = status
        self.message = message

class _RoutineHTTPServer(http.server.ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = sys.platform != "win32"
    routine = None

class _RoutineHandler(http.server.BaseHTTPRequestHandler):
    server_version = "CocoRoutine/" + __version__
    sys_version = ""

    def log_message(self, fmt, *args):
        pass

    def _send(self, status, body, ctype):
        data = body if isinstance(body, bytes) else body.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(data)

    def _json(self, status, obj):
        self._send(status, json.dumps(obj, ensure_ascii=False),
                   "application/json; charset=utf-8")

    def _html(self, status, text):
        self._send(status, text, "text/html; charset=utf-8")

    def _host_ok(self):
        host = (self.headers.get("Host") or "").strip().lower()
        if host.startswith("["):
            host = host.split("]")[0] + "]"
        else:
            host = host.split(":")[0]
        return host in ("127.0.0.1", "localhost", "[::1]", "")

    def _token_ok(self, query):
        srv = self.server.routine
        given = self.headers.get("X-Coco-Token") or ""
        if not given:
            given = (query.get("t") or [""])[0]
        return bool(given) and secrets.compare_digest(str(given), srv.token)

    def _read_body(self):
        try:
            n = int(self.headers.get("Content-Length") or 0)
        except ValueError:
            raise RoutineRequestError(400, "Content-Length 오류")
        if n < 0 or n > ROUTINE_BODY_LIMIT:
            raise RoutineRequestError(400, "본문이 너무 큽니다")
        raw = self.rfile.read(n) if n else b""
        try:
            body = json.loads(raw.decode("utf-8") or "{}")
        except (UnicodeDecodeError, ValueError):
            raise RoutineRequestError(400, "JSON 파싱 실패")
        if not isinstance(body, dict):
            raise RoutineRequestError(400, "JSON 객체가 아닙니다")
        return body

    def do_HEAD(self):
        self.do_GET()

    def do_GET(self):
        srv = self.server.routine
        parts = urllib.parse.urlsplit(self.path)
        path = parts.path.rstrip("/") or "/"
        try:
            if not self._host_ok():
                srv.note("GET", path, 403)
                return self._json(403, {"ok": False, "error": "허용되지 않는 Host"})
            if path in ("/", "/index.html", "/루틴.html"):
                status, page = srv.page()
                srv.note("GET", "/", status)
                return self._html(status, page)
            if path == "/favicon.ico":
                return self._send(204, b"", "image/x-icon")
            if path == "/api/state":
                srv.stats["state"] += 1
                return self._json(200, srv.state())
            if path.startswith("/api/report/"):
                rid = urllib.parse.unquote(path[len("/api/report/"):])
                fp = srv.report_file(rid)
                if not fp:
                    srv.note("GET", path, 404)
                    return self._json(404, {"ok": False,
                                            "error": "리포트를 찾을 수 없습니다"})
                with open(fp, "rb") as f:
                    data = f.read()
                srv.note("GET", f"/api/report/{rid}", 200)
                return self._send(200, data, "text/html; charset=utf-8")
            srv.note("GET", path, 404)
            return self._json(404, {"ok": False, "error": "없는 경로"})
        except Exception as e:
            log_error(f"루틴 서버 GET {path}", e)
            srv.note("GET", path, 500)
            try:
                self._json(500, {"ok": False, "error": f"서버 오류: {e}"})
            except Exception:
                pass

    def do_POST(self):
        srv = self.server.routine
        parts = urllib.parse.urlsplit(self.path)
        path = parts.path.rstrip("/") or "/"
        query = urllib.parse.parse_qs(parts.query)
        status = 500
        try:
            if not self._host_ok():
                raise RoutineRequestError(403, "허용되지 않는 Host")
            if not path.startswith("/api/"):
                raise RoutineRequestError(404, "없는 경로")
            if not self._token_ok(query):
                raise RoutineRequestError(401, "토큰 불일치")
            body = self._read_body()
            handler = {"/api/check": srv.api_check, "/api/score": srv.api_score,
                       "/api/review": srv.api_review,
                       "/api/exam_date": srv.api_exam_date,
                       "/api/action": srv.api_action}.get(path)
            if handler is None:
                raise RoutineRequestError(404, "없는 경로")
            status, payload = handler(body)
            srv.note("POST", path, status, payload)
            return self._json(status, payload)
        except RoutineRequestError as e:
            status = e.status
            srv.note("POST", path, status, {"error": e.message})
            return self._json(status, {"ok": False, "error": e.message})
        except Exception as e:
            log_error(f"루틴 서버 POST {path}", e)
            srv.note("POST", path, 500, {"error": str(e)})
            try:
                self._json(500, {"ok": False, "error": f"서버 오류: {e}"})
            except Exception:
                pass

class RoutineServer:

    def __init__(self, sets_fn=None, ui_call=None, action_fn=None,
                 on_change=None, exam_running_fn=None, port=None,
                 html_path=None, today=None, log=None):
        self.sets_fn = sets_fn or (lambda: [])
        self.ui_call = ui_call
        self.action_fn = action_fn
        self.on_change = on_change
        self.exam_running_fn = exam_running_fn or (lambda: False)
        self.preferred_port = ROUTINE_PORT_DEFAULT if port is None else int(port)
        self._html_path = html_path
        self._today = today
        self.log = log or startup_log
        self.token = secrets.token_urlsafe(18)
        self.httpd = None
        self.thread = None
        self.port = None
        self.stats = {"state": 0, "requests": 0}
        self._fingerprint = None
        self._write_lock = threading.Lock()

    def start(self):
        last = None
        cands = [self.preferred_port]
        if self.preferred_port != 0:
            cands.append(0)
        for p in cands:
            try:
                self.httpd = _RoutineHTTPServer(("127.0.0.1", p), _RoutineHandler)
                break
            except OSError as e:
                last = e
                self.httpd = None
        if self.httpd is None:
            raise last or OSError("루틴 서버 포트를 열 수 없습니다")
        self.httpd.routine = self
        self.port = int(self.httpd.server_address[1])
        self.thread = threading.Thread(target=self.httpd.serve_forever,
                                       kwargs={"poll_interval": 0.5},
                                       daemon=True, name="루틴서버")
        self.thread.start()
        hp = self.html_path()
        self.log(f"루틴 서버 시작: http://127.0.0.1:{self.port}/ "
                 f"(우선 포트 {self.preferred_port}"
                 f"{' 사용 중 → 대체' if self.port != self.preferred_port else ''}"
                 f", 토큰 {self.token[:4]}…) · 페이지="
                 f"{hp if os.path.isfile(hp) else '없음(' + hp + ')'}")
        return self.port

    def stop(self):
        httpd, self.httpd = self.httpd, None
        if httpd is None:
            return
        try:
            httpd.shutdown()
            httpd.server_close()
        except Exception:
            pass
        self.log(f"루틴 서버 종료 (상태 조회 {self.stats['state']}회, "
                 f"기타 요청 {self.stats['requests']}회)")

    @property
    def running(self):
        return self.httpd is not None and self.thread is not None \
            and self.thread.is_alive()

    def url(self, tab=None):
        u = f"http://127.0.0.1:{self.port}/?t={self.token}"
        return u + (f"#{tab}" if tab else "")

    def html_path(self):
        return self._html_path or routine_html_path()

    def today(self):
        t = self._today
        if callable(t):
            t = t()
        return t or date.today()

    def note(self, method, path, status, payload=None):
        self.stats["requests"] += 1
        err = ""
        if isinstance(payload, dict) and payload.get("ok") is False:
            err = f" · {payload.get('error')}"
        elif isinstance(payload, dict) and "error" in payload \
                and "ok" not in payload:
            err = f" · {payload.get('error')}"
        try:
            self.log(f"루틴 서버: {method} {path} → {status}{err}")
        except Exception:
            pass

    def page(self):
        hp = self.html_path()
        try:
            with open(hp, encoding="utf-8-sig") as f:
                return 200, render_routine_page(f.read())
        except OSError:
            return 200, routine_missing_page(hp)

    def _touch_if_files_changed(self):
        fp = []
        for p in (RECORDS_PATH, SET_CONFIG_PATH):
            try:
                st = os.stat(p)
                fp.append((st.st_mtime_ns, st.st_size))
            except OSError:
                fp.append(None)
        fp = tuple(fp)
        if self._fingerprint is not None and fp != self._fingerprint:
            routine_touch()
        self._fingerprint = fp

    def state(self):
        self._touch_if_files_changed()
        return build_state(self.sets_fn(), today=self.today(),
                           exam_running=bool(self.exam_running_fn()))

    def _ok(self):
        st = self.state()
        return 200, dict([("ok", True)] + list(st.items()))

    def report_file(self, rid):
        rid = str(rid or "").strip()
        if not rid or any(c in rid for c in "/\\") or ".." in rid:
            return None
        info = record_index(self.sets_fn()).get(rid)
        return info["html"] if info and info.get("html") else None

    def _run_on_ui(self, fn, timeout):
        if self.ui_call is None:
            return True, fn()
        box = {}
        ev = threading.Event()

        def job():
            try:
                box["r"] = fn()
            except Exception as e:
                box["e"] = e
            finally:
                ev.set()

        self.ui_call(job)
        if not ev.wait(timeout):
            return False, None
        if "e" in box:
            raise box["e"]
        return True, box.get("r")

    def _write(self, fn):
        try:
            with self._write_lock:
                done, _r = self._run_on_ui(fn, ROUTINE_UI_TIMEOUT)
        except Exception as e:
            log_error("루틴 연동 쓰기", e)
            return 200, {"ok": False, "error": f"저장 실패: {e}"}
        if not done:
            return 200, {"ok": False, "error": "시험장 창이 응답하지 않습니다 — "
                                              "잠시 후 다시 시도하세요"}
        return self._ok()

    def _changed(self, kind, info):
        if self.on_change is not None:
            try:
                self.on_change(kind, info)
            except Exception as e:
                log_error(f"루틴 연동 반영({kind})", e)

    def api_check(self, body):
        key = body.get("key")
        if not isinstance(key, str) or not key.strip() or len(key) > 120 \
                or any(ord(c) < 32 for c in key):
            raise RoutineRequestError(400, "key 누락 또는 형식 오류")
        key = key.strip()
        value = body.get("value")
        on = value is True or value == 1 or (
            isinstance(value, str) and value.lower() in ("true", "1", "on"))

        def job():
            cfg = load_set_config(SET_CONFIG_PATH)
            wc = cfg.get(WEB_CHECKS_KEY)
            if not isinstance(wc, dict):
                wc = cfg[WEB_CHECKS_KEY] = {}
            if on:
                wc[key] = True
            else:
                wc.pop(key, None)
            if not save_set_config(cfg, SET_CONFIG_PATH):
                raise OSError("세트설정.json 저장 실패")
            self._changed("check", {"key": key, "value": on})

        return self._write(job)

    def api_exam_date(self, body):
        if "dates" in body:
            raw = body.get("dates")
            if raw is None:
                raw = []
            if not isinstance(raw, (list, tuple)):
                raise RoutineRequestError(400, "dates 는 배열이어야 합니다")
            out = []
            for x in raw:
                d = parse_iso_date(x)
                if d is None:
                    raise RoutineRequestError(
                        400, f"dates 형식 오류 (YYYY-MM-DD): {x}")
                if d not in out:
                    out.append(d)
            out.sort()

            def job_list():
                save_exam_date_list(out, SET_CONFIG_PATH)
                self._changed("exam_date",
                              {"dates": [d.isoformat() for d in out]})

            return self._write(job_list)
        raw1, raw2 = body.get("date"), body.get("date2")
        d1 = parse_iso_date(raw1) if raw1 else None
        d2 = parse_iso_date(raw2) if raw2 else None
        if raw1 and d1 is None:
            raise RoutineRequestError(400, "date 형식 오류 (YYYY-MM-DD)")
        if raw2 and d2 is None:
            raise RoutineRequestError(400, "date2 형식 오류 (YYYY-MM-DD)")
        if d1 is not None and d2 is not None and d2 <= d1:
            raise RoutineRequestError(400, "date2 는 date 보다 뒤여야 합니다")

        def job():
            save_exam_dates(d1, d2, SET_CONFIG_PATH)
            self._changed("exam_date",
                          {"date": d1.isoformat() if d1 else None,
                           "date2": d2.isoformat() if d2 else None})

        return self._write(job)

    def api_score(self, body):
        d_txt = body.get("date")
        try:
            d = datetime.strptime(str(d_txt), "%Y-%m-%d").date()
        except (TypeError, ValueError):
            raise RoutineRequestError(400, "date 형식 오류 (YYYY-MM-DD)")
        norm = body.get("set")
        if not isinstance(norm, str) or not norm:
            raise RoutineRequestError(400, "set(norm) 누락")
        s = next((x for x in self.sets_fn() if x.get("norm") == norm), None)
        if s is None:
            raise RoutineRequestError(400, "알 수 없는 세트")
        total = body.get("total")
        if isinstance(total, bool) or not isinstance(total, (int, float)) \
                or not float(total).is_integer() or not 0 <= total <= 100:
            raise RoutineRequestError(400, "total 은 0~100 정수")
        total = int(total)
        no = routine_day_no(d)
        record = {
            "일시": f"{d.isoformat()} {datetime.now().strftime('%H:%M')}",
            "세트명": s["name"], "점수": total, "소요시간": "-", "리포트": None,
            "mode": "수동", "출처": "루틴페이지",
        }
        if 1 <= no <= ROUTINE_DAYS:
            record["day"] = plan_day_tag(no)
            record["루틴"] = routine_tag()

        def job():
            append_record(record, RECORDS_PATH)
            self._changed("score", {"record": record, "set": s})

        return self._write(job)

    def api_review(self, body):
        rid = body.get("record_id")
        if not isinstance(rid, str) or not rid:
            raise RoutineRequestError(400, "record_id 누락")
        done = body.get("done") is True
        info = record_index(self.sets_fn()).get(rid)
        if info is None:
            return 200, {"ok": False, "error": "기록을 찾을 수 없습니다"}
        if not info.get("json"):
            return 200, {"ok": False, "error": "채점결과 JSON이 없는 기록은 오답노트 "
                                              "상태를 저장할 수 없습니다"}
        json_name = os.path.basename(info["json"])
        count = int(info.get("count") or 0)

        def job():
            checks = set(range(count)) if done else set()
            if not record_review_state(info["norm"], json_name, checks,
                                       SET_CONFIG_PATH, done=done):
                raise OSError("세트설정.json 저장 실패")
            self._changed("review", {"record_id": rid, "done": done,
                                     "norm": info["norm"], "json": json_name})

        return self._write(job)

    def api_action(self, body):
        action = body.get("action")
        if action not in ROUTINE_ACTIONS:
            raise RoutineRequestError(400, "알 수 없는 action")
        norm = body.get("set")
        rid = body.get("record_id")
        norm = norm if isinstance(norm, str) and norm else None
        rid = rid if isinstance(rid, str) and rid else None
        sets = self.sets_fn()
        s = next((x for x in sets if x.get("norm") == norm), None) if norm else None
        rec = record_index(sets).get(rid) if rid else None
        if norm and s is None:
            return 200, {"ok": False, "error": "알 수 없는 세트"}
        if rid and rec is None:
            return 200, {"ok": False, "error": "기록을 찾을 수 없습니다"}
        if rec is not None and s is None:
            s = rec.get("set")
        if action in ("start_exam", "start_retry") and self.exam_running_fn():
            return 200, {"ok": False, "error": "시험 진행 중"}
        if action == "open_pdf":
            if s is None:
                return 200, {"ok": False, "error": "세트를 지정하세요"}
            if not s.get("pdf") or not os.path.isfile(str(s.get("pdf"))):
                return 200, {"ok": False, "error": "문제지 PDF 없음"}
        if action in ("open_review", "start_retry") and s is None and rec is None:
            return 200, {"ok": False, "error": "세트 또는 기록을 지정하세요"}
        if action == "open_report" and (rec is None or not rec.get("html")):
            return 200, {"ok": False, "error": "리포트 파일이 없습니다"}
        if self.action_fn is None:
            return self._ok()
        try:
            done, result = self._run_on_ui(
                lambda: self.action_fn(action, s, rec, norm, rid),
                ROUTINE_ACTION_TIMEOUT)
        except Exception as e:
            log_error(f"루틴 연동 동작({action})", e)
            return 200, {"ok": False, "error": f"동작 실패: {e}"}
        if done and isinstance(result, tuple) and len(result) == 2 \
                and not result[0]:
            return 200, {"ok": False, "error": str(result[1] or "동작 실패")}
        return self._ok()

def format_elapsed(seconds):
    seconds = max(0, int(seconds))
    m, s = divmod(seconds, 60)
    return f"{m}분 {s:02d}초"

def beep(widget=None):
    try:
        import winsound
        winsound.Beep(880, 180)
        winsound.Beep(660, 180)
        return
    except Exception:
        pass
    try:
        if widget is not None:
            widget.bell()
    except Exception:
        pass

if HAS_TK:

    class CollapsibleErrorDialog(tk.Toplevel):

        def __init__(self, master, title, message, detail=""):
            super().__init__(master)
            self.title(title)
            self.configure(bg=BG)
            self.resizable(True, True)
            self.transient(master)
            frm = tk.Frame(self, bg=BG, padx=16, pady=14)
            frm.pack(fill="both", expand=True)
            tk.Label(frm, text=message, bg=BG, fg=INK, font=UI_FONT,
                     justify="left", wraplength=460).pack(anchor="w")
            self._detail = (detail or "").strip()
            self._text = None
            btns = tk.Frame(frm, bg=BG)
            btns.pack(fill="x", pady=(10, 0))
            self._message = str(message)
            if self._detail:
                self._toggle_btn = tk.Button(
                    btns, text="자세히 보기", command=self._toggle,
                    font=UI_FONT, relief="groove")
                self._toggle_btn.pack(side="left")
            self.copy_btn = tk.Button(btns, text="오류 내용 복사",
                                      command=self.copy_error, font=UI_FONT,
                                      relief="groove")
            self.copy_btn.pack(side="left", padx=6)
            tk.Button(btns, text="닫기", command=self.destroy,
                      font=UI_FONT, relief="groove").pack(side="right")
            self._body = frm
            try:
                self.grab_set()
            except Exception:
                pass

        def copy_text(self):
            return (f"[{APP_TITLE} {__version__} 오류]\n{self._message}\n\n"
                    + self._detail).strip()

        def copy_error(self):
            try:
                self.clipboard_clear()
                self.clipboard_append(self.copy_text())
                self.update_idletasks()
                self.copy_btn.configure(text="복사됨")
                return True
            except Exception:
                return False

        def _toggle(self):
            if self._text is None:
                self._text = tk.Text(self._body, height=12, width=64,
                                     font=("Consolas", 9), wrap="word")
                self._text.insert("1.0", self._detail)
                self._text.configure(state="disabled")
                self._text.pack(fill="both", expand=True, pady=(10, 0))
                self._toggle_btn.configure(text="상세 접기")
            else:
                self._text.destroy()
                self._text = None
                self._toggle_btn.configure(text="자세히 보기")

    _LAST_ERROR_DIALOG = [None]

    def _tk_report_callback_exception(root, exc, val, tb):
        text = log_error("tk callback", (exc, val, tb))
        try:
            dlg = CollapsibleErrorDialog(
                root, f"{APP_TITLE} - 오류",
                f"작업 중 오류가 발생했습니다: {val}\n\n"
                f"오류 내용은 {os.path.basename(ERROR_LOG_PATH)}에 기록되었습니다. "
                "[오류 내용 복사]로 복사해 채팅에 붙여넣어 주세요.",
                text)
            _LAST_ERROR_DIALOG[0] = dlg
        except Exception:
            try:
                messagebox.showerror(APP_TITLE, f"오류: {val}\n\n{text[-800:]}")
            except Exception:
                pass

    tk.Tk.report_callback_exception = _tk_report_callback_exception

    _TK_VARS = []

    def keep_var(var):
        _TK_VARS.append(var)
        return var

    class ResultWindow(tk.Toplevel):

        def __init__(self, master, result, html_path, folder=None,
                     copied=False, goal=None, linked=False, progress=None):
            super().__init__(master)
            self.title(f"{APP_TITLE} - 채점 결과")
            self.configure(bg=BG)
            self.attributes("-topmost", True)
            total = result.get("total", 0)
            partial = result.get("mode") == "partial"
            max_total = result.get("max_total") or 100
            passed = total >= result.get("pass_line", PASS_LINE)
            frm = tk.Frame(self, bg=BG, padx=28, pady=20)
            frm.pack(fill="both", expand=True)
            head_txt = "부분 연습 결과" if partial else "채점 결과"
            tk.Label(frm, text=head_txt, bg=BG, fg=SUB,
                     font=UI_FONT).pack()
            if partial:
                names = ", ".join(result.get("graded_sheets") or [])
                tk.Label(frm, text=f"{total} / {max_total:g}점", bg=BG,
                         fg=BRAND,
                         font=("Malgun Gothic", 38, "bold")).pack()
                tk.Label(frm, text=f"채점 영역: {names}",
                         bg=BG, fg=INK, font=UI_FONT_BOLD).pack(pady=(0, 10))
            else:
                tk.Label(frm, text=f"{total}점", bg=BG,
                         fg=BRAND if passed else RED,
                         font=("Malgun Gothic", 42, "bold")).pack()
                verdict = "합격권" if passed else "미달"
                delta_txt = ""
                if progress and progress.get("prev_best") is not None:
                    d = total - progress["prev_best"]
                    delta_txt = ("  ·  직전 최고 "
                                 f"{progress['prev_best']:g}점 대비 "
                                 + (f"+{d:g}점" if d > 0 else
                                    (f"{d:g}점" if d < 0 else "동점")))
                elif progress:
                    delta_txt = "  ·  첫 기록"
                tk.Label(frm,
                         text=f"합격선 {PASS_LINE}점 기준: {verdict}{delta_txt}",
                         bg=BG, fg=INK,
                         font=UI_FONT_BOLD).pack(pady=(0, 8))
                self.score_gauge = tk.Canvas(frm, height=16, width=360, bg=BG,
                                             bd=0, highlightthickness=0)
                self.score_gauge.pack(fill="x", pady=(0, 4))
                self._gauge_target = max(0.0, min(1.0, float(total) / 100.0))
                self._gauge_passed = passed
                self._gauge_step = 0
                self._animate_gauge()
                tk.Label(frm, text=f"0 —— 합격선 {PASS_LINE} —— 100",
                         bg=BG, fg=SUB,
                         font=("Malgun Gothic", 8)).pack(pady=(0, 8))
            sheets = result.get("sheets") or []
            if sheets:
                box = tk.Frame(frm, bg=CARD, highlightbackground=LINE,
                               highlightthickness=1, padx=12, pady=8)
                box.pack(fill="x")
                for s in sheets:
                    earned, alloc = s.get("earned", 0), s.get("alloc", 0)
                    ok = earned >= alloc
                    row = tk.Frame(box, bg=CARD)
                    row.pack(fill="x")
                    tk.Label(row, text=s.get("name", "?"), bg=CARD, fg=INK,
                             font=UI_FONT, anchor="w").pack(side="left")
                    tk.Label(row, text=f"{earned:g} / {alloc:g}", bg=CARD,
                             fg=BRAND if ok else (AMBER if earned > 0 else RED),
                             font=UI_FONT_BOLD).pack(side="right")
            if goal is not None and not partial:
                diff = total - goal
                if diff >= 0:
                    nxt = next((t for t in GOAL_TIERS if t > goal), None)
                    g_txt = (f"목표 {goal}점 달성! (+{diff}점)"
                             + (f" — 다음 목표 {nxt}점" if nxt
                                else " — 만점 달성!"))
                    g_fg, g_bg = BRAND_DARK, BRAND_SOFT
                else:
                    g_txt = (f"목표 {goal}점까지 {-diff}점 — 완주·채점까지 "
                             "했으니 오늘은 클리어입니다. 리포트에서 틀린 "
                             "부분만 확인해 두세요(선택).")
                    g_fg, g_bg = "#8A5A00", "#FBF0DC"
                tk.Label(frm, text=g_txt, bg=g_bg, fg=g_fg,
                         font=("Malgun Gothic", 9, "bold"), padx=10, pady=5,
                         wraplength=360).pack(pady=(6, 0))
            if progress and not partial:
                self._render_progress(frm, progress, total)
            if linked:
                self.link_lbl = tk.Label(
                    frm, text="채점 결과가 루틴 페이지에 자동 반영되었습니다 — "
                    "열려 있는 페이지는 5초 안에 갱신됩니다 (시작 화면 "
                    "[루틴 열기])", bg=BRAND_SOFT, fg=BRAND_DARK,
                    font=("Malgun Gothic", 9, "bold"), wraplength=360,
                    padx=10, pady=6, justify="left")
                self.link_lbl.pack(pady=(10, 0))
            elif copied:
                tk.Label(frm, text="성적이 복사되었습니다 — 루틴 웹페이지에서 "
                         "[성적 붙여넣기]를 누르면 자동 기록됩니다",
                         bg=BRAND_SOFT, fg=BRAND_DARK,
                         font=("Malgun Gothic", 9, "bold"), wraplength=360,
                         padx=10, pady=6, justify="left").pack(pady=(10, 0))
            if folder and os.path.isdir(folder):
                tk.Label(frm, text=f"저장 폴더: {folder}", bg=BG, fg=SUB,
                         font=("Malgun Gothic", 8), wraplength=360,
                         justify="left").pack(pady=(10, 0))
            btns = tk.Frame(frm, bg=BG)
            btns.pack(pady=(12, 0))
            if html_path and os.path.isfile(html_path):
                tk.Button(btns, text="리포트 열기", font=UI_FONT_BOLD,
                          bg=BRAND, fg="white", activebackground=BRAND_DARK,
                          relief="flat", padx=14, pady=4,
                          command=lambda: open_file(html_path)).pack(
                    side="left", padx=6)
            if folder and os.path.isdir(folder):
                tk.Button(btns, text="폴더 열기", font=UI_FONT, relief="groove",
                          padx=14, pady=4,
                          command=lambda: open_file(folder)).pack(
                    side="left", padx=6)
            tk.Button(btns, text="닫기", font=UI_FONT, relief="groove",
                      padx=14, pady=4, command=self.destroy).pack(
                side="left", padx=6)

        STEPS = 14

        def _animate_gauge(self):
            cv = getattr(self, "score_gauge", None)
            if cv is None or not cv.winfo_exists():
                return
            self._gauge_step = min(self.STEPS, self._gauge_step + 1)
            frac = self._gauge_target * self._gauge_step / self.STEPS
            w = max(1, cv.winfo_width() or 360)
            h = int(cv["height"])
            cv.delete("all")
            cv.create_rectangle(0, 0, w, h, fill=CARD, outline=LINE)
            if frac > 0:
                cv.create_rectangle(
                    0, 0, max(2, int(w * frac)), h,
                    fill=BRAND if self._gauge_passed else AMBER, outline="")
            x = int(w * PASS_LINE / 100)
            cv.create_line(x, 0, x, h, fill=INK, width=2)
            cv.create_text(min(w - 22, x + 22), h // 2, text=str(PASS_LINE),
                           fill=INK, font=("Malgun Gothic", 8))
            if self._gauge_step < self.STEPS:
                self.after(30, self._animate_gauge)

        def _render_progress(self, frm, prog, total):
            ups = prog.get("tier_up") or []
            if ups:
                top = max(ups)
                nxt = prog.get("level_after", {}).get("next_goal")
                txt = f"{top}점 달성! {tier_label(top)} 단계" + (
                    f" — 다음 목표 {nxt}점({tier_label(nxt)})"
                    if nxt and nxt != top else " — 만점까지 왔습니다")
                tk.Label(frm, text=txt, bg=BRAND, fg="white",
                         font=("Malgun Gothic", 11, "bold"), padx=12, pady=6,
                         wraplength=360).pack(pady=(2, 6))
            badges = prog.get("new_badges") or []
            if badges:
                box = tk.Frame(frm, bg=BRAND_SOFT, padx=10, pady=6)
                box.pack(fill="x", pady=(0, 6))
                tk.Label(box, text=f"새 도전 과제 {len(badges)}개 달성",
                         bg=BRAND_SOFT, fg=BRAND_DARK,
                         font=("Malgun Gothic", 9, "bold"),
                         anchor="w").pack(fill="x")
                for a in badges[:4]:
                    tk.Label(box, text=f"✓ [{a['icon']}] {a['name']} — "
                                       f"{a['desc']}", bg=BRAND_SOFT, fg=INK,
                             font=("Malgun Gothic", 9), anchor="w",
                             justify="left", wraplength=340).pack(fill="x")
            tk.Label(frm, text=f"도전 과제 {prog.get('badges_done', 0)}/"
                               f"{prog.get('badges_total', 0)} 달성",
                     bg=BG, fg=SUB, font=("Malgun Gothic", 9)).pack()

    class TimerOverlay(tk.Toplevel):

        _live = []

        def __init__(self, master, head="", monitor_index=None,
                     monitors=None, saved=None):
            super().__init__(master)
            self._drag = None
            self._moved = False
            self.monitors = list(monitors or [])
            try:
                self.withdraw()
            except Exception:
                pass
            for setup in (lambda: self.overrideredirect(True),
                          lambda: self.attributes("-topmost", True),
                          lambda: self.title(f"{APP_TITLE} - 남은 시간")):
                try:
                    setup()
                except Exception:
                    pass
            self.configure(bg=INK, highlightthickness=1,
                           highlightbackground="#3A5F49")
            body = tk.Frame(self, bg=INK, padx=14, pady=8)
            body.pack(fill="both", expand=True)
            self.head_lbl = tk.Label(body, text=head, bg=INK, fg="#A7C8B2",
                                     font=("Malgun Gothic", 8))
            self.head_lbl.pack()
            self.time_lbl = tk.Label(body, text="--:--", bg=INK, fg="#7BD59A",
                                     font=OVERLAY_DIGIT_FONT)
            self.time_lbl.pack()
            self.note_lbl = tk.Label(body, text="남은 시간", bg=INK,
                                     fg="#8FA69A", font=("Malgun Gothic", 8))
            self.note_lbl.pack()
            for w in (self, body, self.head_lbl, self.time_lbl, self.note_lbl):
                w.bind("<ButtonPress-1>", self._grab)
                w.bind("<B1-Motion>", self._drag_to)
                w.bind("<ButtonRelease-1>", self._drop)
            self.place_self(monitor_index, saved)
            for show in (self.deiconify, self.lift,
                         lambda: self.attributes("-topmost", True)):
                try:
                    show()
                except Exception:
                    pass
            TimerOverlay._live.append(self)

        def _size(self):
            w = h = 0
            try:
                self.update_idletasks()
                w, h = self.winfo_reqwidth(), self.winfo_reqheight()
            except Exception:
                pass
            return (max(int(w), OVERLAY_SIZE[0]), max(int(h), OVERLAY_SIZE[1]))

        def place_self(self, monitor_index=None, saved=None):
            size = self._size()
            pos = None
            try:
                if self.monitors:
                    pos = overlay_position(self.monitors, monitor_index,
                                           size=size, saved=saved)
                else:
                    pos = overlay_position_for_screen(
                        self.winfo_screenwidth(), self.winfo_screenheight(),
                        size=size, saved=saved)
            except Exception as e:
                log_error("타이머 오버레이 위치 계산", e)
            if pos is None:
                return None
            try:
                self.geometry(f"+{int(pos[0])}+{int(pos[1])}")
            except Exception:
                return None
            return (int(pos[0]), int(pos[1]))

        def _grab(self, event):
            try:
                self._drag = (event.x_root - self.winfo_x(),
                              event.y_root - self.winfo_y())
            except Exception:
                self._drag = None
            return "break"

        def _drag_to(self, event):
            if not self._drag:
                return None
            self._moved = True
            try:
                self.geometry(f"+{int(event.x_root - self._drag[0])}"
                              f"+{int(event.y_root - self._drag[1])}")
            except Exception:
                pass
            return "break"

        def _drop(self, event=None):
            self._drag = None
            if not self._moved:
                return None
            self._moved = False
            try:
                pos = clamp_overlay_position(self.winfo_x(), self.winfo_y(),
                                             self.monitors, size=self._size())
                self.geometry(f"+{pos[0]}+{pos[1]}")
                save_overlay_pos(pos[0], pos[1])
                return pos
            except Exception as e:
                log_error("타이머 오버레이 위치 저장", e)
            return None

        def update_time(self, text, color=None, note=None):
            try:
                self.time_lbl.configure(text=text,
                                        **({"fg": color} if color else {}))
                if note is not None:
                    self.note_lbl.configure(text=note)
            except Exception:
                pass
            return text

        def close(self):
            try:
                TimerOverlay._live.remove(self)
            except ValueError:
                pass
            except Exception:
                pass
            try:
                if self.winfo_exists():
                    self.destroy()
            except Exception:
                pass
            return True

        @classmethod
        def close_all(cls):
            gone = 0
            for ov in list(cls._live):
                try:
                    ov.close()
                    gone += 1
                except Exception:
                    pass
            cls._live = []
            return gone

    class TimerWindow(tk.Toplevel):

        ALERTS = (600, 300, 60)

        def __init__(self, app, exam):
            super().__init__(app)
            self.app = app
            self.exam = exam
            self.total_seconds = exam["minutes"] * 60
            self.remaining = self.total_seconds
            self.paused = False
            self.finished = False
            self.practice = False
            self.started_mono = time.monotonic()
            self.paused_accum = 0.0
            self._pause_started = None
            self._alerted = set()
            self._after_id = None

            self.title(f"{APP_TITLE} - 타이머")
            self.configure(bg=INK)
            self.attributes("-topmost", True)
            self.resizable(False, False)
            self.protocol("WM_DELETE_WINDOW", self._on_close)

            practice = exam.get("practice_info")
            head = exam["set"]["name"]
            if practice and practice.get("mode") == "오답재풀이":
                head += " · 오답 재풀이(" + ", ".join(
                    practice.get("sheets") or []) + ")"
            elif practice:
                head += f" · 부분 연습({practice['label']})"
            tk.Label(self, text=head, bg=INK,
                     fg="#F2B24C" if practice else "#A7C8B2",
                     font=("Malgun Gothic", 9)).pack(padx=16, pady=(10, 0))
            self.time_lbl = tk.Label(self, text=self._fmt(), bg=INK,
                                     fg="#7BD59A", font=DIGIT_FONT)
            self.time_lbl.pack(padx=24, pady=(0, 2))
            self.status_lbl = tk.Label(
                self,
                text=(("오답 재풀이 — 오답 시트만 채점"
                       if practice.get("mode") == "오답재풀이"
                       else f"부분 연습 — {practice['label']}")
                      if practice else "시험 진행 중"),
                bg=INK, fg="#A7C8B2", font=("Malgun Gothic", 9))
            self.status_lbl.pack()
            if not exam["set"].get("pdf"):
                tk.Label(self, text="문제지 미연결", bg=INK, fg="#8FA69A",
                         font=("Malgun Gothic", 8)).pack()
            btns = tk.Frame(self, bg=INK)
            btns.pack(pady=(6, 12))
            self.pause_btn = tk.Button(btns, text="일시정지", width=9,
                                       font=UI_FONT, relief="flat",
                                       bg="#2C4A38", fg="white",
                                       activebackground="#3A5F49",
                                       command=self.toggle_pause)
            self.pause_btn.pack(side="left", padx=5)
            tk.Button(btns, text="제출", width=9, font=UI_FONT_BOLD,
                      relief="flat", bg=BRAND, fg="white",
                      activebackground=BRAND_DARK,
                      command=self.submit).pack(side="left", padx=5)
            self.reopen_btn = tk.Button(
                btns, text="풀이 파일 열기", width=12, font=UI_FONT,
                relief="flat", bg="#2C4A38", fg="white",
                activebackground="#3A5F49", command=self.reopen_file)
            self.reopen_btn.pack(side="left", padx=5)
            self.overlay = None
            self._open_overlay(head)
            self._tick()

        def _open_overlay(self, head):
            self.overlay = None
            try:
                if not timer_overlay_enabled():
                    return None
                layout = (self.exam or {}).get("monitor_layout") or {}
                mons = layout.get("monitors") or list_monitors()
                index = layout.get("timer")
                if index is None:
                    index = resolve_monitor_assignment(mons).get("timer")
                self.overlay = TimerOverlay(
                    self, head=head, monitor_index=index, monitors=mons,
                    saved=saved_overlay_pos())
                self._sync_overlay()
            except Exception as e:
                log_error("타이머 오버레이", e)
                self.overlay = None
            return self.overlay

        def _sync_overlay(self):
            ov = getattr(self, "overlay", None)
            if ov is None:
                return None
            try:
                if not ov.winfo_exists():
                    self.overlay = None
                    return None
                if self.finished:
                    note, color = "시험 시간 종료", "#FF8B80"
                elif self.paused:
                    note, color = "일시정지 (연습)", "#8FA69A"
                else:
                    note, color = "남은 시간", self._color()
                return ov.update_time(self._fmt(), color, note)
            except Exception:
                return None

        def close_overlay(self):
            ov = getattr(self, "overlay", None)
            self.overlay = None
            if ov is None:
                return False
            try:
                ov.close()
            except Exception:
                pass
            return True

        def destroy(self):
            self.close_overlay()
            return super().destroy()

        def reopen_file(self):
            self.app.reopen_student(self.exam)

        def _fmt(self):
            m, s = divmod(max(0, self.remaining), 60)
            return f"{m:02d}:{s:02d}"

        def _color(self):
            if self.remaining <= 300:
                return "#FF8B80"
            if self.remaining <= 600:
                return "#F2B24C"
            return "#7BD59A"

        def _tick(self):
            if self.finished:
                return
            if not self.paused:
                self.time_lbl.configure(text=self._fmt(), fg=self._color())
                self._sync_overlay()
                for mark in self.ALERTS:
                    if self.remaining == mark and mark not in self._alerted:
                        self._alerted.add(mark)
                        beep(self)
                if self.remaining <= 0:
                    self._time_up()
                    return
                self.remaining -= 1
            self.after(1000, self._tick)

        def elapsed_seconds(self):
            paused = self.paused_accum
            if self._pause_started is not None:
                paused += time.monotonic() - self._pause_started
            return time.monotonic() - self.started_mono - paused

        def toggle_pause(self):
            self.paused = not self.paused
            if self.paused:
                self._pause_started = time.monotonic()
                self.practice = True
                self.pause_btn.configure(text="재개")
                self.status_lbl.configure(text="연습 모드 (일시정지)")
                self.time_lbl.configure(fg="#8FA69A")
            else:
                if self._pause_started is not None:
                    self.paused_accum += time.monotonic() - self._pause_started
                    self._pause_started = None
                self.pause_btn.configure(text="일시정지")
                self.status_lbl.configure(
                    text="시험 진행 중 (연습 모드)" if self.practice
                    else "시험 진행 중")
                self.time_lbl.configure(fg=self._color())
            self._sync_overlay()

        def _time_up(self):
            self.finished = True
            self.remaining = 0
            self.time_lbl.configure(text="00:00", fg="#FF8B80")
            self.status_lbl.configure(text="시험 시간 종료")
            self._sync_overlay()
            beep(self)
            beep(self)
            self.submit(time_up=True)

        def _on_close(self):
            if self.finished:
                self.destroy()
                self.app.exam_closed()
                return
            if messagebox.askyesno(
                    APP_TITLE, "시험을 중단할까요?\n"
                    "(제출하지 않으면 채점되지 않습니다)", parent=self):
                self.finished = True
                self.destroy()
                self.app.exam_closed()

        def submit(self, time_up=False):
            self.finished = True
            head = "시험 시간이 종료되었습니다.\n\n" if time_up else ""
            ok = messagebox.askokcancel(
                f"{APP_TITLE} - 제출",
                head + "Excel에서 풀이 파일을 Ctrl+S로 저장했는지 확인하세요.\n"
                "저장하지 않으면 마지막 저장 상태로 채점됩니다.\n\n"
                "[확인]을 누르면 채점을 시작합니다.",
                parent=self)
            if not ok:
                if self.remaining > 0 and not time_up:
                    self.finished = False
                    self.after(1000, self._tick)
                else:
                    self.status_lbl.configure(
                        text="시험 시간 종료 - [제출]을 눌러 채점하세요")
                return
            elapsed = self.elapsed_seconds()
            self.status_lbl.configure(text="채점 중입니다...")
            ov = getattr(self, "overlay", None)
            if ov is not None:
                try:
                    ov.update_time(self._fmt(), "#A7C8B2", "채점 중")
                except Exception:
                    pass
            self.app.start_grading(self.exam, elapsed, self.practice,
                                   on_done=self._grading_done)

        def _grading_done(self):
            self.close_overlay()
            try:
                self.destroy()
            except Exception:
                pass
            self.app.exam_closed()

    class ExcelWarnWindow(tk.Toplevel):

        def __init__(self, app, exam, detail=""):
            super().__init__(app)
            self.app = app
            self.exam = exam
            self.title(f"{APP_TITLE} - Excel 확인")
            self.configure(bg=BG)
            self.attributes("-topmost", True)
            self.resizable(False, False)
            student = exam.get("student") or "(풀이 파일 없음)"
            frm = tk.Frame(self, bg=BG, padx=16, pady=14)
            frm.pack(fill="both", expand=True)
            tk.Label(frm, text="Excel 창이 열리지 않은 것 같습니다",
                     bg=BG, fg=RED, font=("Malgun Gothic", 12, "bold")).pack(
                anchor="w")
            tk.Label(frm, text=(
                "풀이 파일이 Excel에서 열려 있으면 이 창은 그냥 닫으면 됩니다.\n"
                "Excel이 깜빡하고 사라졌거나 아예 안 떴다면 아래 버튼으로 "
                "다시 열어 보세요.\n(작업 표시줄의 Excel 아이콘·PDF 뷰어 뒤에 "
                "숨은 창도 확인해 주세요)"),
                bg=BG, fg=INK, font=UI_FONT, justify="left",
                wraplength=540).pack(anchor="w", pady=(6, 0))
            tk.Label(frm, text=f"파일: {student}", bg=BG, fg=SUB,
                     font=("Malgun Gothic", 9), justify="left",
                     wraplength=540).pack(anchor="w", pady=(6, 0))
            self.detail_lbl = tk.Label(frm, text="", bg=BG, fg=SUB,
                                       font=("Malgun Gothic", 9),
                                       justify="left", wraplength=540)
            self.detail_lbl.pack(anchor="w")
            self.update_detail(detail)
            btns = tk.Frame(frm, bg=BG)
            btns.pack(fill="x", pady=(12, 0))
            self.buttons = []
            for text, cmd, primary in (
                    ("다시 열기(Excel 직접)", self.reopen_excel, True),
                    ("기본 프로그램으로 열기", self.reopen_default, False),
                    ("폴더 열기", self.open_folder, False)):
                b = tk.Button(
                    btns, text=text, command=cmd, padx=10, pady=4,
                    font=UI_FONT_BOLD if primary else UI_FONT,
                    relief="flat" if primary else "groove",
                    bg=BRAND if primary else "SystemButtonFace"
                    if sys.platform == "win32" else CARD,
                    fg="white" if primary else INK,
                    activebackground=BRAND_DARK if primary else BRAND_SOFT)
                b.pack(side="left", padx=(0, 6))
                self.buttons.append(b)
            tk.Button(btns, text="닫기", font=UI_FONT, relief="groove",
                      padx=10, pady=4, command=self.destroy).pack(side="right")
            self.status_lbl = tk.Label(frm, text="", bg=BG, fg=BRAND_DARK,
                                       font=("Malgun Gothic", 9), anchor="w",
                                       wraplength=540, justify="left")
            self.status_lbl.pack(fill="x", pady=(8, 0))

        def update_detail(self, detail):
            self.detail_lbl.configure(
                text=f"확인 결과: {detail}" if detail else "")

        def _status(self, text):
            try:
                self.status_lbl.configure(text=text)
            except Exception:
                pass

        def reopen_excel(self):
            ok, method, err = self.app.reopen_student(self.exam,
                                                      prefer_excel=True)
            self._status(("다시 열었습니다: " + method + " — 5초 뒤 다시 확인합니다")
                         if ok else f"열지 못했습니다: {err}")

        def reopen_default(self):
            ok, method, err = self.app.reopen_student(self.exam,
                                                      prefer_excel=False)
            self._status(("기본 프로그램으로 열었습니다 — 5초 뒤 다시 확인합니다")
                         if ok else f"열지 못했습니다: {err}")

        def open_folder(self):
            student = self.exam.get("student") or ""
            folder = os.path.dirname(os.path.abspath(student)) if student \
                else (self.exam.get("set") or {}).get("dir") or BASE_DIR
            ok, err = open_file(folder)
            startup_log(f"폴더 열기: {'성공' if ok else '실패 ' + err} · {folder}")
            self._status(f"폴더를 열었습니다: {folder}" if ok
                         else f"폴더를 열지 못했습니다: {err}")

    class ReviewWindow(tk.Toplevel):

        def __init__(self, app, set_info, json_path, items, note=None):
            super().__init__(app)
            self.app = app
            self.set_info = set_info
            self.json_name = os.path.basename(json_path)
            self.items = items
            self.checks = load_review_state(set_info.get("norm") or "",
                                            self.json_name)
            self.checks &= set(range(len(items)))
            self.title(f"{APP_TITLE} - 오답노트: {set_info['name']}")
            self.configure(bg=BG)
            self.minsize(760, 460)

            top = tk.Frame(self, bg=BG, padx=12, pady=8)
            top.pack(fill="x")
            tk.Label(top, text=f"{set_info['name']} 오답노트", bg=BG, fg=INK,
                     font=UI_FONT_BOLD).pack(side="left")
            self.progress_lbl = tk.Label(top, text="", bg=BG, fg=SUB,
                                         font=UI_FONT)
            self.progress_lbl.pack(side="right")
            if note:
                tk.Label(self, text=note, bg=BG, fg="#B45309",
                         font=("Malgun Gothic", 9)).pack(fill="x", padx=12)

            body = tk.Frame(self, bg=BG, padx=12, pady=6)
            body.pack(fill="both", expand=True)
            body.columnconfigure(1, weight=1)
            body.rowconfigure(0, weight=1)
            listfrm = tk.Frame(body, bg=CARD, highlightbackground=LINE,
                               highlightthickness=1)
            listfrm.grid(row=0, column=0, sticky="nsw", padx=(0, 10))
            self.listbox = tk.Listbox(listfrm, font=UI_FONT, width=34, bd=0,
                                      highlightthickness=0, bg=CARD, fg=INK,
                                      selectbackground=BRAND_SOFT,
                                      selectforeground=BRAND_DARK,
                                      activestyle="none",
                                      exportselection=False)
            sb = tk.Scrollbar(listfrm, command=self.listbox.yview)
            self.listbox.configure(yscrollcommand=sb.set)
            self.listbox.pack(side="left", fill="both", expand=True,
                              padx=(6, 0), pady=6)
            sb.pack(side="right", fill="y")
            self.listbox.bind("<<ListboxSelect>>",
                              lambda e: self._show_detail())
            self.detail = tk.Text(body, font=("Malgun Gothic", 10),
                                  wrap="word", bg=CARD, fg=INK,
                                  relief="flat", padx=12, pady=10,
                                  state="disabled")
            self.detail.grid(row=0, column=1, sticky="nsew")

            btns = tk.Frame(self, bg=BG, padx=12, pady=10)
            btns.pack(fill="x")
            self.check_btn = tk.Button(
                btns, text="이해했음 체크", font=UI_FONT_BOLD, bg=BRAND,
                fg="white", activebackground=BRAND_DARK, relief="flat",
                padx=14, pady=4, command=self.toggle_check)
            self.check_btn.pack(side="left")
            self.done_lbl = tk.Label(btns, text="", bg=BG, fg=BRAND,
                                     font=UI_FONT_BOLD)
            self.done_lbl.pack(side="left", padx=12)
            tk.Button(btns, text="닫기", font=UI_FONT, relief="groove",
                      padx=12, pady=4, command=self.destroy).pack(
                side="right")
            self.retake_btn = tk.Button(
                btns, text="재응시 시작", font=UI_FONT_BOLD, bg=BRAND,
                fg="white", activebackground=BRAND_DARK, relief="flat",
                padx=14, pady=4, state="disabled", command=self._retake)
            self.retake_btn.pack(side="right", padx=8)

            self._refresh_list()
            if items:
                self.listbox.selection_set(0)
                self._show_detail()

        def _refresh_list(self):
            sel = self.listbox.curselection()
            self.listbox.delete(0, "end")
            for i, it in enumerate(self.items):
                mark = "[v]" if i in self.checks else "[  ]"
                self.listbox.insert(
                    "end", f" {mark} {it.get('label', '?')} "
                           f"(-{it.get('lost', 0):g}점)")
            if sel:
                self.listbox.selection_set(sel[0])
            done = len(self.checks)
            total = len(self.items)
            self.progress_lbl.configure(text=f"이해함 {done} / {total}")
            if total and done == total:
                self.done_lbl.configure(
                    text="모든 항목 이해 완료 — 이제 시험 모드로 재응시해 "
                         "보세요")
                self.retake_btn.configure(state="normal")
            else:
                self.done_lbl.configure(text="")
                self.retake_btn.configure(state="disabled")

        def _show_detail(self):
            sel = self.listbox.curselection()
            if not sel or sel[0] >= len(self.items):
                return
            text = review_item_text(self.items[sel[0]])
            self.detail.configure(state="normal")
            self.detail.delete("1.0", "end")
            self.detail.insert("1.0", text)
            self.detail.configure(state="disabled")

        def toggle_check(self):
            sel = self.listbox.curselection()
            if not sel:
                return
            i = sel[0]
            if i in self.checks:
                self.checks.discard(i)
            else:
                self.checks.add(i)
            record_review_state(self.set_info.get("norm") or "",
                                self.json_name, self.checks,
                                total=len(self.items))
            self._refresh_list()

        def _retake(self):
            self.destroy()
            self.app.start_exam()

    class PracticeDialog(tk.Toplevel):

        def __init__(self, app):
            super().__init__(app)
            self.result = None
            self.title(f"{APP_TITLE} - 부분 연습")
            self.configure(bg=BG)
            self.resizable(False, False)
            frm = tk.Frame(self, bg=BG, padx=18, pady=14)
            frm.pack(fill="both", expand=True)
            tk.Label(frm, text="연습할 영역을 선택하세요", bg=BG, fg=INK,
                     font=UI_FONT_BOLD).pack(anchor="w")
            tk.Label(frm, text="루틴 주차 프리셋:", bg=BG, fg=SUB,
                     font=("Malgun Gothic", 9)).pack(anchor="w", pady=(8, 2))
            pf = tk.Frame(frm, bg=BG)
            pf.pack(fill="x")
            for label, areas in PRACTICE_PRESETS:
                tk.Button(pf, text=label, font=("Malgun Gothic", 8),
                          relief="groove", padx=5, pady=2,
                          command=lambda a=areas: self.apply_preset(a)
                          ).pack(side="left", padx=2)
            self.area_vars = {}
            af = tk.Frame(frm, bg=BG)
            af.pack(fill="x", pady=(8, 2))
            for name, sheet_list, minutes in PRACTICE_AREAS:
                v = keep_var(tk.BooleanVar(value=False))
                self.area_vars[name] = v
                tk.Checkbutton(
                    af, text=f"{name}  ({' · '.join(sheet_list)}, "
                             f"권장 {minutes}분)",
                    variable=v, bg=BG, fg=INK, anchor="w", font=UI_FONT,
                    activebackground=BG, command=self._sync_minutes
                ).pack(anchor="w")
            self.adv_open = False
            self.adv_btn = tk.Button(
                frm, text="개별 시트 선택 ▸", relief="flat", bg=BG, fg=SUB,
                font=("Malgun Gothic", 9), command=self._toggle_adv)
            self.adv_btn.pack(anchor="w", pady=(4, 0))
            self.adv_frame = tk.Frame(frm, bg=BG)
            self.sheet_vars = {}
            for _name, sheet_list, _m in PRACTICE_AREAS:
                for sh in sheet_list:
                    v = keep_var(tk.BooleanVar(value=False))
                    self.sheet_vars[sh] = v
                    tk.Checkbutton(self.adv_frame, text=sh, variable=v,
                                   bg=BG, fg=INK, activebackground=BG,
                                   font=("Malgun Gothic", 9)
                                   ).pack(anchor="w", padx=16)
            tf = tk.Frame(frm, bg=BG)
            tf.pack(fill="x", pady=(10, 0))
            tk.Label(tf, text="연습 시간(분):", bg=BG, fg=INK,
                     font=UI_FONT).pack(side="left")
            self.minutes_var = keep_var(tk.IntVar(value=DEFAULT_MINUTES))
            tk.Spinbox(tf, from_=MIN_MINUTES, to=MAX_MINUTES,
                       textvariable=self.minutes_var, width=4,
                       font=UI_FONT).pack(side="left", padx=6)
            bf = tk.Frame(frm, bg=BG)
            bf.pack(fill="x", pady=(12, 0))
            tk.Button(bf, text="연습 시작", font=UI_FONT_BOLD, bg=BRAND,
                      fg="white", activebackground=BRAND_DARK, relief="flat",
                      padx=16, pady=4, command=self._start).pack(side="left")
            tk.Button(bf, text="취소", font=UI_FONT, relief="groove",
                      padx=12, pady=4, command=self.destroy).pack(
                side="right")
            self.grab_set()

        def apply_preset(self, areas):
            for name, v in self.area_vars.items():
                v.set(name in areas)
            for v in self.sheet_vars.values():
                v.set(False)
            self._sync_minutes()

        def _toggle_adv(self):
            self.adv_open = not self.adv_open
            if self.adv_open:
                self.adv_frame.pack(fill="x")
                self.adv_btn.configure(text="개별 시트 선택 ▾")
            else:
                self.adv_frame.pack_forget()
                self.adv_btn.configure(text="개별 시트 선택 ▸")

        def _sync_minutes(self):
            areas = [n for n, v in self.area_vars.items() if v.get()]
            if areas:
                self.minutes_var.set(practice_minutes(areas))

        def selection(self):
            sheets = [sh for sh, v in self.sheet_vars.items() if v.get()]
            if sheets:
                label = "+".join(sheets) if len(sheets) <= 2 \
                    else f"개별{len(sheets)}시트"
                return sheets, label
            areas = [n for n, v in self.area_vars.items() if v.get()]
            return sheets_for_areas(areas), "+".join(areas)

        def _start(self):
            sheets, label = self.selection()
            if not sheets:
                messagebox.showinfo(APP_TITLE, "연습할 영역이나 시트를 "
                                    "선택하세요.", parent=self)
                return
            try:
                minutes = int(self.minutes_var.get())
            except Exception:
                minutes = DEFAULT_MINUTES
            self.result = {"sheets": sheets, "label": label,
                           "minutes": minutes}
            self.destroy()

    class DiagnosisWindow(tk.Toplevel):

        def __init__(self, app, text, title=None, hint=None):
            super().__init__(app)
            self.text_value = text
            self.title(title or f"{APP_TITLE} - 세트 인식 진단")
            self.configure(bg=BG)
            self.geometry("760x520")
            frm = tk.Frame(self, bg=BG, padx=14, pady=10)
            frm.pack(fill="both", expand=True)
            tk.Label(frm, text=hint or (
                "파일명 → 정규화 키 → 역할 → 세트 → 일정 슬롯 매칭 결과입니다. "
                "[복사]해서 채팅에 붙여넣으면 진단해 드립니다."),
                bg=BG, fg=INK, font=("Malgun Gothic", 9),
                wraplength=720, justify="left").pack(anchor="w")
            self.text = tk.Text(frm, font=("Consolas", 9), wrap="none",
                                bg=CARD, fg=INK)
            self.text.insert("1.0", text)
            self.text.configure(state="disabled")
            self.text.pack(fill="both", expand=True, pady=(6, 6))
            bf = tk.Frame(frm, bg=BG)
            bf.pack(fill="x")
            self.copy_btn = tk.Button(
                bf, text="클립보드 복사", font=UI_FONT_BOLD, bg=BRAND,
                fg="white", activebackground=BRAND_DARK, relief="flat",
                padx=12, pady=3, command=self.copy)
            self.copy_btn.pack(side="left")
            tk.Button(bf, text="닫기", font=UI_FONT, relief="groove",
                      padx=12, pady=3, command=self.destroy).pack(side="right")
            self.copy()

        def copy(self):
            try:
                self.clipboard_clear()
                self.clipboard_append(self.text_value)
                self.update_idletasks()
                self.copy_btn.configure(text="복사됨 — 채팅에 붙여넣으세요")
                return True
            except Exception:
                return False

    class UpdateConsentDialog(tk.Toplevel):

        def __init__(self, app, text=None):
            super().__init__(app)
            self.app = app
            self.agreed = False
            self.answered = False
            self.title(f"{APP_TITLE} - 자동 업데이트 동의")
            self.configure(bg=BG)
            self.resizable(False, False)
            frm = tk.Frame(self, bg=BG, padx=18, pady=14)
            frm.pack(fill="both", expand=True)
            tk.Label(frm, text="자동 업데이트를 쓸까요?", bg=BG, fg=INK,
                     font=("Malgun Gothic", 12, "bold")).pack(anchor="w")
            tk.Label(frm, text=text or first_run_notice_text(), bg=BG, fg=INK,
                     font=UI_FONT, justify="left", wraplength=520).pack(
                anchor="w", pady=(8, 12))
            bf = tk.Frame(frm, bg=BG)
            bf.pack(fill="x")
            self.yes_btn = tk.Button(
                bf, text=UPDATE_CONSENT_YES, font=UI_FONT_BOLD, bg=BRAND,
                fg="white", activebackground=BRAND_DARK, relief="flat",
                padx=16, pady=5, command=self.accept)
            self.yes_btn.pack(side="left")
            self.no_btn = tk.Button(
                bf, text=UPDATE_CONSENT_NO, font=UI_FONT, relief="groove",
                padx=14, pady=4, command=self.decline)
            self.no_btn.pack(side="right")
            try:
                self.protocol("WM_DELETE_WINDOW", self.decline)
                self.bind("<Escape>", lambda e: self.decline())
                self.transient(app)
                self.grab_set()
                self.yes_btn.focus_set()
            except Exception:
                pass

        def accept(self):
            self.agreed, self.answered = True, True
            self.destroy()
            return True

        def decline(self):
            self.agreed, self.answered = False, True
            self.destroy()
            return False

    class ExamDateDialog(tk.Toplevel):

        def __init__(self, app, on_save=None, first_run=False):
            super().__init__(app)
            self.app = app
            self.on_save = on_save
            self.saved = None
            self.first_run = first_run
            self.title(f"{APP_TITLE} - 시험일 설정")
            self.configure(bg=BG)
            self.resizable(False, False)
            self.dates = list(exam_dates())
            frm = tk.Frame(self, bg=BG, padx=16, pady=12)
            frm.pack(fill="both", expand=True)
            head = ("시험 날짜를 넣으세요. 그 날 기준으로 "
                    f"D-{ROUTINE_DAYS} ~ D-1 일주일 일정이 만들어집니다."
                    if first_run else
                    "시험일을 바꾸면 일정이 다시 계산됩니다.")
            tk.Label(frm, text=head, bg=BG, fg=INK, font=UI_FONT_BOLD,
                     wraplength=460, justify="left").pack(anchor="w")
            tk.Label(frm, text="여러 개를 넣을 수 있고 하나씩 지울 수도 "
                               "있습니다. 일정은 언제나 **가장 가까운 미래 "
                               "시험일** 기준이고, 지난 시험일은 목록에 남되 "
                               "기준에서 빠집니다. 시험일이 하나도 없으면 "
                               "일정 대신 자유 연습 모드로 돕니다.",
                     bg=BG, fg=SUB, font=("Malgun Gothic", 9),
                     wraplength=460, justify="left").pack(anchor="w",
                                                          pady=(2, 8))
            row = tk.Frame(frm, bg=BG)
            row.pack(fill="x", pady=3)
            tk.Label(row, text="추가할 날짜", bg=BG, fg=INK, font=UI_FONT,
                     width=11, anchor="w").pack(side="left")
            self.var = tk.StringVar(value=(self.dates[0].isoformat()
                                           if self.dates else
                                           date.today().isoformat()))
            ent = tk.Entry(row, textvariable=self.var, font=UI_FONT, width=13)
            ent.pack(side="left")
            for label, days in (("-1일", -1), ("+1일", 1), ("+7일", 7)):
                tk.Button(row, text=label, font=UI_FONT, relief="groove",
                          padx=6, command=lambda d=days: self._shift(d)
                          ).pack(side="left", padx=(5, 0))
            self.add_btn = tk.Button(
                row, text="추가", font=UI_FONT_BOLD, bg=BRAND, fg="white",
                activebackground=BRAND_DARK, relief="flat", padx=14, pady=2,
                command=self.add)
            self.add_btn.pack(side="left", padx=(8, 0))
            tk.Label(frm, text="예: 2026-11-21", bg=BG, fg=SUB,
                     font=("Malgun Gothic", 9)).pack(anchor="w", pady=(2, 6))
            tk.Label(frm, text="넣은 시험일", bg=BG, fg=INK,
                     font=UI_FONT_BOLD).pack(anchor="w")
            lf = tk.Frame(frm, bg=CARD, highlightbackground=LINE,
                          highlightthickness=1)
            lf.pack(fill="both", expand=True, pady=(3, 6))
            self.listbox = tk.Listbox(lf, font=UI_FONT, height=6, bd=0,
                                      highlightthickness=0, bg=CARD, fg=INK,
                                      selectbackground=BRAND_SOFT,
                                      selectforeground=BRAND_DARK,
                                      activestyle="none", exportselection=False)
            self.listbox.pack(fill="both", expand=True, padx=6, pady=5)
            bf2 = tk.Frame(frm, bg=BG)
            bf2.pack(fill="x")
            self.del_btn = tk.Button(bf2, text="선택 삭제", font=UI_FONT,
                                     relief="groove", padx=12, pady=3,
                                     command=self.remove_selected)
            self.del_btn.pack(side="left")
            self.clear_btn = tk.Button(bf2, text="전체 비우기", font=UI_FONT,
                                       relief="groove", padx=12, pady=3,
                                       command=self.clear_all)
            self.clear_btn.pack(side="left", padx=6)
            self.msg = tk.Label(frm, text="", bg=BG, fg=RED,
                                font=("Malgun Gothic", 9), wraplength=460,
                                justify="left")
            self.msg.pack(anchor="w", pady=(6, 0))
            bf = tk.Frame(frm, bg=BG)
            bf.pack(fill="x", pady=(10, 0))
            self.save_btn = tk.Button(
                bf, text="저장", font=UI_FONT_BOLD, bg=BRAND, fg="white",
                activebackground=BRAND_DARK, relief="flat", padx=16, pady=4,
                command=self.save)
            self.save_btn.pack(side="left")
            tk.Button(bf, text="나중에" if first_run else "취소", font=UI_FONT,
                      relief="groove", padx=12, pady=3,
                      command=self.destroy).pack(side="right")
            self._render()
            try:
                self.transient(app)
                self.grab_set()
            except Exception:
                pass

        def _shift(self, days):
            base = parse_iso_date(self.var.get()) or date.today()
            self.var.set((base + timedelta(days=days)).isoformat())

        def _render(self):
            self.dates = _norm_date_list(self.dates)
            self.listbox.delete(0, "end")
            today = date.today()
            act = effective_exams(self.dates, today)
            for d in self.dates:
                if d < today:
                    tag = "지남 (일정 기준에서 제외)"
                elif act and d == act[0]:
                    tag = f"기준 시험일 · {dday_label((d - today).days)}"
                else:
                    tag = dday_label((d - today).days)
                self.listbox.insert("end",
                                    f" {d.isoformat()} ({'월화수목금토일'[d.weekday()]})"
                                    f"  —  {tag}")
            if not self.dates:
                self.listbox.insert("end", " (시험일 없음 — 자유 연습 모드)")
            for b, st in ((self.del_btn, bool(self.dates)),
                          (self.clear_btn, bool(self.dates))):
                b.configure(state="normal" if st else "disabled")

        def values(self):
            return [d.isoformat() for d in _norm_date_list(self.dates)]

        def add(self, value=None):
            txt = str(value if value is not None else self.var.get()).strip()
            if not txt:
                self.msg.configure(text="추가할 날짜를 입력하세요 (YYYY-MM-DD).")
                return False
            d = parse_iso_date(txt)
            if d is None:
                self.msg.configure(text=f"날짜 형식이 올바르지 않습니다: {txt}")
                return False
            if d in self.dates:
                self.msg.configure(text=f"이미 들어 있는 시험일입니다: "
                                        f"{d.isoformat()}")
                return False
            self.dates.append(d)
            self.msg.configure(text="")
            self._render()
            return True

        def remove(self, value):
            d = parse_iso_date(value)
            if d is None or d not in self.dates:
                return False
            self.dates = [x for x in self.dates if x != d]
            self.msg.configure(text="")
            self._render()
            return True

        def remove_selected(self):
            sel = self.listbox.curselection()
            if not sel or sel[0] >= len(self.dates):
                self.msg.configure(text="목록에서 지울 시험일을 먼저 고르세요.")
                return False
            return self.remove(self.dates[sel[0]])

        def clear_all(self):
            self.dates = []
            self.msg.configure(text="")
            self._render()
            return True

        def validate(self):
            return _norm_date_list(self.dates), ""

        def save(self):
            dates, err = self.validate()
            if err:
                self.msg.configure(text=err)
                return False
            self.saved = save_exam_date_list(dates)
            if callable(self.on_save):
                try:
                    self.on_save(self.saved)
                except Exception as e:
                    log_error("시험일 저장 후 갱신", e)
            self.destroy()
            return True

    class SettingsDialog(tk.Toplevel):

        MONITOR_ROWS = (("pdf", "문제지(PDF)를 띄울 화면"),
                        ("excel", "Excel(풀이 파일)을 띄울 화면"),
                        ("timer", "타이머를 띄울 화면"))

        def __init__(self, app):
            super().__init__(app)
            self.app = app
            self.rows = []
            self.mon_vars = {}
            self.mon_msg = None
            self.overlay_msg = None
            try:
                self.monitors = list_monitors()
            except Exception as e:
                log_error("설정 창 모니터 목록", e)
                self.monitors = []
            self.title(f"{APP_TITLE} - 설정")
            self.configure(bg=BG)
            self.geometry(f"660x{self._window_height()}")
            self.minsize(540, 460)
            frm = tk.Frame(self, bg=BG, padx=16, pady=12)
            frm.pack(fill="both", expand=True)

            tk.Label(frm, text="세트 스캔 폴더", bg=BG, fg=INK,
                     font=UI_FONT_BOLD).pack(anchor="w")
            tk.Label(frm, text="이 폴더와 하위 폴더에서 '<이름>_문제.xlsx' + "
                               "'<이름>_정답.xlsx' 짝을 찾습니다. 기본값은 "
                               "시험장.py 가 있는 폴더의 상위 폴더라, 기존 "
                               "공부 폴더 근처에 풀면 갖고 있던 세트가 함께 "
                               "잡힙니다 — 원치 않으면 폴더를 좁히거나 세트를 "
                               "개별로 숨기세요.",
                     bg=BG, fg=SUB, font=("Malgun Gothic", 9), wraplength=580,
                     justify="left").pack(anchor="w", pady=(2, 4))
            self.root_lbl = tk.Label(frm, text="", bg=CARD, fg=INK,
                                     font=("Consolas", 9), justify="left",
                                     anchor="w", padx=8, pady=6,
                                     wraplength=580,
                                     highlightbackground=LINE,
                                     highlightthickness=1)
            self.root_lbl.pack(fill="x")
            rf = tk.Frame(frm, bg=BG)
            rf.pack(fill="x", pady=(4, 10))
            tk.Button(rf, text="폴더 선택", font=UI_FONT_BOLD, bg=BRAND,
                      fg="white", activebackground=BRAND_DARK, relief="flat",
                      padx=14, pady=3, command=self.choose_root).pack(side="left")
            tk.Button(rf, text="기본값으로", font=UI_FONT, relief="groove",
                      padx=12, pady=3, command=self.reset_root).pack(
                side="left", padx=6)
            tk.Button(rf, text="세트 다시 스캔", font=UI_FONT, relief="groove",
                      padx=12, pady=3, command=self.app.refresh_all).pack(
                side="left", padx=6)

            tk.Label(frm, text="세트 숨기기", bg=BG, fg=INK,
                     font=UI_FONT_BOLD).pack(anchor="w")
            tk.Label(frm, text="숨긴 세트는 세트 목록 · 오늘의 학습 일정 · "
                               "점수대 보드에서 빠집니다. 파일과 응시 기록은 "
                               "그대로 두므로 다시 보이게 하면 점수도 그대로 "
                               "돌아옵니다.",
                     bg=BG, fg=SUB, font=("Malgun Gothic", 9), wraplength=580,
                     justify="left").pack(anchor="w", pady=(2, 4))
            lf = tk.Frame(frm, bg=CARD, highlightbackground=LINE,
                          highlightthickness=1)
            lf.pack(fill="both", expand=True)
            self.listbox = tk.Listbox(lf, font=UI_FONT, bd=0,
                                      highlightthickness=0, bg=CARD, fg=INK,
                                      selectbackground=BRAND_SOFT,
                                      selectforeground=BRAND_DARK,
                                      activestyle="none", exportselection=False)
            sb = tk.Scrollbar(lf, command=self.listbox.yview)
            self.listbox.configure(yscrollcommand=sb.set)
            self.listbox.pack(side="left", fill="both", expand=True,
                              padx=(6, 0), pady=6)
            sb.pack(side="right", fill="y")
            self.listbox.bind("<Double-Button-1>", lambda e: self.toggle())
            hf = tk.Frame(frm, bg=BG)
            hf.pack(fill="x", pady=(4, 10))
            tk.Button(hf, text="숨기기 / 다시 보이기", font=UI_FONT_BOLD,
                      relief="groove", padx=12, pady=3,
                      command=self.toggle).pack(side="left")
            tk.Button(hf, text="전부 다시 보이기", font=UI_FONT,
                      relief="groove", padx=12, pady=3,
                      command=self.unhide_all).pack(side="left", padx=6)
            self.hide_msg = tk.Label(hf, text="", bg=BG, fg=SUB,
                                     font=("Malgun Gothic", 9))
            self.hide_msg.pack(side="left", padx=8)

            if monitor_ui_visible(self.monitors):
                self._build_monitor_section(frm)

            tf = tk.Frame(frm, bg=BG)
            tf.pack(fill="x")
            self.overlay_var = tk.BooleanVar(value=timer_overlay_enabled())
            tk.Checkbutton(tf, text="시험 중 화면 우측 상단에 남은 시간 표시 "
                                    "(끌어서 옮길 수 있음)",
                           variable=self.overlay_var, bg=BG, fg=INK,
                           activebackground=BG, font=UI_FONT,
                           command=self._toggle_overlay).pack(anchor="w")
            self.auto_update_var = tk.BooleanVar(value=auto_update_enabled())
            tk.Checkbutton(tf, text="자동 업데이트 (실행할 때마다 새 버전 적용)",
                           variable=self.auto_update_var, bg=BG, fg=INK,
                           activebackground=BG, font=UI_FONT,
                           command=self._toggle_auto_update).pack(anchor="w")
            self.auto_open_var = tk.BooleanVar(
                value=bool(get_app_setting(ROUTINE_AUTO_OPEN_SETTING, False)))
            tk.Checkbutton(tf, text="시작할 때 루틴 웹페이지 자동 열기",
                           variable=self.auto_open_var, bg=BG, fg=INK,
                           activebackground=BG, font=UI_FONT,
                           command=self._toggle_auto_open).pack(anchor="w")
            of = tk.Frame(frm, bg=BG)
            of.pack(fill="x", pady=(2, 0))
            tk.Button(of, text="타이머 위치 초기화", font=UI_FONT,
                      relief="groove", padx=12, pady=2,
                      command=self.reset_overlay_pos).pack(side="left")
            self.overlay_msg = tk.Label(of, text="", bg=BG, fg=SUB,
                                        font=("Malgun Gothic", 9))
            self.overlay_msg.pack(side="left", padx=8)
            tk.Label(frm, text="자동 업데이트를 꺼도 [업데이트 확인]으로 언제든 "
                               "직접 받을 수 있습니다.",
                     bg=BG, fg=SUB, font=("Malgun Gothic", 9),
                     wraplength=600, justify="left").pack(anchor="w",
                                                          pady=(2, 8))
            tk.Button(frm, text="닫기", font=UI_FONT, relief="groove",
                      padx=16, pady=4, command=self.destroy).pack(anchor="e")
            self.refresh_root_label()
            self.refresh_sets_list()
            try:
                self.transient(app)
            except Exception:
                pass

        def _window_height(self):
            h = 760 if monitor_ui_visible(self.monitors) else 620
            try:
                h = min(h, max(460, int(self.winfo_screenheight()) - 80))
            except Exception:
                pass
            return int(h)

        def _build_monitor_section(self, frm):
            tk.Label(frm, text="모니터 배치", bg=BG, fg=INK,
                     font=UI_FONT_BOLD).pack(anchor="w")
            tk.Label(frm, text="실제 시험장처럼 한 화면에는 문제지, 다른 화면에는 "
                               "Excel 을 띄웁니다. [시험 시작] 때 창을 자동으로 "
                               "옮기며, 문제지 뷰어(Edge·Acrobat·Chrome 등)는 "
                               "프로그램마다 창이 달라 못 옮길 수도 있습니다 "
                               "— 그럴 때는 안내만 나오고 시험은 그대로 "
                               "진행됩니다.",
                     bg=BG, fg=SUB, font=("Malgun Gothic", 9), wraplength=600,
                     justify="left").pack(anchor="w", pady=(2, 4))
            cur = resolve_monitor_assignment(self.monitors)
            labels = [m["label"] for m in self.monitors]
            grid = tk.Frame(frm, bg=BG)
            grid.pack(fill="x", pady=(0, 4))
            for r, (key, text) in enumerate(self.MONITOR_ROWS):
                tk.Label(grid, text=text, bg=BG, fg=INK, font=UI_FONT,
                         anchor="w", width=22).grid(row=r, column=0,
                                                    sticky="w", pady=2)
                i = cur.get(key) or 0
                var = tk.StringVar(value=labels[i] if 0 <= i < len(labels)
                                   else labels[0])
                self.mon_vars[key] = var
                box = ttk.Combobox(grid, textvariable=var, values=labels,
                                   state="readonly", font=UI_FONT, width=36)
                box.grid(row=r, column=1, sticky="we", padx=(6, 0), pady=2)
                box.bind("<<ComboboxSelected>>",
                         lambda _e: self.apply_monitors())
            grid.columnconfigure(1, weight=1)
            mb = tk.Frame(frm, bg=BG)
            mb.pack(fill="x", pady=(0, 10))
            tk.Button(mb, text="기본값으로", font=UI_FONT, relief="groove",
                      padx=12, pady=3,
                      command=self.reset_monitors).pack(side="left")
            self.mon_msg = tk.Label(
                mb, text=f"모니터 {len(self.monitors)}대 인식됨", bg=BG, fg=SUB,
                font=("Malgun Gothic", 9))
            self.mon_msg.pack(side="left", padx=8)

        def _monitor_choice(self, key):
            var = self.mon_vars.get(key)
            if var is None:
                return None
            label = var.get()
            for m in self.monitors:
                if m.get("label") == label:
                    return int(m["index"])
            return None

        def apply_monitors(self):
            saved = save_monitor_assignment(
                **{k: self._monitor_choice(k) for k, _t in self.MONITOR_ROWS})
            startup_log(f"모니터 배치 저장: 문제지={saved.get('pdf')} · "
                        f"Excel={saved.get('excel')} · "
                        f"타이머={saved.get('timer')}")
            if self.mon_msg is not None:
                self.mon_msg.configure(text="다음 [시험 시작]부터 적용됩니다")
            return saved

        def reset_monitors(self):
            clear_monitor_assignment()
            cur = resolve_monitor_assignment(self.monitors)
            for key, _t in self.MONITOR_ROWS:
                var = self.mon_vars.get(key)
                i = cur.get(key) or 0
                if var is not None and 0 <= i < len(self.monitors):
                    var.set(self.monitors[i]["label"])
            if self.mon_msg is not None:
                self.mon_msg.configure(text="기본값으로 되돌렸습니다")
            return cur

        def _toggle_overlay(self):
            on = set_timer_overlay_enabled(self.overlay_var.get())
            startup_log(f"우측 상단 타이머 {'켜짐' if on else '꺼짐'} "
                        f"(세트설정 _설정.{TIMER_OVERLAY_SETTING})")
            if self.overlay_msg is not None:
                self.overlay_msg.configure(
                    text="" if on else "타이머 창에서만 시간을 봅니다")
            return on

        def reset_overlay_pos(self):
            save_overlay_pos(None, None)
            if self.overlay_msg is not None:
                self.overlay_msg.configure(text="다음 응시부터 우측 상단")
            return None

        def refresh_root_label(self):
            cur = self.app.scan_root
            tag = ("기본값 (시험장.py 상위 폴더)" if os.path.abspath(cur) ==
                   os.path.abspath(default_scan_root()) else "직접 지정")
            extra = ""
            if getattr(self.app, "scan_root_fixed", False):
                extra = ("\n※ 이번 실행은 --scan-root 로 지정된 폴더를 씁니다 "
                         "— 저장한 값은 다음 실행부터 적용됩니다.")
            self.root_lbl.configure(text=f"{cur}\n({tag})" + extra)

        def choose_root(self):
            folder = filedialog.askdirectory(
                parent=self, title="세트를 찾을 폴더 선택",
                initialdir=self.app.scan_root or default_scan_root())
            if not folder:
                return None
            return self.apply_root(folder)

        def apply_root(self, folder):
            applied = save_scan_root(folder)
            self.app.scan_root = applied
            self.app.scan_root_fixed = False
            startup_log(f"스캔 폴더 변경: {applied} "
                        f"(세트설정 _설정.{SCAN_ROOT_SETTING})")
            self.refresh_root_label()
            self.app.refresh_all()
            self.app.show_toast(f"세트 스캔 폴더를 바꿨습니다 — {applied} "
                                "(다음 실행에도 유지됩니다)", seconds=8)
            return applied

        def reset_root(self):
            applied = save_scan_root(None)
            self.app.scan_root = applied
            self.app.scan_root_fixed = False
            startup_log(f"스캔 폴더 기본값 복귀: {applied}")
            self.refresh_root_label()
            self.app.refresh_all()
            self.app.show_toast(f"세트 스캔 폴더를 기본값으로 되돌렸습니다 — "
                                f"{applied}", seconds=8)
            return applied

        def refresh_sets_list(self):
            hidden = hidden_set_keys()
            rows = list(self.app.sets) + list(
                getattr(self.app, "hidden_sets", []) or [])
            rows.sort(key=lambda s: str(s.get("name") or ""))
            self.rows = rows
            sel = self.listbox.curselection()
            self.listbox.delete(0, "end")
            for s in rows:
                mark = "[숨김] " if is_hidden_set(s, hidden) else "        "
                self.listbox.insert("end", f" {mark}{s.get('name') or '?'}")
            if not rows:
                self.listbox.insert("end", " (인식된 세트가 없습니다)")
            if sel and sel[0] < len(rows):
                self.listbox.selection_set(sel[0])
            self.hide_msg.configure(
                text=f"숨긴 세트 {len(hidden)}개" if hidden else "숨긴 세트 없음")
            self.refresh_root_label()

        def _selected(self):
            sel = self.listbox.curselection()
            if not sel or sel[0] >= len(self.rows):
                return None
            return self.rows[sel[0]]

        def toggle(self):
            s = self._selected()
            if s is None:
                self.hide_msg.configure(text="목록에서 세트를 먼저 고르세요.")
                return None
            hidden = is_hidden_set(s)
            for k in set_hidden_keys(s):
                set_set_hidden(k, not hidden)
            startup_log(f"세트 {'다시 보이기' if hidden else '숨기기'}: "
                        f"{s.get('name')} ({', '.join(set_hidden_keys(s))})")
            self.app.refresh_all()
            self.app.show_toast(
                f"'{s.get('name')}' 세트를 다시 목록에 넣었습니다"
                if hidden else
                f"'{s.get('name')}' 세트를 숨겼습니다 — 파일과 기록은 그대로 "
                "두고 목록·일정·점수대 보드에서만 뺍니다", seconds=7)
            return not hidden

        def unhide_all(self):
            clear_hidden_sets()
            startup_log("숨긴 세트 전부 해제")
            self.app.refresh_all()
            self.app.show_toast("숨긴 세트를 전부 다시 보이게 했습니다",
                                seconds=6)
            return True

        def _toggle_auto_update(self):
            on = bool(self.auto_update_var.get())
            try:
                self.app.auto_update_var.set(on)
            except Exception:
                pass
            return self.app._toggle_auto_update()

        def _toggle_auto_open(self):
            on = bool(self.auto_open_var.get())
            try:
                self.app.routine_auto_open_var.set(on)
            except Exception:
                pass
            return self.app._toggle_routine_auto_open()

    class ScoreBoardWindow(tk.Toplevel):

        BAND_COLORS = {
            "none": (BG, SUB), "fail": ("#FDECEA", RED),
            "pass": (BRAND_SOFT, BRAND_DARK), "skilled": (BRAND_SOFT, BRAND_DARK),
            "high": (BRAND_SOFT, BRAND_DARK), "perfect": (BRAND, "white"),
        }

        def __init__(self, app):
            super().__init__(app)
            self.app = app
            self.rows = []
            self.selected = None
            self.title(f"{APP_TITLE} - 점수대 보드")
            self.configure(bg=BG)
            self.geometry("620x560")
            self.minsize(420, 380)
            frm = tk.Frame(self, bg=BG, padx=14, pady=10)
            frm.pack(fill="both", expand=True)
            tk.Label(frm, text="점수대 세트 보드", bg=BG, fg=INK,
                     font=UI_FONT_BOLD, anchor="w").pack(fill="x")
            self.sum_lbl = tk.Label(
                frm, text="", bg=BG, fg=SUB, font=("Malgun Gothic", 9),
                anchor="w", justify="left", wraplength=580)
            self.sum_lbl.pack(fill="x", pady=(2, 6))
            wrap = tk.Frame(frm, bg=CARD, highlightbackground=LINE,
                            highlightthickness=1)
            wrap.pack(fill="both", expand=True)
            self.canvas = tk.Canvas(wrap, bg=CARD, bd=0, highlightthickness=0)
            sb = tk.Scrollbar(wrap, command=self.canvas.yview)
            self.canvas.configure(yscrollcommand=sb.set)
            self.canvas.pack(side="left", fill="both", expand=True)
            sb.pack(side="right", fill="y")
            self.inner = tk.Frame(self.canvas, bg=CARD)
            self._win = self.canvas.create_window((0, 0), window=self.inner,
                                                  anchor="nw")
            self.inner.bind("<Configure>", lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")))
            self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(
                self._win, width=e.width))
            bf = tk.Frame(frm, bg=BG)
            bf.pack(fill="x", pady=(8, 0))
            self.start_btn = tk.Button(
                bf, text="이 세트 응시", font=UI_FONT_BOLD, bg=BRAND,
                fg="white", activebackground=BRAND_DARK, relief="flat",
                padx=16, pady=4, state="disabled", command=self.start_selected)
            self.start_btn.pack(side="left")
            self.pick_lbl = tk.Label(bf, text="세트를 고르세요", bg=BG, fg=SUB,
                                     font=("Malgun Gothic", 9))
            self.pick_lbl.pack(side="left", padx=8)
            tk.Button(bf, text="닫기", font=UI_FONT, relief="groove",
                      padx=12, pady=3, command=self.destroy).pack(side="right")
            self.refresh()

        def destroy(self):
            if getattr(self.app, "board_win", None) is self:
                self.app.board_win = None
            super().destroy()

        def refresh(self):
            try:
                board = self.app.current_board()
            except Exception as e:
                log_error("점수대 보드", e)
                return
            by_norm = {s.get("norm"): s for s in self.app.sets}
            rows = {r["norm"]: r for r in board["sets"]}
            todo = sum(b["count"] for b in board["bands"]
                       if b["id"] in ("none", "fail"))
            self.sum_lbl.configure(
                text=band_summary_text(board["bands"])
                + (f"  —  아직 합격선을 못 넘은 세트 {todo}개. 고른 뒤 "
                   "[이 세트 응시]로 바로 다시 풀어 점수를 갱신하세요 "
                   "(응시 횟수 제한 없음)." if todo else
                   "  —  전 세트가 합격선을 넘었습니다. 더 올리고 싶은 세트를 "
                   "골라 다시 풀 수 있어요 (응시 횟수 제한 없음)."))
            for w in self.inner.winfo_children():
                w.destroy()
            self.rows = []
            for b in board["bands"]:
                if not b["count"]:
                    continue
                bg_c, fg_c = self.BAND_COLORS.get(b["id"], (BG, SUB))
                head = tk.Frame(self.inner, bg=CARD, padx=8, pady=3)
                head.pack(fill="x", pady=(6, 0))
                tk.Label(head, text=f" {b['name']} {b['count']} ", bg=bg_c,
                         fg=fg_c, font=("Malgun Gothic", 9, "bold"),
                         padx=4).pack(side="left")
                tk.Label(head, text=b["desc"], bg=CARD, fg=SUB,
                         font=("Malgun Gothic", 8)).pack(side="left", padx=6)
                for norm in b["sets"]:
                    r = rows.get(norm)
                    s = by_norm.get(norm)
                    if r is None:
                        continue
                    self._add_row(r, s, bg_c, fg_c)
            self._select(self.selected)

        def _add_row(self, r, s, bg_c, fg_c):
            row = tk.Frame(self.inner, bg=CARD, padx=10, pady=3,
                           highlightbackground=LINE, highlightthickness=1)
            row.pack(fill="x", padx=6, pady=2)
            best = r["best"]
            head = f"{r['name']}"
            meta = [f"최고 {best:g}점" if best is not None else "기록 없음",
                    f"{len(r['scores'])}회 응시"]
            if r["last_date"]:
                meta.append(f"마지막 {r['last_date']}")
            hist = " → ".join(f"{x:g}" for x in r["scores"])
            if hist:
                meta.append(f"이력 {hist}")
            up = best_update_text(r)
            if up:
                meta.append(up)
            name_lbl = tk.Label(row, text=head, bg=CARD, fg=INK,
                                font=UI_FONT_BOLD, anchor="w")
            name_lbl.pack(fill="x")
            meta_lbl = tk.Label(row, text=" · ".join(meta), bg=CARD, fg=SUB,
                                font=("Malgun Gothic", 9), anchor="w",
                                justify="left")
            meta_lbl.pack(fill="x")
            if s is None:
                tk.Label(row, text="(목록에 없는 세트 — [새로 고침])", bg=CARD,
                         fg=RED, font=("Malgun Gothic", 8),
                         anchor="w").pack(fill="x")
                return
            self.rows.append((r["norm"], s, row, (name_lbl, meta_lbl)))
            for w in (row, name_lbl, meta_lbl):
                w.bind("<Button-1>",
                       lambda e, n=r["norm"]: self._select(n))
                w.bind("<Double-Button-1>",
                       lambda e, n=r["norm"]: (self._select(n),
                                               self.start_selected()))

        def _select(self, norm):
            self.selected = norm if any(x[0] == norm for x in self.rows) \
                else None
            picked = None
            for n, s, row, lbls in self.rows:
                on = (n == self.selected)
                bg_c = BRAND_SOFT if on else CARD
                row.configure(bg=bg_c,
                              highlightbackground=BRAND if on else LINE)
                for lb in lbls:
                    lb.configure(bg=bg_c)
                if on:
                    picked = s
            if picked is None:
                self.start_btn.configure(state="disabled")
                self.pick_lbl.configure(text="세트를 고르세요")
            else:
                self.start_btn.configure(state="normal")
                self.pick_lbl.configure(text=f"선택: {picked['name']}")

        def start_selected(self):
            s = next((x[1] for x in self.rows if x[0] == self.selected), None)
            if s is None:
                messagebox.showinfo(APP_TITLE, "먼저 세트를 고르세요.",
                                    parent=self)
                return
            self.app.start_set_now(s)

    class SetChooserDialog(tk.Toplevel):

        def __init__(self, guide, sets, slot_count, on_apply):
            super().__init__(guide)
            self.sets = sets
            self.on_apply = on_apply
            self.title(f"{APP_TITLE} - 다른 세트로 바꾸기")
            self.configure(bg=BG)
            self.geometry("420x360")
            frm = tk.Frame(self, bg=BG, padx=14, pady=10)
            frm.pack(fill="both", expand=True)
            top = tk.Frame(frm, bg=BG)
            top.pack(fill="x")
            tk.Label(top, text="바꿀 세트 슬롯:", bg=BG, fg=INK,
                     font=UI_FONT).pack(side="left")
            self.slot_var = keep_var(tk.IntVar(value=1))
            for k in range(max(1, slot_count)):
                tk.Radiobutton(top, text=f"{k + 1}번", variable=self.slot_var,
                               value=k + 1, bg=BG, fg=INK, font=UI_FONT,
                               selectcolor=CARD).pack(side="left", padx=4)
            listfrm = tk.Frame(frm, bg=CARD, highlightbackground=LINE,
                               highlightthickness=1)
            listfrm.pack(fill="both", expand=True, pady=(8, 8))
            self.listbox = tk.Listbox(
                listfrm, font=UI_FONT, bd=0, highlightthickness=0, bg=CARD,
                fg=INK, selectbackground=BRAND_SOFT,
                selectforeground=BRAND_DARK, activestyle="none",
                exportselection=False)
            sb = tk.Scrollbar(listfrm, command=self.listbox.yview)
            self.listbox.configure(yscrollcommand=sb.set)
            self.listbox.pack(side="left", fill="both", expand=True,
                              padx=(6, 0), pady=6)
            sb.pack(side="right", fill="y")
            for s in sets:
                self.listbox.insert("end", " " + s["name"])
            bf = tk.Frame(frm, bg=BG)
            bf.pack(fill="x")
            tk.Button(bf, text="이 세트로 바꾸기", font=UI_FONT_BOLD,
                      bg=BRAND, fg="white", activebackground=BRAND_DARK,
                      relief="flat", padx=14, pady=4,
                      command=self.apply).pack(side="left")
            tk.Button(bf, text="취소", font=UI_FONT, relief="groove",
                      padx=12, pady=4, command=self.destroy).pack(
                side="right")

        def apply(self):
            sel = self.listbox.curselection()
            if not sel or sel[0] >= len(self.sets):
                messagebox.showinfo(APP_TITLE, "목록에서 세트를 선택하세요.",
                                    parent=self)
                return
            self.on_apply(int(self.slot_var.get()) - 1, self.sets[sel[0]])
            self.destroy()

    class StepGuideWindow(tk.Toplevel):

        def __init__(self, app, plan):
            super().__init__(app)
            self.app = app
            self.plan = plan
            self.day_tag = plan_day_tag(plan["no"])
            self.slot_sets = []
            self.steps = []
            self._resolve_slots(save=True)
            self.done = load_step_progress(self.day_tag) \
                if self.day_tag else set()
            self._celebrated = self._all_done()
            self.title(f"{APP_TITLE} - 오늘 일정")
            self.configure(bg=BG)
            self.geometry("600x540")
            frm = tk.Frame(self, bg=BG, padx=16, pady=12)
            frm.pack(fill="both", expand=True)
            self.title_lbl = tk.Label(frm, text="", bg=BG, fg=BRAND_DARK,
                                      font=UI_FONT_BOLD, wraplength=560,
                                      justify="left", anchor="w")
            self.title_lbl.pack(anchor="w", fill="x")
            pick_row = tk.Frame(frm, bg=BG)
            pick_row.pack(fill="x", pady=(2, 0))
            self.pick_lbl = tk.Label(pick_row, text="", bg=BG, fg=INK,
                                     font=("Malgun Gothic", 9),
                                     justify="left", anchor="w",
                                     wraplength=420)
            self.pick_lbl.pack(side="left", fill="x", expand=True)
            self.change_btn = tk.Button(
                pick_row, text="다른 세트로 바꾸기", font=UI_FONT,
                relief="groove", padx=8, pady=2, command=self.change_set)
            if plan.get("세트"):
                self.change_btn.pack(side="right")
            self.missing_frame = tk.Frame(frm, bg=BG)
            self.missing_frame.pack(fill="x")
            self.missing_rows = []
            self.progress_lbl = tk.Label(frm, text="", bg=BG, fg=SUB,
                                         font=("Malgun Gothic", 9))
            self.progress_lbl.pack(anchor="w", pady=(2, 6))
            listfrm = tk.Frame(frm, bg=CARD, highlightbackground=LINE,
                               highlightthickness=1)
            listfrm.pack(fill="both", expand=True)
            self.listbox = tk.Listbox(
                listfrm, font=UI_FONT, bd=0, highlightthickness=0,
                bg=CARD, fg=INK, selectbackground=BRAND_SOFT,
                selectforeground=BRAND_DARK, activestyle="none",
                exportselection=False)
            sb = tk.Scrollbar(listfrm, command=self.listbox.yview)
            self.listbox.configure(yscrollcommand=sb.set)
            self.listbox.pack(side="left", fill="both", expand=True,
                              padx=(6, 0), pady=6)
            sb.pack(side="right", fill="y")
            self.listbox.bind("<<ListboxSelect>>",
                              lambda e: self._show_detail())
            self.detail_lbl = tk.Label(
                frm, text="", bg=CARD, fg=INK, font=("Malgun Gothic", 9),
                justify="left", anchor="nw", padx=10, pady=8, wraplength=540,
                highlightbackground=LINE, highlightthickness=1)
            self.detail_lbl.pack(fill="x", pady=(8, 0))
            bf = tk.Frame(frm, bg=BG)
            bf.pack(fill="x", pady=(10, 0))
            self.start_btn = tk.Button(
                bf, text="이 단계 시작", font=UI_FONT_BOLD, bg=BRAND,
                fg="white", activebackground=BRAND_DARK, relief="flat",
                padx=16, pady=4, command=self.start_step)
            self.start_btn.pack(side="left")
            self.check_btn = tk.Button(
                bf, text="완료 체크", font=UI_FONT_BOLD, bg=BRAND_SOFT,
                fg=BRAND_DARK, activebackground="#CFE9DA", relief="flat",
                padx=12, pady=4, command=self.toggle_check)
            self.check_btn.pack(side="left", padx=6)
            self.web_btn = tk.Button(
                bf, text="웹에서 퀴즈 풀기", font=UI_FONT_BOLD, bg="#1F5FBF",
                fg="white", activebackground="#174A96", relief="flat",
                padx=12, pady=4, command=self.open_web_step)
            tk.Button(bf, text="닫기", font=UI_FONT, relief="groove",
                      padx=12, pady=4, command=self.destroy).pack(
                side="right")
            self._render_header()
            self.refresh(select=self.current_index())

        def web_tab_of(self, i=None):
            i = self.selected_index() if i is None else i
            if i is None or i >= len(self.steps):
                return None
            return self.steps[i].get("웹탭") or None

        def open_web_step(self):
            tab = self.web_tab_of()
            if not tab:
                return None
            return self.app.open_routine_page(tab=tab)

        def _resolve_slots(self, save=False):
            plan = self.plan
            if plan.get("세트"):
                try:
                    saved = load_auto_picks(self.day_tag) if self.day_tag \
                        else []
                    self.slot_sets = resolve_day_sets(
                        plan, self.app.sets, load_records(), saved=saved,
                        mapping=load_slot_mapping())
                    for s, _r in self.slot_sets:
                        if s is not None:
                            self.app.add_set_to_list(s)
                    if save and self.day_tag and AUTO in plan["세트"]:
                        self._save_picks()
                except Exception as e:
                    log_error("세트 슬롯 해석", e)
                    self.slot_sets = [(None, f"세트 해석 오류: {e}")
                                      for _ in plan["세트"]]
            else:
                self.slot_sets = []
            self.slot_names = [
                (s["name"] if s else (spec if spec != AUTO
                                      else "자동 선택(세트 없음)"))
                for (s, _r), spec in zip(self.slot_sets, plan["세트"])]
            self.steps = build_day_steps(plan, self.slot_names)

        def _save_picks(self):
            picks = []
            for (s, reason), spec in zip(self.slot_sets, self.plan["세트"]):
                picks.append({"세트": s["name"],
                              "이유": reason.replace(" (유지)", "")}
                             if (s and spec == AUTO) else None)
            try:
                save_auto_picks(self.day_tag, picks)
            except Exception:
                pass

        def pick_text(self):
            lines = []
            for k, ((s, reason), spec) in enumerate(
                    zip(self.slot_sets, self.plan["세트"])):
                head = f"세트 {k + 1}: " if len(self.slot_sets) > 1 else "세트: "
                how = "자동 선택" if spec == AUTO else "일정 지정"
                if s:
                    lines.append(f"{head}{s['name']}  [{how} — {reason}]")
                else:
                    lines.append(f"{head}(없음)  [{how} — {reason}]")
            return "\n".join(lines)

        def _render_header(self):
            self.title_lbl.configure(
                text=plan_title(self.plan, today=date.today(),
                                set_names=self.slot_names or None))
            missing = any(s is None for s, _r in self.slot_sets)
            self.pick_lbl.configure(text=self.pick_text(),
                                    fg=RED if missing else INK)
            self._render_missing_rows()

        def missing_slots(self):
            return [k for k, (s, _r) in enumerate(self.slot_sets) if s is None]

        def _render_missing_rows(self):
            for w in self.missing_frame.winfo_children():
                w.destroy()
            self.missing_rows = []
            for k in self.missing_slots():
                spec = self.plan["세트"][k]
                row = tk.Frame(self.missing_frame, bg=BG)
                row.pack(fill="x", pady=(2, 0))
                head = f"세트 {k + 1} " if len(self.slot_sets) > 1 else ""
                what = spec if spec != AUTO else "자동 선택"
                tk.Label(row, text=f"{head}'{what}' 세트를 찾지 못함 —",
                         bg=BG, fg=RED, font=("Malgun Gothic", 9, "bold")
                         ).pack(side="left")
                b1 = tk.Button(row, text="직접 선택", font=UI_FONT_BOLD,
                               bg=RED, fg="white", activebackground="#8A2A22",
                               relief="flat", padx=8, pady=1,
                               command=lambda k=k: self.direct_select_slot(k))
                b1.pack(side="left", padx=4)
                b2 = tk.Button(row, text="파일명 안내", font=UI_FONT,
                               relief="groove", padx=8, pady=1,
                               command=lambda k=k: self.filename_guide(k))
                b2.pack(side="left")
                self.missing_rows.append((k, row, b1, b2))

        def filename_guide(self, k):
            spec = self.plan["세트"][k]
            idt = "·".join(sorted(slot_identity_tokens(spec))) \
                if spec != AUTO else "(자동 선택)"
            ex = spec if spec != AUTO else "세트명"
            messagebox.showinfo(
                f"{APP_TITLE} - 파일명 안내",
                f"'{spec}' 세트로 인식되려면:\n\n"
                f"1) 문제와 정답 파일이 같은 폴더에 있고\n"
                f"2) 파일명에 식별 토큰 [{idt}] 이 들어가며\n"
                f"3) 이름 끝에 '문제'/'정답'이 붙어 짝을 이루면 됩니다.\n"
                f"   예: '{ex}_문제.xlsx' + '{ex}_정답.xlsm' (확장자 달라도 됨)\n\n"
                "지금 바로 하려면 [직접 선택]으로 문제/정답 파일을 고르세요 — "
                "한 번 고르면 저장되어 다음부터 자동으로 잡힙니다.\n"
                "인식 상태는 시작 화면의 [세트 인식 진단]으로 확인할 수 있습니다.",
                parent=self)

        def direct_select_slot(self, k):
            spec = self.plan["세트"][k]
            try:
                problem = filedialog.askopenfilename(
                    parent=self, title=f"'{spec}' 문제 파일 선택",
                    filetypes=[("Excel 파일", "*.xlsx *.xlsm"),
                               ("모든 파일", "*.*")])
                if not problem:
                    return
                answer = filedialog.askopenfilename(
                    parent=self, title=f"'{spec}' 정답 파일 선택",
                    initialdir=os.path.dirname(problem),
                    filetypes=[("Excel 파일", "*.xlsx *.xlsm"),
                               ("모든 파일", "*.*")])
                if not answer:
                    return
                s = build_direct_set(problem, answer)
                s["direct"] = True
                self.apply_direct_slot(k, s)
            except Exception as e:
                log_error(f"직접 선택 슬롯 {k}", e)
                messagebox.showerror(
                    APP_TITLE, f"'{spec}' 세트를 지정하지 못했습니다: {e}",
                    parent=self)

        def apply_direct_slot(self, k, s):
            spec = self.plan["세트"][k]
            if spec != AUTO:
                try:
                    save_slot_mapping(spec, s)
                    remember_set(s)
                except Exception as e:
                    log_error("슬롯 매핑 저장", e)
            self.app.add_set_to_list(s)
            new = list(self.slot_sets)
            new[k] = (s, "직접 선택")
            self.slot_sets = new
            self.slot_names = [
                (x["name"] if x else (sp if sp != AUTO else "자동 선택(세트 없음)"))
                for (x, _r), sp in zip(self.slot_sets, self.plan["세트"])]
            self.steps = build_day_steps(self.plan, self.slot_names)
            if spec == AUTO and self.day_tag:
                self._save_picks()
            self._render_header()
            self.refresh(select=self.selected_index())

        def change_set(self):
            if not self.plan.get("세트"):
                return
            if not self.app.sets:
                messagebox.showinfo(APP_TITLE, "바꿀 세트가 없습니다. 세트를 "
                                    "스캔 폴더에 넣거나 [직접 선택]하세요.",
                                    parent=self)
                return
            SetChooserDialog(self, self.app.sets, len(self.plan["세트"]),
                             self.apply_set_change)

        def apply_set_change(self, slot, s):
            if not (0 <= slot < len(self.slot_sets)):
                return
            new = list(self.slot_sets)
            for k, (x, _r) in enumerate(new):
                if k != slot and x and x.get("norm") == s.get("norm"):
                    new[k] = (new[slot][0], "직접 선택 (맞교환)")
            new[slot] = (s, "직접 선택")
            self.slot_sets = new
            self.slot_names = [
                (x["name"] if x else spec)
                for (x, _r), spec in zip(self.slot_sets, self.plan["세트"])]
            self.steps = build_day_steps(self.plan, self.slot_names)
            if self.day_tag:
                picks = [{"세트": x["name"], "이유": r} if x else None
                         for (x, r) in self.slot_sets]
                try:
                    save_auto_picks(self.day_tag, picks)
                except Exception:
                    pass
            self._render_header()
            self.refresh(select=self.selected_index())

        def _required_steps(self):
            return mandatory_step_indexes(self.steps)

        def _all_done(self):
            req = self._required_steps()
            return bool(req) and all(i in self.done for i in req)

        def current_index(self):
            for i in range(len(self.steps)):
                if i not in self.done:
                    return i
            return max(0, len(self.steps) - 1)

        def selected_index(self):
            sel = self.listbox.curselection()
            return sel[0] if sel and sel[0] < len(self.steps) else None

        def refresh(self, select=None):
            if select is None:
                select = self.selected_index()
            cur = self.current_index()
            self.listbox.delete(0, "end")
            for i, st in enumerate(self.steps):
                if i in self.done:
                    mark = "[v]"
                elif i == cur:
                    mark = " ▶ "
                else:
                    mark = "    "
                self.listbox.insert(
                    "end",
                    f" {mark} {i + 1}. {st['이름']}  ({st.get('분', '?')}분)")
                if i in self.done:
                    self.listbox.itemconfigure(i, foreground=SUB)
            req = self._required_steps()
            n_req = len([i for i in req if i in self.done])
            opt = [i for i in range(len(self.steps)) if i not in req]
            n_opt = len([i for i in opt if i in self.done])
            txt = f"필수 {n_req}/{len(req)} 완료"
            if opt:
                txt += f" · 선택 {n_opt}/{len(opt)}"
            txt += (" · 필수(완주·채점)만 끝내면 오늘 클리어 · 체크는 자동 저장"
                    if req else " · 체크는 자동 저장")
            self.progress_lbl.configure(text=txt)
            if select is None:
                select = cur
            select = max(0, min(select, len(self.steps) - 1)) \
                if self.steps else None
            if select is not None:
                self.listbox.selection_clear(0, "end")
                self.listbox.selection_set(select)
                self.listbox.see(select)
            self._show_detail()

        def _show_detail(self):
            i = self.selected_index()
            if i is None:
                self.detail_lbl.configure(text="스텝을 선택하세요.")
                return
            st = self.steps[i]
            state = "완료" if i in self.done else (
                "지금 할 차례" if i == self.current_index() else "대기")
            kind = st.get("형", "안내")
            need = "선택" if st.get("선택") else "필수"
            lines = [f"{i + 1}. {st['이름']}  ·  {kind}  ·  {need}  ·  예상 "
                     f"{st.get('분', '?')}분  ·  {state}",
                     "", st.get("설명", "")]
            lines.append("")
            if st.get("선택"):
                lines.append("선택 단계입니다 — 체크하지 않아도 오늘 클리어에 "
                             "영향이 없습니다.")
            if kind in ("안내", "채점"):
                lines.append("이 단계는 직접 하고 [완료 체크]를 누르면 "
                             "됩니다." if kind == "안내" else
                             "채점이 끝나면 자동으로 체크됩니다 (직접 [완료 "
                             "체크]도 가능).")
            elif kind == "모의":
                lines.append("[이 단계 시작]을 누르면 풀이 사본과 40분 타이머가 "
                             "열립니다. 제출·채점까지 끝나면 자동으로 "
                             "체크됩니다.")
            elif kind == "오답노트":
                lines.append("[이 단계 시작]을 누르면 오답노트 모드가 열립니다 "
                             "(선택 도구). 채점 리포트에서 틀린 부분을 이미 "
                             "확인했다면 건너뛰어도 됩니다.")
            elif kind == "오답재풀이":
                lines.append("[이 단계 시작]을 누르면 오답 시트 목록으로 새 "
                             "사본과 15분 타이머가 열리고, 제출하면 그 "
                             "시트들만 채점된 뒤 자동 체크됩니다.")
            elif kind in ("부분연습", "드릴"):
                lines.append("[이 단계 시작]을 누르면 풀이 사본과 타이머가 "
                             "열립니다. 채점까지 끝나면 자동으로 "
                             "체크됩니다.")
            if st.get("웹탭"):
                lines.append("[웹에서 퀴즈 풀기]를 누르면 루틴 페이지의 퀴즈 "
                             "탭이 브라우저로 열립니다. 페이지에서 ⑤를 "
                             "체크하면 이 단계도 자동으로 체크됩니다.")
                if not self.web_btn.winfo_manager():
                    self.web_btn.pack(side="left", padx=6)
            elif self.web_btn.winfo_manager():
                self.web_btn.pack_forget()
            self.detail_lbl.configure(text="\n".join(lines))

        def _auto_steps_for(self, i):
            out = [i]
            if i + 1 < len(self.steps) and \
                    self.steps[i + 1].get("형") == "채점":
                out.append(i + 1)
            return out

        def start_step(self):
            i = self.selected_index()
            if i is None:
                return
            st = self.steps[i]
            try:
                self._start_step(i, st)
            except Exception as e:
                log_error(f"스텝 시작 실패: {st.get('이름')}", e)
                k = st.get("슬롯")
                name = (self.slot_names[k] if k is not None
                        and k < len(self.slot_names) else st.get("세트") or "")
                messagebox.showerror(
                    APP_TITLE,
                    f"'{name or st.get('이름')}' 세트를 시작하지 못했습니다: "
                    f"{e}\n\n오류 내용은 {os.path.basename(ERROR_LOG_PATH)}에 "
                    "기록되었습니다.", parent=self)

        def _start_step(self, i, st):
            kind, payload = resolve_step_action(st, self.app.sets,
                                                self.slot_sets)
            if kind == "info":
                if (payload or {}).get("웹탭"):
                    self.app.open_routine_page(tab=payload["웹탭"])
                    return
                messagebox.showinfo(f"{APP_TITLE} - {st['이름']}",
                                    (payload or {}).get("메시지") or
                                    st.get("설명", ""), parent=self)
                if (payload or {}).get("자동완료"):
                    self.mark_step_done(i)
                return
            if kind == "missing":
                messagebox.showinfo(APP_TITLE, (payload or {}).get("이유")
                                    or "필요한 세트를 찾지 못했습니다.",
                                    parent=self)
                return
            if kind == "review":
                s = (payload or {}).get("set")
                if s:
                    self.app._select_set_in_list(s)
                self.app.open_review_mode(full_only=True)
                return
            if self.app.exam_running:
                messagebox.showinfo(APP_TITLE, "이미 시험이 진행 중입니다.",
                                    parent=self)
                return
            if kind == "retry":
                s = payload["set"]
                self.app._select_set_in_list(s)
                self.app._pending_plan = {"day": self.day_tag, "목표": None,
                                          "step": i, "done_steps": [i]}
                self.app.start_exam(practice={
                    "sheets": payload["sheets"], "label": payload["label"],
                    "minutes": payload["minutes"], "mode": "오답재풀이"})
                return
            if kind == "practice":
                s = payload.get("set")
                if s:
                    self.app._select_set_in_list(s)
                elif not self.app._selected_set():
                    ment = payload.get("세트문구")
                    messagebox.showinfo(
                        APP_TITLE,
                        (f"'{ment}' 세트를 자동으로 찾지 못했습니다.\n"
                         if ment else "") +
                        "아래 목록에서 연습할 세트를 선택한 뒤 다시 "
                        "누르세요.", parent=self)
                    return
                self.app._pending_plan = {"day": self.day_tag, "목표": None,
                                          "step": i, "done_steps": [i]}
                self.app.start_exam(practice={
                    "sheets": payload["sheets"], "label": payload["label"],
                    "minutes": payload["minutes"]})
                return
            if kind == "exam":
                self.app._select_set_in_list(payload["set"])
                try:
                    self.app.minutes_var.set(int(payload.get("minutes", 40)))
                except Exception:
                    pass
                self.app._pending_plan = {"day": self.day_tag,
                                          "목표": payload.get("목표"),
                                          "step": i,
                                          "done_steps": self._auto_steps_for(i)}
                self.app.start_exam()

        def toggle_check(self):
            i = self.selected_index()
            if i is None:
                return
            if i in self.done:
                self.done.discard(i)
            else:
                self.done.add(i)
            self._persist()
            self.refresh(select=self.current_index())
            self.maybe_celebrate()

        def mark_step_done(self, idx):
            if 0 <= idx < len(self.steps) and idx not in self.done:
                self.done.add(idx)
                self._persist()
                self.refresh(select=self.current_index())
                self.maybe_celebrate()

        def _persist(self):
            if self.day_tag:
                try:
                    save_step_progress(self.day_tag, self.done)
                except Exception:
                    pass

        def maybe_celebrate(self):
            if self._all_done() and not self._celebrated:
                self._celebrated = True
                messagebox.showinfo(APP_TITLE, STEP_DONE_MESSAGE,
                                    parent=self)
            elif not self._all_done():
                self._celebrated = False

    class ExamApp(tk.Tk):

        def __init__(self, scan_root=None, auto_update=True):
            super().__init__()
            self.title(APP_TITLE)
            self.configure(bg=BG)
            self.minsize(680, 520)
            self.scan_root = scan_root or configured_scan_root()
            self.scan_root_fixed = bool(scan_root)
            self.sets = []
            self.hidden_sets = []
            self.settings_win = None
            self.grade_py = find_grade_py()
            self.exam_running = False
            self._grade_state = None
            self.auto_update = auto_update
            self.updater = UpdateCoordinator()
            self._update_state = None
            self._update_progress = None
            self._poll_token = 0
            self._toast_after = None
            self._toast_click = None
            self._regrade_summary = None
            self.status_card = None
            self._gauge_goal = GOAL_TIERS[-1]
            self._gauge_frac = 0.0
            self.board_win = None
            self._warned_dirty = set()
            self.plan_no = routine_day_no()
            self.adaptive = None
            self._pending_plan = None
            self.step_guide = None
            self._current_exam = None
            self._excel_warn = None
            self.routine = None
            self.routine_error = ""
            self._routine_jobs = queue.Queue()
            self.exam_date_win = None
            self._build_ui()
            self.refresh_sets()
            self.refresh_records()
            self._start_routine_server()
            excel = find_excel_exe()
            startup_log(f"시작 화면 준비: 스캔 루트={self.scan_root} · "
                        f"세트 {len(self.sets)}개 · "
                        f"grade.py={self.grade_py or '못 찾음'} · "
                        f"Excel 실행 파일={excel or '못 찾음(기본 프로그램으로 엶)'}")
            notice = consume_update_notice()
            if notice:
                startup_log(f"자동 업데이트 완료 확인: v{notice.get('from')} → "
                            f"v{notice.get('to')} · {notice.get('notes') or ''}")
                self.after(300, lambda: self.show_toast(
                    f"v{notice.get('to')}(으)로 자동 업데이트됨 — "
                    f"{notice.get('notes') or '변경 사항 안내 없음'}",
                    seconds=15))
            if auto_update and not first_run_notice_pending():
                self._begin_update_check()
            elif auto_update:
                startup_log("첫 실행 — 자동 업데이트 동의를 받은 뒤에 확인합니다")
            if auto_update:
                self.after(600, self._start_regrade)
            if auto_update and self.routine_auto_open_var.get() \
                    and self.routine is not None:
                self.after(900, self.open_routine_page)
            self.after(150, self._first_run_flow)

        def _begin_update_check(self):
            startup_log("업데이트 확인 시작 (백그라운드, 타임아웃 3초)")
            self._poll_token += 1
            tok = self._poll_token
            threading.Thread(target=self._bg_update_check, daemon=True).start()
            self.after(1200, lambda: self._update_poll(tok))
            return tok

        def ask_update_consent(self):
            startup_log("첫 실행 자동 업데이트 동의 창 표시")
            agreed = False
            try:
                dlg = UpdateConsentDialog(self)
                self.wait_window(dlg)
                agreed = bool(getattr(dlg, "agreed", False))
            except Exception as e:
                log_error("자동 업데이트 동의 창", e)
                agreed = False
            apply_update_consent(agreed)
            try:
                self.auto_update_var.set(agreed)
            except Exception:
                pass
            startup_log(f"자동 업데이트 {'동의' if agreed else '거절'} → "
                        f"세트설정 _설정.{AUTO_UPDATE_SETTING}="
                        f"{'true' if agreed else 'false'}")
            self.show_toast(
                "자동 업데이트에 동의했습니다 — 실행할 때마다 새 버전을 "
                "자동으로 적용합니다 ([설정]에서 끌 수 있습니다)" if agreed else
                "자동 업데이트를 끈 채로 시작합니다 — 필요할 때 [업데이트 "
                "확인]으로 직접 받으면 됩니다 ([설정]에서 켤 수 있습니다)",
                seconds=10)
            return agreed

        def _first_run_flow(self):
            try:
                if first_run_notice_pending():
                    agreed = self.ask_update_consent()
                    if agreed and self.auto_update:
                        self._begin_update_check()
                if not exam_date_set():
                    self.open_exam_date_dialog(first_run=True)
            except Exception as e:
                log_error("첫 실행 안내", e)

        def open_settings(self):
            win = getattr(self, "settings_win", None)
            if win is not None and win.winfo_exists():
                win.lift()
                win.refresh_sets_list()
                return win
            self.settings_win = SettingsDialog(self)
            return self.settings_win

        def open_exam_date_dialog(self, first_run=False):
            win = getattr(self, "exam_date_win", None)
            if win is not None and win.winfo_exists():
                win.lift()
                return win
            self.exam_date_win = ExamDateDialog(
                self, on_save=self._on_exam_date_saved, first_run=first_run)
            return self.exam_date_win

        def _on_exam_date_saved(self, dates):
            startup_log("시험일 설정: "
                        + (" · ".join(d.isoformat() for d in dates) or "없음"))
            self.plan_no = routine_day_no()
            self.recompute_plan()
            self.refresh_records()
            routine_touch()

        def destroy(self):
            srv, self.routine = self.routine, None
            if srv is not None:
                try:
                    srv.stop()
                except Exception:
                    pass
            super().destroy()

        def _build_ui(self):
            header = tk.Frame(self, bg=BRAND)
            header.pack(fill="x")
            tk.Label(header, text=APP_TITLE, bg=BRAND, fg="white",
                     font=("Malgun Gothic", 16, "bold")).pack(
                side="left", padx=18, pady=10)
            tk.Label(header, text="컴활 2급 실기 모의고사 런처", bg=BRAND,
                     fg="#CFE9DA", font=UI_FONT).pack(side="left")

            self.toast = tk.Frame(self, bg="#FFF4D6", padx=12, pady=6,
                                  highlightbackground="#E5C87A",
                                  highlightthickness=1, cursor="hand2")
            self.toast_lbl = tk.Label(
                self.toast, text="", bg="#FFF4D6", fg="#5A4300",
                font=("Malgun Gothic", 9), justify="left", anchor="w",
                wraplength=600)
            self.toast_lbl.pack(side="left", fill="x", expand=True)
            self.toast_close_btn = tk.Button(
                self.toast, text="✕", bg="#FFF4D6", fg="#5A4300",
                relief="flat", font=("Malgun Gothic", 9), padx=4,
                command=self.hide_toast)
            self.toast_close_btn.pack(side="right")
            self.toast_lbl.bind("<Button-1>", self._on_toast_click)
            self.toast.bind("<Button-1>", self._on_toast_click)

            plan_card = tk.Frame(self, bg=BRAND_SOFT, padx=14, pady=8)
            plan_card.pack(fill="x", padx=16, pady=(10, 0))
            self.plan_card = plan_card
            row1 = tk.Frame(plan_card, bg=BRAND_SOFT)
            row1.pack(fill="x")
            tk.Button(row1, text="◀", font=UI_FONT, relief="flat",
                      bg=BRAND_SOFT, fg=BRAND_DARK, padx=4,
                      command=lambda: self._shift_plan(-1)).pack(side="left")
            self.plan_title_lbl = tk.Label(row1, text="", bg=BRAND_SOFT,
                                           fg=BRAND_DARK, font=UI_FONT_BOLD)
            self.plan_title_lbl.pack(side="left", padx=6)
            tk.Button(row1, text="▶", font=UI_FONT, relief="flat",
                      bg=BRAND_SOFT, fg=BRAND_DARK, padx=4,
                      command=lambda: self._shift_plan(1)).pack(side="left")
            self.plan_start_btn = tk.Button(
                row1, text="오늘 일정 시작", font=UI_FONT_BOLD, bg=BRAND,
                fg="white", activebackground=BRAND_DARK, relief="flat",
                padx=16, pady=4, command=self.start_today_plan)
            self.plan_start_btn.pack(side="right")
            self.recalc_btn = tk.Button(
                row1, text="일정 다시 계산", font=UI_FONT, relief="groove",
                bg=BRAND_SOFT, fg=BRAND_DARK, padx=8, pady=2,
                command=self.recompute_plan)
            self.recalc_btn.pack(side="right", padx=(0, 8))
            self.exam_date_btn = tk.Button(
                row1, text="시험일 설정", font=UI_FONT, relief="groove",
                bg=BRAND_SOFT, fg=BRAND_DARK, padx=8, pady=2,
                command=self.open_exam_date_dialog)
            self.exam_date_btn.pack(side="right", padx=(0, 8))
            self.routine_btn = tk.Button(
                row1, text="루틴 열기", font=UI_FONT_BOLD, bg="#1F5FBF",
                fg="white", activebackground="#174A96", relief="flat",
                padx=10, pady=2, command=self.open_routine_page)
            self.routine_btn.pack(side="right", padx=(0, 8))
            self.plan_todo_lbl = tk.Label(plan_card, text="", bg=BRAND_SOFT,
                                          fg=INK, font=("Malgun Gothic", 9),
                                          justify="left", anchor="w",
                                          wraplength=620)
            self.plan_todo_lbl.pack(fill="x", pady=(4, 0))
            self.trust_hint_lbl = tk.Label(
                plan_card, text="매크로 차단 배너가 뜨면 [Excel 신뢰 위치로 "
                "등록]을 눌러주세요", bg=BRAND_SOFT, fg="#8A5A00",
                font=("Malgun Gothic", 8), anchor="w")
            self._render_plan_card()

            self._build_routine_bar(self).pack(fill="x", padx=16, pady=(8, 0))
            self._build_status_card(self).pack(fill="x", padx=16, pady=(8, 0))

            body = tk.Frame(self, bg=BG, padx=16, pady=12)
            body.pack(fill="both", expand=True)
            body.columnconfigure(0, weight=3)
            body.columnconfigure(1, weight=2)
            body.rowconfigure(1, weight=1)

            listhead = tk.Frame(body, bg=BG)
            listhead.grid(row=0, column=0, sticky="w")
            tk.Label(listhead, text="응시 가능한 모의고사", bg=BG, fg=INK,
                     font=UI_FONT_BOLD).pack(side="left")
            self.hidden_note_lbl = tk.Label(
                listhead, text="", bg=BG, fg=SUB, font=("Malgun Gothic", 9))
            self.hidden_note_lbl.pack(side="left", padx=(8, 0))
            listfrm = tk.Frame(body, bg=CARD, highlightbackground=LINE,
                               highlightthickness=1)
            listfrm.grid(row=1, column=0, sticky="nsew", pady=(4, 8))
            self.listbox = tk.Listbox(listfrm, font=UI_FONT, bd=0,
                                      highlightthickness=0, bg=CARD, fg=INK,
                                      selectbackground=BRAND_SOFT,
                                      selectforeground=BRAND_DARK,
                                      activestyle="none",
                                      exportselection=False)
            sb = tk.Scrollbar(listfrm, command=self.listbox.yview)
            self.listbox.configure(yscrollcommand=sb.set)
            self.listbox.pack(side="left", fill="both", expand=True,
                              padx=(6, 0), pady=6)
            sb.pack(side="right", fill="y")
            self.listbox.bind("<<ListboxSelect>>", lambda e: self._show_info())

            tk.Label(body, text="세트 정보", bg=BG, fg=INK,
                     font=UI_FONT_BOLD).grid(row=0, column=1, sticky="w",
                                             padx=(12, 0))
            infofrm = tk.Frame(body, bg=CARD, highlightbackground=LINE,
                               highlightthickness=1)
            infofrm.grid(row=1, column=1, sticky="nsew", padx=(12, 0),
                         pady=(4, 8))
            self.info_lbl = tk.Label(
                infofrm, text="왼쪽 목록에서 세트를 선택하세요.", bg=CARD,
                fg=SUB, font=UI_FONT, justify="left", anchor="nw",
                padx=10, pady=8, wraplength=230)
            self.info_lbl.pack(fill="both", expand=True)
            self.pdf_warn_lbl = tk.Label(
                infofrm, text="", bg="#FDECEA", fg=RED,
                font=("Malgun Gothic", 9, "bold"), justify="left",
                anchor="w", padx=10, pady=6, wraplength=230)

            ctrl = tk.Frame(body, bg=BG)
            ctrl.grid(row=2, column=0, columnspan=2, sticky="ew")
            tk.Label(ctrl, text="시험 시간(분):", bg=BG, fg=INK,
                     font=UI_FONT).pack(side="left")
            self.minutes_var = keep_var(tk.IntVar(value=DEFAULT_MINUTES))
            tk.Spinbox(ctrl, from_=MIN_MINUTES, to=MAX_MINUTES, increment=5,
                       textvariable=self.minutes_var, width=4,
                       font=UI_FONT).pack(side="left", padx=(4, 14))
            self.start_btn = tk.Button(
                ctrl, text="시험 시작", font=UI_FONT_BOLD, bg=BRAND,
                fg="white", activebackground=BRAND_DARK, relief="flat",
                padx=18, pady=5, command=self.start_exam)
            self.start_btn.pack(side="left", padx=4)
            self.practice_btn = tk.Button(
                ctrl, text="부분 연습", font=UI_FONT_BOLD, bg=BRAND_SOFT,
                fg=BRAND_DARK, activebackground="#CFE9DA", relief="flat",
                padx=12, pady=5, command=self.open_practice_dialog)
            self.practice_btn.pack(side="left", padx=4)
            self.review_btn = tk.Button(
                ctrl, text="오답노트 모드", font=UI_FONT_BOLD, bg="#B45309",
                fg="white", activebackground="#8A5A00", relief="flat",
                padx=12, pady=5, state="disabled",
                command=self.open_review_mode)
            self.review_btn.pack(side="left", padx=4)
            tk.Button(ctrl, text="직접 선택...", font=UI_FONT, relief="groove",
                      padx=10, pady=4, command=self.choose_direct).pack(
                side="left", padx=4)
            tk.Button(ctrl, text="새로 고침", font=UI_FONT, relief="groove",
                      padx=10, pady=4, command=self.refresh_all).pack(
                side="left", padx=4)
            tk.Button(ctrl, text="문제지 연결", font=UI_FONT, relief="groove",
                      padx=10, pady=4, command=self.connect_pdf).pack(
                side="left", padx=4)
            self.diag_btn = tk.Button(
                ctrl, text="세트 인식 진단", font=UI_FONT, relief="groove",
                padx=10, pady=4, command=self.show_scan_diagnosis)
            self.diag_btn.pack(side="left", padx=4)
            tk.Button(ctrl, text="시작 로그 보기", font=UI_FONT, relief="groove",
                      padx=10, pady=4, command=self.show_startup_log).pack(
                side="left", padx=4)
            tk.Button(ctrl, text="설정", font=UI_FONT_BOLD, relief="groove",
                      padx=12, pady=4, command=self.open_settings).pack(
                side="left", padx=4)
            tk.Button(ctrl, text="업데이트 확인", font=UI_FONT, relief="groove",
                      padx=10, pady=4,
                      command=self.manual_update_check).pack(
                side="right", padx=4)
            self.auto_update_var = keep_var(
                tk.BooleanVar(value=auto_update_enabled()))
            self.auto_update_chk = tk.Checkbutton(
                ctrl, text="자동 업데이트", variable=self.auto_update_var,
                bg=BG, fg=INK, activebackground=BG, font=UI_FONT,
                command=self._toggle_auto_update)
            self.auto_update_chk.pack(side="right", padx=(0, 2))
            self.routine_auto_open_var = keep_var(tk.BooleanVar(
                value=bool(get_app_setting(ROUTINE_AUTO_OPEN_SETTING, False))))
            self.routine_auto_open_chk = tk.Checkbutton(
                ctrl, text="루틴 자동 열기", variable=self.routine_auto_open_var,
                bg=BG, fg=INK, activebackground=BG, font=UI_FONT,
                command=self._toggle_routine_auto_open)
            self.routine_auto_open_chk.pack(side="right", padx=(0, 2))
            tk.Button(ctrl, text="Excel 신뢰 위치로 등록", font=UI_FONT,
                      relief="groove", padx=10, pady=4,
                      command=self.on_register_trust).pack(
                side="right", padx=4)

            tk.Label(body, text="최근 응시 기록", bg=BG, fg=INK,
                     font=UI_FONT_BOLD).grid(row=3, column=0, columnspan=2,
                                             sticky="w", pady=(10, 0))
            self.records_lbl = tk.Label(
                body, text="아직 응시 기록이 없습니다.", bg=CARD, fg=SUB,
                font=UI_FONT, justify="left", anchor="nw", padx=10, pady=8,
                highlightbackground=LINE, highlightthickness=1)
            self.records_lbl.grid(row=4, column=0, columnspan=2,
                                  sticky="ew", pady=(4, 0))

        def refresh_all(self):
            self.refresh_sets()
            self.refresh_records()

        def refresh_sets(self):
            direct = [s for s in self.sets if s.get("direct")]
            scanned = scan_sets(self.scan_root)
            norms = {s["norm"] for s in scanned}
            merged = scanned + [s for s in direct
                                if s.get("norm") not in norms]
            self.sets, self.hidden_sets = split_hidden_sets(merged)
            self.listbox.delete(0, "end")
            for s in self.sets:
                mark = "[직접] " if s.get("direct") else (
                    "[저장] " if s.get("saved") else "")
                pdf_mark = "" if s.get("pdf") else "  (문제지 미연결)"
                self.listbox.insert("end", f" {mark}{s['name']}{pdf_mark}")
            if self.sets:
                self.listbox.selection_set(0)
                self._show_info()
            if getattr(self, "plan_title_lbl", None) is not None:
                self.recompute_plan()
            self._sync_hidden_note()
            win = getattr(self, "settings_win", None)
            if win is not None and win.winfo_exists():
                try:
                    win.refresh_sets_list()
                except Exception as e:
                    log_error("설정 창 세트 목록 갱신", e)

        def refresh_records(self):
            records = load_records()
            if getattr(self, "plan_title_lbl", None) is not None:
                self.recompute_plan()
            self._sync_start_btn()
            self._render_status_card(records)
            self.refresh_board()
            if not records:
                self.records_lbl.configure(text="아직 응시 기록이 없습니다.")
                return
            lines = []
            for r in records[-5:][::-1]:
                score = r.get("점수")
                score_s = f"{score}점" if score is not None else "채점 실패"
                if r.get("mode") in ("부분연습", "오답재풀이"):
                    mx = r.get("만점")
                    score_s = (f"{score}/{mx:g}점" if score is not None
                               and isinstance(mx, (int, float)) and mx
                               else score_s)
                    score_s += f" [{r.get('mode')} · {r.get('영역', '?')}]"
                lines.append(f"{r.get('일시', '?')}  |  {r.get('세트명', '?')}"
                             f"  |  {score_s}  |  {r.get('소요시간', '-')}")
            self.records_lbl.configure(text="\n".join(lines))

        def _selected_set(self):
            sel = self.listbox.curselection()
            if not sel or sel[0] >= len(self.sets):
                return None
            return self.sets[sel[0]]

        def _sync_start_btn(self):
            btn = getattr(self, "start_btn", None)
            if btn is None or self.exam_running:
                return
            s = self._selected_set()
            btn.configure(state="normal", text="재응시 (점수 갱신)"
                          if (s and set_records(s["name"])) else "시험 시작")

        def _show_info(self):
            s = self._selected_set()
            if not s:
                return
            recs = set_records(s["name"])
            self._sync_start_btn()
            self.review_btn.configure(
                state="normal" if find_latest_result_json(s) else "disabled")
            summ = records_summary(load_records()).get(s["name"], {})
            best = summ.get("best")
            recent = summ.get("recent") or []
            pdf = s.get("pdf")
            lines = [
                f"세트: {s['name']}",
                f"폴더: {s.get('dir', '')}",
                f"정답 파일: 있음",
                f"기대값: {key_link_label(s)}",
                "문제지 PDF: " + (os.path.basename(pdf) if pdf
                                else "미연결 ([문제지 연결]로 지정 가능)"),
                "",
                f"응시 기록: {len(recs)}회",
                f"최고 점수: {best if best is not None else '-'}"
                + (f"  [{band_label(score_band(best))}]"
                   if best is not None else "  [미응시]"),
                f"최근 3회: "
                + (" → ".join(str(x) for x in recent) if recent else "-"),
            ]
            up = best_update_text(self._board_row(s))
            if up:
                lines.append(up + " 갱신")
            self.info_lbl.configure(text="\n".join(lines), fg=INK)
            warn = ""
            if pdf and pdf_conflicts_with_set(s):
                warn = PDF_MISMATCH_WARNING
                if s.get("pdf_확인됨"):
                    warn += " (사용자 확인 연결)"
            elif s.get("pdf_warning"):
                warn = "⚠ " + s["pdf_warning"] + " — [문제지 연결]로 바로잡으세요"
            if warn:
                self.pdf_warn_lbl.configure(text=warn)
                self.pdf_warn_lbl.pack(fill="x", side="bottom")
            else:
                self.pdf_warn_lbl.pack_forget()

        def _build_routine_bar(self, parent):
            bar = tk.Frame(parent, bg="#EAF1FB", padx=14, pady=8,
                           highlightbackground="#B9D0EF", highlightthickness=1)
            self.routine_bar = bar
            top = tk.Frame(bar, bg="#EAF1FB")
            top.pack(fill="x")
            tk.Label(top, text="루틴 웹페이지", bg="#EAF1FB", fg="#174A96",
                     font=UI_FONT_BOLD).pack(side="left")
            self.routine_open_btn = tk.Button(
                top, text="루틴 페이지 열기", font=UI_FONT_BOLD, bg="#1F5FBF",
                fg="white", activebackground="#174A96", relief="flat",
                padx=14, pady=3, command=self.open_routine_page)
            self.routine_open_btn.pack(side="right")
            self.routine_file_btn = tk.Button(
                top, text="파일로 열기", font=UI_FONT, relief="groove",
                bg="#EAF1FB", fg="#174A96", padx=10, pady=2,
                command=self.open_routine_file)
            self.routine_file_btn.pack(side="right", padx=(0, 8))
            self.routine_quiz_btn = tk.Button(
                top, text="함수 퀴즈", font=UI_FONT, relief="groove",
                bg="#EAF1FB", fg="#174A96", padx=10, pady=2,
                command=lambda: self.open_routine_page("quiz"))
            self.routine_quiz_btn.pack(side="right", padx=(0, 8))
            self.routine_bar_lbl = tk.Label(
                bar, text="", bg="#EAF1FB", fg=INK, font=("Malgun Gothic", 9),
                justify="left", anchor="w", wraplength=640)
            self.routine_bar_lbl.pack(fill="x", pady=(4, 0))
            return bar

        def routine_bar_text(self):
            if self.routine_alive():
                return ("오늘 할 일 · 날짜별 일정 · 실수 노트 · 함수 퀴즈를 "
                        "브라우저에서 봅니다. 시험장이 켜져 있는 동안 점수·"
                        "체크가 자동으로 오갑니다 — 따로 내려받는 페이지가 "
                        f"아니라 이 프로그램이 여는 화면입니다.\n주소: "
                        f"{self.routine.url().split('?')[0]} "
                        "(주소창에 직접 입력해도 됩니다)", True)
            why = getattr(self, "routine_error", "") or "알 수 없는 이유"
            return ("⚠ 루틴 연동 서버를 띄우지 못했습니다 — " + why +
                    "\n방화벽이 127.0.0.1 접속을 막았거나 포트를 못 잡은 "
                    "경우입니다. [파일로 열기]를 누르면 페이지 자체는 열 수 "
                    "있지만(폴백 모드) 점수·체크 동기화는 되지 않고 브라우저에 "
                    "저장된 값으로만 돕니다. [새로 고침] 뒤에도 같으면 프로그램을 "
                    "다시 실행해 보세요.", False)

        def _sync_routine_bar(self):
            lbl = getattr(self, "routine_bar_lbl", None)
            if lbl is None:
                return
            text, ok = self.routine_bar_text()
            lbl.configure(text=text, fg=INK if ok else RED)
            for name in ("routine_open_btn", "routine_quiz_btn"):
                btn = getattr(self, name, None)
                if btn is not None:
                    btn.configure(state="normal" if ok else "disabled")
            btn = getattr(self, "routine_btn", None)
            if btn is not None:
                btn.configure(text="루틴 열기" if ok else "루틴 (서버 꺼짐)")

        def _sync_hidden_note(self):
            lbl = getattr(self, "hidden_note_lbl", None)
            if lbl is None:
                return
            n = len(getattr(self, "hidden_sets", []) or [])
            lbl.configure(text=f"· 숨긴 세트 {n}개 ([설정]에서 다시 보이게)"
                          if n else "")

        def open_routine_file(self):
            hp = routine_html_path()
            if not os.path.isfile(hp):
                messagebox.showwarning(
                    APP_TITLE, f"루틴 페이지 파일을 찾지 못했습니다:\n{hp}\n"
                    "[업데이트 확인]으로 내려받거나 배포 ZIP 의 시험장/루틴.html "
                    "을 같은 폴더에 두세요.", parent=self)
                return None
            url = "file:///" + os.path.abspath(hp).replace("\\", "/").lstrip("/")
            try:
                ok = bool(webbrowser.open(url))
            except Exception as e:
                log_error("루틴 페이지 파일 열기", e)
                ok = False
            startup_log(f"루틴 페이지 파일 열기(폴백): {hp} → "
                        f"{'브라우저 실행' if ok else '실패'}")
            self.show_toast(
                "루틴 페이지를 파일로 열었습니다 — 폴백 모드라 시험장과 "
                "동기화되지 않고 브라우저에 저장된 값으로만 돕니다. 연동하려면 "
                "[루틴 페이지 열기]를 쓰세요." if ok else
                f"브라우저를 열지 못했습니다. 이 파일을 직접 여세요: {hp}",
                seconds=10)
            return url if ok else None

        def _start_routine_server(self):
            prefer = ROUTINE_PORT_DEFAULT
            srv = None
            try:
                srv = RoutineServer(
                    sets_fn=lambda: self.sets, ui_call=self._routine_jobs.put,
                    action_fn=self.routine_action,
                    on_change=self.on_routine_changed,
                    exam_running_fn=lambda: self.exam_running, port=prefer)
                srv.start()
                self.routine_error = ""
            except Exception as e:
                log_error("루틴 서버 시작", e)
                startup_log(f"루틴 서버를 띄우지 못했습니다 ({e}) — 연동 없이 계속")
                self.routine_error = str(e) or e.__class__.__name__
                srv = None
            self.routine = srv
            if srv is not None:
                try:
                    if get_app_setting(ROUTINE_PORT_SETTING) != srv.port:
                        set_app_setting(ROUTINE_PORT_SETTING, srv.port)
                except Exception:
                    pass
            self.after(100, self._routine_pump)
            self._sync_routine_bar()
            return srv

        def _routine_pump(self):
            try:
                while True:
                    job = self._routine_jobs.get_nowait()
                    try:
                        job()
                    except Exception as e:
                        log_error("루틴 연동 작업", e)
            except queue.Empty:
                pass
            except Exception:
                pass
            try:
                if self.winfo_exists():
                    self.after(100, self._routine_pump)
            except Exception:
                pass

        def routine_alive(self):
            return self.routine is not None and self.routine.running

        def routine_url(self, tab=None):
            return self.routine.url(tab) if self.routine_alive() else None

        def open_routine_page(self, tab=None):
            if not self.routine_alive():
                why = getattr(self, "routine_error", "") or "알 수 없는 이유"
                if messagebox.askyesno(
                        APP_TITLE,
                        f"루틴 연동 서버를 띄우지 못해 연동 모드로는 열 수 "
                        f"없습니다.\n사유: {why}\n\n대신 루틴 페이지를 "
                        "파일로 열까요? (폴백 모드 — 점수·체크가 시험장과 "
                        "오가지 않고 브라우저에 저장된 값으로만 돕니다)\n\n"
                        "[아니요]를 고르면 [시작 로그 보기]에서 '루틴 서버' "
                        "줄을 확인하고 프로그램을 다시 실행해 보세요.",
                        parent=self):
                    return self.open_routine_file()
                return None
            url = self.routine.url(tab)
            try:
                ok = bool(webbrowser.open(url))
            except Exception as e:
                log_error("루틴 페이지 열기", e)
                ok = False
            base = url.split("?")[0]
            startup_log(f"루틴 페이지 열기: {base}"
                        f"{' #' + tab if tab else ''} → "
                        f"{'브라우저 실행' if ok else '실패'}")
            self.show_toast(
                f"브라우저에서 루틴 페이지를 열었습니다 — 주소: {base} "
                "(페이지는 시험장이 켜져 있는 동안 5초마다 동기화)" if ok else
                f"브라우저를 열지 못했습니다. 주소창에 직접 입력하세요: {url}",
                seconds=10)
            return url if ok else None

        def _toggle_routine_auto_open(self):
            on = bool(self.routine_auto_open_var.get())
            set_app_setting(ROUTINE_AUTO_OPEN_SETTING, on)
            startup_log(f"루틴 자동 열기 {'켬' if on else '끔'} "
                        f"(세트설정 _설정.{ROUTINE_AUTO_OPEN_SETTING})")
            self.show_toast(
                "시험장을 시작할 때 루틴 페이지를 브라우저로 자동으로 엽니다"
                if on else "루틴 페이지 자동 열기를 껐습니다 — [루틴 열기]로 "
                "언제든 열 수 있습니다", seconds=5)
            return on

        def bring_to_front(self):
            try:
                self.deiconify()
                self.lift()
                self.focus_force()
                self.attributes("-topmost", True)
                self.after(400, lambda: self.attributes("-topmost", False))
            except Exception:
                pass

        def _web_pending_plan(self, s, step_kind):
            no = routine_day_no()
            plan = {"day": plan_day_tag(no), "목표": None, "step": None,
                    "done_steps": []}
            try:
                today_plan = self.plan_for(no)
                for i, st in enumerate(today_plan.get("스텝") or []):
                    if st.get("형") == step_kind and st.get("세트") == s["name"]:
                        plan["step"] = i
                        plan["done_steps"] = [i]
                        if step_kind == "모의" and i + 1 < len(today_plan["스텝"]) \
                                and today_plan["스텝"][i + 1].get("형") == "채점":
                            plan["done_steps"].append(i + 1)
                        plan["목표"] = st.get("목표")
                        break
                if plan["목표"] is None and self.adaptive \
                        and self.adaptive.get("kind") == "study":
                    for sl in self.adaptive["days"].get(date.today(), []):
                        if sl.get("set") and sl["set"].get("norm") == s.get("norm"):
                            plan["목표"] = sl.get("goal")
                            break
            except Exception as e:
                log_error("웹 시작 스텝 연결", e)
            return plan

        def routine_action(self, action, s, rec, norm, record_id):
            self.bring_to_front()
            if action == "show":
                return True, None
            if action in ("start_exam", "start_retry") and self.exam_running:
                return False, "시험 진행 중"
            if action == "start_exam":
                if s is None:
                    picks = pick_set_for_retry(self.sets, load_records(), 1)
                    if not picks:
                        return False, "시작할 세트가 없습니다"
                    s = picks[0][0]
                    startup_log(f"웹 요청 자동 선택: {s['name']} ({picks[0][1]})")
                if not self._select_set_in_list(s):
                    return False, "목록에 없는 세트"
                self._pending_plan = self._web_pending_plan(s, "모의")
                startup_log(f"웹 요청: 시험 시작 '{s['name']}'")
                self.start_exam()
                return True, None
            if action == "open_review":
                if s is None:
                    return False, "세트를 찾을 수 없습니다"
                self._select_set_in_list(s)
                startup_log(f"웹 요청: 오답노트 '{s['name']}'")
                self.open_review_mode(full_only=rec is None,
                                      json_path=rec.get("json") if rec else None,
                                      set_info=s)
                return True, None
            if action == "start_retry":
                if s is None:
                    return False, "세트를 찾을 수 없습니다"
                kind, payload = retry_payload_for_set(
                    s, 15, json_path=rec.get("json") if rec else None)
                if kind != "retry":
                    return False, (payload or {}).get("이유") or \
                        (payload or {}).get("메시지") or "재풀이할 오답이 없습니다"
                self._select_set_in_list(s)
                self._pending_plan = self._web_pending_plan(s, "오답재풀이")
                startup_log(f"웹 요청: 오답 재풀이 '{s['name']}' "
                            f"({', '.join(payload['sheets'])})")
                self.start_exam(practice={
                    "sheets": payload["sheets"], "label": payload["label"],
                    "minutes": payload["minutes"], "mode": "오답재풀이"})
                return True, None
            if action == "open_report":
                if not rec or not rec.get("html"):
                    return False, "리포트 파일이 없습니다"
                ok, err = open_file(rec["html"])
                return (True, None) if ok else (False, f"리포트를 열지 못했습니다: {err}")
            if action == "open_pdf":
                if not s or not s.get("pdf") or not os.path.isfile(s["pdf"]):
                    return False, "문제지 PDF 없음"
                ok, err = open_file(s["pdf"])
                return (True, None) if ok else (False, f"PDF를 열지 못했습니다: {err}")
            return False, "알 수 없는 action"

        def on_routine_changed(self, kind, info):
            info = info or {}
            if kind == "score":
                self.refresh_records()
                self._show_info()
                rec = info.get("record") or {}
                self.show_toast(f"루틴 페이지에서 점수 기록: {rec.get('세트명', '?')} "
                                f"{rec.get('점수', '?')}점 (수동)", seconds=6)
                return
            if kind == "review":
                self._show_info()
                return
            if kind == "check":
                key = str(info.get("key") or "")
                if not key.endswith("|quiz"):
                    return
                if key.split("|")[0] != date.today().isoformat():
                    return
                self._sync_quiz_step(bool(info.get("value")))

        def _sync_quiz_step(self, on):
            no = routine_day_no()
            day_tag = plan_day_tag(no)
            if not day_tag:
                return False
            try:
                steps = self.plan_for(no).get("스텝") or []
                idx = next((i for i, st in enumerate(steps)
                            if st.get("웹탭") == "quiz"), None)
                if idx is None:
                    return False
                g = getattr(self, "step_guide", None)
                if g is not None and g.winfo_exists() and g.day_tag == day_tag:
                    if on:
                        g.mark_step_done(idx)
                    elif idx in g.done:
                        g.done.discard(idx)
                        g._persist()
                        g.refresh(select=g.current_index())
                else:
                    done = load_step_progress(day_tag)
                    if on:
                        done.add(idx)
                    else:
                        done.discard(idx)
                    save_step_progress(day_tag, done)
                self._render_plan_card()
                return True
            except Exception as e:
                log_error("웹 퀴즈 체크 동기화", e)
                return False

        def show_toast(self, text, seconds=8, on_click=None, tone=None):
            self._toast_click = on_click
            bg, fg, line = (("#E3F2E8", "#0B5D31", "#9ED3B4") if tone == "good"
                            else ("#FFF4D6", "#5A4300", "#E5C87A"))
            for w in (self.toast, self.toast_lbl, self.toast_close_btn):
                try:
                    w.configure(bg=bg)
                except Exception:
                    pass
            try:
                self.toast_lbl.configure(fg=fg)
                self.toast_close_btn.configure(fg=fg)
                self.toast.configure(highlightbackground=line)
            except Exception:
                pass
            self.toast_lbl.configure(text=text)
            if not self.toast.winfo_manager():
                self.toast.pack(fill="x", padx=16, pady=(8, 0),
                                before=self.plan_card)
            if self._toast_after is not None:
                try:
                    self.after_cancel(self._toast_after)
                except Exception:
                    pass
                self._toast_after = None
            if seconds:
                self._toast_after = self.after(int(seconds * 1000),
                                               self.hide_toast)
            return text

        def hide_toast(self, _event=None):
            self._toast_after = None
            self._toast_click = None
            if self.toast.winfo_manager():
                self.toast.pack_forget()

        def _on_toast_click(self, _event=None):
            cb = self._toast_click
            if cb is None:
                self.hide_toast()
                return
            try:
                cb()
            except Exception as e:
                log_error("안내 띠 상세", e)

        def _start_regrade(self):
            if getattr(self, "_regrade_running", False):
                return
            self._regrade_running = True
            sets = list(self.sets)
            coord = RegradeCoordinator(sets, grade_py=self.grade_py,
                                       progress=self._regrade_progress)
            startup_log("자동 재채점 확인 시작 (백그라운드)")

            def worker():
                try:
                    summary = coord.run()
                except Exception as e:
                    log_error("자동 재채점", e)
                    summary = {"checked": 0, "targets": 0, "regraded": 0,
                               "changes": [], "errors": [str(e)],
                               "data_applied": 0, "skipped": 0}
                try:
                    self.after(0, lambda: self._regrade_done(summary))
                except Exception:
                    pass

            threading.Thread(target=worker, daemon=True).start()

        def _regrade_progress(self, i, n, name):
            if n < 2:
                return
            try:
                self.after(0, lambda: self.show_toast(
                    f"지난 기록 다시 채점 중 {i}/{n} — {name}", seconds=0))
            except Exception:
                pass

        def _regrade_done(self, summary):
            self._regrade_running = False
            self._regrade_summary = summary
            changes = [c for c in (summary or {}).get("changes") or []
                       if isinstance(c, dict)]
            for err in (summary or {}).get("errors") or []:
                startup_log(f"자동 재채점 건너뜀: {err}")
            if not changes:
                if summary and summary.get("targets"):
                    startup_log(
                        f"자동 재채점 완료: 대상 {summary['targets']}건 · "
                        f"다시 채점 {summary.get('regraded', 0)}건 · "
                        "점수 변화 없음")
                self.hide_toast()
                return
            startup_log("이의제기 반영: "
                        + " / ".join(f"{c['name']} {c['before']} → "
                                     f"{c['after']}점 ({c['reason']})"
                                     for c in changes))
            try:
                self.refresh_records()
            except Exception as e:
                log_error("자동 재채점 후 갱신", e)
            self.show_toast(correction_summary_text(changes), seconds=0,
                            on_click=self.show_corrections, tone="good")

        def corrections_report_text(self, changes=None):
            summary = self._regrade_summary or {}
            changes = changes if changes is not None else (
                summary.get("changes") or [])
            lines = ["[이의제기 반영 — 지난 기록 자동 재채점]", ""]
            for c in changes:
                before = c.get("before")
                b = f"{before:g}" if isinstance(before, (int, float)) else "—"
                lines.append(f"· {c.get('date', '?')}  {c.get('name', '?')}"
                             f"  {b} → {c.get('after'):g}점"
                             f"  ({c.get('source', '재채점')})")
                lines.append(f"    사유: {c.get('reason', '')}")
                if c.get("report"):
                    lines.append(f"    새 리포트: {c['report']}")
                lines.append("")
            lines.append(f"확인한 기록 {summary.get('checked', 0)}건 · "
                         f"대상 {summary.get('targets', 0)}건 · "
                         f"다시 채점 {summary.get('regraded', 0)}건 · "
                         f"정정 데이터 {summary.get('data_applied', 0)}건")
            if summary.get("skipped"):
                lines.append(f"풀이 파일이 없어 건너뛴 기록 "
                             f"{summary['skipped']}건 — 정정 데이터가 배포되면 "
                             "자동으로 반영됩니다.")
            for err in summary.get("errors") or []:
                lines.append(f"오류: {err}")
            lines.append("")
            lines.append("예전 리포트 파일은 지우지 않았습니다. 정정된 점수로 "
                         "목표 단계·연속 응시·합격 세트 수·일정 배정·루틴 "
                         "페이지 진행률이 다시 계산됩니다.")
            return "\n".join(lines)

        def show_corrections(self):
            changes = (self._regrade_summary or {}).get("changes") or []
            if not changes:
                self.hide_toast()
                return
            DiagnosisWindow(
                self, self.corrections_report_text(changes),
                title=f"{APP_TITLE} - 이의제기 반영",
                hint="채점기·정답 파일이 바뀌어 지난 기록을 다시 채점했습니다. "
                     "원래 점수는 기록.json 의 정정이력에 남아 있습니다.")

        def _build_status_card(self, parent):
            card = tk.Frame(parent, bg=CARD, highlightbackground=LINE,
                            highlightthickness=1, padx=14, pady=9)
            self.status_card = card
            top = tk.Frame(card, bg=CARD)
            top.pack(fill="x")
            tk.Label(top, text="목표 단계", bg=CARD, fg=SUB,
                     font=("Malgun Gothic", 9)).pack(side="left", padx=(0, 8))
            self.tier_lbls = []
            for g in GOAL_TIERS:
                lbl = tk.Label(top, text="", bg=CARD, fg=SUB,
                               font=("Malgun Gothic", 9, "bold"),
                               padx=7, pady=1, highlightthickness=1,
                               highlightbackground=LINE)
                lbl.pack(side="left", padx=2)
                self.tier_lbls.append((g, lbl))
            self.badge_count_lbl = tk.Label(
                top, text="", bg=CARD, fg=SUB, font=("Malgun Gothic", 9),
                cursor="hand2")
            self.badge_count_lbl.pack(side="right")
            self.badge_count_lbl.bind("<Button-1>",
                                      lambda e: self.show_achievements())
            tk.Button(top, text="도전 과제", font=("Malgun Gothic", 9),
                      relief="groove", padx=8, pady=1,
                      command=self.show_achievements).pack(side="right",
                                                           padx=(0, 8))
            self.gauge = tk.Canvas(card, height=13, bg=CARD, bd=0,
                                   highlightthickness=0)
            self.gauge.pack(fill="x", pady=(7, 3))
            self.gauge.bind("<Configure>", lambda e: self._draw_status_gauge())
            info = tk.Frame(card, bg=CARD)
            info.pack(fill="x")
            self.gauge_lbl = tk.Label(info, text="", bg=CARD, fg=INK,
                                      font=UI_FONT, anchor="w")
            self.gauge_lbl.pack(side="left")
            self.streak_lbl = tk.Label(info, text="", bg=CARD, fg=SUB,
                                       font=("Malgun Gothic", 9), anchor="e")
            self.streak_lbl.pack(side="right")
            self.badge_row = tk.Frame(card, bg=CARD)
            self.badge_row.pack(fill="x", pady=(6, 0))
            self.badge_chips = []
            cols = 6
            for i in range(cols):
                self.badge_row.columnconfigure(i, weight=1, uniform="badge")
            for i, a in enumerate(ACHIEVEMENT_DEFS):
                chip = tk.Label(self.badge_row, text=a["icon"], bg=BG, fg=SUB,
                                font=("Malgun Gothic", 8), padx=4, pady=1,
                                highlightthickness=1,
                                highlightbackground=LINE, cursor="hand2")
                chip.grid(row=i // cols, column=i % cols, padx=2, pady=1,
                          sticky="ew")
                chip.bind("<Button-1>", lambda e: self.show_achievements())
                self.badge_chips.append((a["id"], chip))
            band_row = tk.Frame(card, bg=CARD)
            band_row.pack(fill="x", pady=(7, 0))
            self.band_lbl = tk.Label(
                band_row, text="", bg=CARD, fg=SUB,
                font=("Malgun Gothic", 9), anchor="w", justify="left")
            self.band_lbl.pack(side="left", fill="x", expand=True)
            self.board_btn = tk.Button(
                band_row, text="점수대 보드", font=("Malgun Gothic", 9),
                relief="groove", padx=8, pady=1, command=self.show_score_board)
            self.board_btn.pack(side="right")
            return card

        def _draw_status_gauge(self, frac=None):
            cv = getattr(self, "gauge", None)
            if cv is None or not cv.winfo_exists():
                return
            if frac is None:
                frac = getattr(self, "_gauge_frac", 0.0)
            self._gauge_frac = frac
            w = max(1, cv.winfo_width())
            h = int(cv["height"])
            cv.delete("all")
            cv.create_rectangle(0, 0, w, h, fill=BRAND_SOFT, outline=LINE)
            if frac > 0:
                cv.create_rectangle(0, 0, max(2, int(w * frac)), h,
                                    fill=BRAND, outline="")
            goal = getattr(self, "_gauge_goal", GOAL_TIERS[-1]) or GOAL_TIERS[-1]
            for g in GOAL_TIERS:
                if g >= goal:
                    continue
                cv.create_line(int(w * g / goal), 0, int(w * g / goal), h,
                               fill=CARD)

        def _render_status_card(self, records=None, ser=None):
            if getattr(self, "status_card", None) is None:
                return
            try:
                if ser is None:
                    recs = records if records is not None else load_records()
                    ser = serialize_records(recs, self.sets)[0]
                else:
                    recs = records if records is not None else []
                lv = level_from_best(serialized_best(ser))
                stk = streak_info(recs if records is not None
                                  else load_records())
                summ = achievement_summary(ser)
            except Exception as e:
                log_error("상태 카드 계산", e)
                return
            best, goal = lv["best_total"], lv["current_goal"]
            for g, lbl in self.tier_lbls:
                done = g in lv["achieved"]
                cur = (g == goal) and not done
                lbl.configure(
                    text=f"{'✓' if done else ('▶' if cur else '○')} "
                         f"{tier_label(g)} {g}",
                    bg=BRAND_SOFT if done else (CARD if not cur else "#FBF0DC"),
                    fg=BRAND_DARK if done else ("#8A5A00" if cur else SUB),
                    highlightbackground=BRAND if done else LINE)
            self._gauge_goal = goal
            self._draw_status_gauge(1.0 if lv["tier"] == "perfect"
                                    else gauge_fraction(best, goal))
            if best is None:
                self.gauge_lbl.configure(
                    text=f"아직 기록이 없습니다 — 첫 목표 {goal}점"
                         f"({tier_label(goal)})")
            elif lv["tier"] == "perfect":
                self.gauge_lbl.configure(
                    text=f"최고 {best:g}점 · 만점 달성 — 남은 세트도 100점 도전")
            else:
                self.gauge_lbl.configure(
                    text=f"최고 {best:g}점 / 목표 {goal}점({tier_label(goal)}) "
                         f"— {max(0, goal - best):g}점 남음")
            self.streak_lbl.configure(
                text=f"연속 {stk['current']}일 (최고 {stk['best']}일) · "
                     + ("오늘 응시함" if stk["today_done"] else "오늘 미응시"))
            self.badge_count_lbl.configure(
                text=f"도전 과제 {summ['done']}/{summ['total']}")
            done_ids = {a["id"] for a in summ["items"] if a["done"]}
            by_id = {a["id"]: a for a in summ["items"]}
            for aid, chip in self.badge_chips:
                a = by_id.get(aid) or {}
                on = aid in done_ids
                chip.configure(
                    bg=BRAND_SOFT if on else BG,
                    fg=BRAND_DARK if on else "#93A79A",
                    highlightbackground=BRAND if on else LINE,
                    text=(("✓ " if on else "· ") + str(a.get("icon") or "")))
            self._apply_band_line(ser)

        def current_board(self, records=None):
            recs = records if records is not None else load_records()
            return score_board(self.sets, serialize_records(recs, self.sets)[0])

        def _board_row(self, s):
            if not s:
                return None
            try:
                board = self.current_board()
            except Exception as e:
                log_error("점수대 보드 계산", e)
                return None
            norm = s.get("norm") or ""
            return next((r for r in board["sets"] if r["norm"] == norm), None)

        def _apply_band_line(self, ser=None):
            lbl = getattr(self, "band_lbl", None)
            if lbl is None:
                return
            try:
                recs = ser if ser is not None \
                    else serialize_records(load_records(), self.sets)[0]
                bands = score_board(self.sets, recs)["bands"]
            except Exception as e:
                log_error("점수대 요약", e)
                return
            todo = sum(b["count"] for b in bands
                       if b["id"] in ("none", "fail"))
            tail = (f" — 미완성 {todo}세트" if todo
                    else " — 전 세트 합격선 통과")
            lbl.configure(text="점수대 " + band_summary_text(bands) + tail)

        def show_score_board(self):
            win = getattr(self, "board_win", None)
            if win is not None and win.winfo_exists():
                win.refresh()
                win.lift()
                win.focus_force()
                return
            self.board_win = ScoreBoardWindow(self)

        def refresh_board(self):
            win = getattr(self, "board_win", None)
            if win is not None and win.winfo_exists():
                try:
                    win.refresh()
                except Exception as e:
                    log_error("점수대 보드 갱신", e)

        def start_set_now(self, s):
            if self.exam_running:
                messagebox.showinfo(APP_TITLE, "이미 시험이 진행 중입니다.",
                                    parent=self)
                return
            if not self._select_set_in_list(s):
                messagebox.showinfo(APP_TITLE,
                                    "세트 목록에서 찾지 못했습니다 — [새로 "
                                    "고침] 뒤 다시 시도해 주세요.", parent=self)
                return
            startup_log(f"점수대 보드에서 시작: {s['name']}")
            self.start_exam()

        def show_achievements(self):
            try:
                ser = serialize_records(load_records(), self.sets)[0]
                summ = achievement_summary(ser)
            except Exception as e:
                log_error("도전 과제 계산", e)
                return
            lines = [f"[도전 과제 {summ['done']}/{summ['total']}]", ""]
            group = None
            for a in summ["items"]:
                if a["group"] != group:
                    group = a["group"]
                    lines.append(f"— {group} —")
                mark = "✓ 달성" if a["done"] else "○ 미달성"
                when = f"  ({a['when']})" if a["when"] else ""
                lines.append(f"  {mark}  [{a['icon']}] {a['name']}{when}")
                lines.append(f"        {a['desc']}")
            lines.append("")
            lines.append("판정은 기록.json 의 시험·수동 기록만 보고 계산하며 "
                         "루틴 웹페이지와 같은 정의를 씁니다.")
            DiagnosisWindow(self, "\n".join(lines),
                            title=f"{APP_TITLE} - 도전 과제",
                            hint="완주·점수·연속 응시·정확도·속도로 나뉜 12개 "
                                 "과제입니다. 달성하면 ✓ 로 바뀝니다.")

        def _toggle_auto_update(self):
            on = bool(self.auto_update_var.get())
            set_app_setting(AUTO_UPDATE_SETTING, on)
            startup_log(f"자동 업데이트 {'켬' if on else '끔'} "
                        f"(세트설정 _설정.{AUTO_UPDATE_SETTING})")
            self.show_toast(
                "자동 업데이트를 켰습니다 — 실행할 때마다 새 버전을 자동으로 "
                "적용합니다" if on else
                "자동 업데이트를 껐습니다 — [업데이트 확인]으로 수동 적용할 수 "
                "있습니다", seconds=5)
            return on

        def _bg_update_check(self, force=False):
            try:
                self._update_state = self.updater.check(force=force)
            except Exception as e:
                log_error("업데이트 확인", e)
                self._update_state = "failed"

        def _update_poll(self, token=None, tries=0, manual=False):
            if token is None:
                token = self._poll_token
            if token != self._poll_token:
                return
            state = self._update_state
            if state is None:
                if tries < 600:
                    self.after(500, lambda: self._update_poll(token, tries + 1,
                                                              manual))
                return
            self._update_state = None
            self._on_update_checked(state, manual=manual)

        def _on_update_checked(self, state, manual=False):
            prog = self._update_progress
            self._update_progress = None
            if prog is not None:
                try:
                    prog.destroy()
                except Exception:
                    pass
            if state == "staged":
                return self._apply_update_now(manual=manual)
            if state == "data":
                self.refresh_sets()
                self.show_toast(f"기대값 파일 {self.updater.data_synced}개를 "
                                f"자동으로 내려받았습니다 → "
                                f"{expected_values_dir()}", seconds=10)
            elif manual:
                if state == "latest":
                    messagebox.showinfo(
                        APP_TITLE, f"현재 최신 버전입니다 (v{__version__}).",
                        parent=self)
                elif state == "offline":
                    messagebox.showinfo(
                        APP_TITLE, "업데이트 서버에 연결할 수 없습니다.\n"
                        "인터넷 연결을 확인한 뒤 다시 시도해 주세요.",
                        parent=self)
                elif state == "stuck":
                    messagebox.showwarning(APP_TITLE, self.updater.detail,
                                           parent=self)
                else:
                    CollapsibleErrorDialog(
                        self, APP_TITLE,
                        "업데이트에 실패해 기존 버전을 유지합니다.",
                        self.updater.detail)
            elif state == "failed":
                self.show_toast("자동 업데이트 실패 — 기존 버전을 유지합니다 "
                                "([시작 로그 보기]에서 원인 확인)", seconds=10)
            return state

        def _apply_update_now(self, manual=False):
            r = self.updater.apply_pending(exam_running=self.exam_running)
            ver = (self.updater.info or {}).get("version")
            if r == "applied":
                self.show_toast(f"v{ver}(으)로 업데이트했습니다 — 잠시 후 "
                                "자동으로 다시 시작합니다", seconds=0)
                self._restart_when_idle()
            elif r == "deferred":
                self.show_toast(f"새 버전 v{ver} 준비됨 — 시험이 끝나면 "
                                "자동으로 적용하고 다시 시작합니다", seconds=0)
                if manual:
                    messagebox.showinfo(
                        APP_TITLE, f"새 버전 v{ver}을(를) 내려받았습니다.\n"
                        "시험이 끝나면 자동으로 적용하고 다시 시작합니다.",
                        parent=self)
            elif r == "failed":
                if manual:
                    CollapsibleErrorDialog(
                        self, APP_TITLE,
                        "업데이트에 실패해 기존 버전을 유지합니다.",
                        self.updater.detail)
                else:
                    self.show_toast("자동 업데이트 실패 — 기존 버전을 유지합니다 "
                                    "([시작 로그 보기]에서 원인 확인)",
                                    seconds=10)
            return r

        def _busy_for_restart(self):
            if self.exam_running:
                return True
            try:
                if self.grab_current() is not None:
                    return True
            except Exception:
                pass
            guide = getattr(self, "step_guide", None)
            for w in self.winfo_children():
                if not isinstance(w, tk.Toplevel) or w is guide:
                    continue
                try:
                    if w.winfo_exists() and w.winfo_viewable():
                        return True
                except Exception:
                    pass
            return False

        def _restart_when_idle(self, tries=0):
            if self.exam_running:
                return False
            if self._busy_for_restart():
                if tries == 0:
                    startup_log("재시작 대기: 열린 창·대화상자가 닫히면 재시작")
                self.after(2000, lambda: self._restart_when_idle(tries + 1))
                return False
            return self._try_restart()

        def _try_restart(self):
            if self.exam_running:
                return False
            args, kwargs = restart_command()
            try:
                subprocess.Popen(args, **kwargs)
                startup_log(f"재시작 실행: {args}")
            except Exception as e:
                log_error("재시작", e)
                messagebox.showwarning(
                    APP_TITLE,
                    f"자동 재시작에 실패했습니다 ({e}).\n\n이 창을 닫고 "
                    "코코시험장.pyw(또는 시험장.py)를 직접 다시 실행해 "
                    "주세요.", parent=self)
                return False
            try:
                self.destroy()
            except Exception:
                pass
            return True

        def manual_update_check(self):
            if self.updater.checking or self._update_progress is not None:
                messagebox.showinfo(APP_TITLE, "이미 업데이트를 확인하는 중입니다.",
                                    parent=self)
                return None
            if self.updater.pending:
                return self._apply_update_now(manual=True)
            prog = tk.Toplevel(self)
            prog.title(APP_TITLE)
            prog.attributes("-topmost", True)
            prog.resizable(False, False)
            tk.Label(prog, text="업데이트를 확인하는 중입니다...",
                     padx=26, pady=18, font=UI_FONT).pack()
            prog.update()
            self._update_progress = prog
            self._update_state = None
            self._poll_token += 1
            tok = self._poll_token
            threading.Thread(target=self._bg_update_check,
                             kwargs={"force": True}, daemon=True).start()
            self.after(300, lambda: self._update_poll(tok, manual=True))
            return prog

        def connect_pdf(self):
            s = self._selected_set()
            if not s:
                messagebox.showinfo(APP_TITLE, "먼저 세트를 선택하세요.",
                                    parent=self)
                return
            path = filedialog.askopenfilename(
                parent=self, title=f"'{s['name']}' 문제지 PDF 선택",
                initialdir=s["dir"],
                filetypes=[("PDF 파일", "*.pdf"), ("모든 파일", "*.*")])
            if not path:
                return
            path = os.path.abspath(path)
            confirmed = False
            if pdf_conflicts_with_set(s, path):
                if not messagebox.askyesno(
                        APP_TITLE,
                        "선택한 문제지 파일명의 회차·형·연도가 세트와 다릅니다.\n"
                        f"세트: {s['name']}\n문제지: {os.path.basename(path)}\n\n"
                        "그래도 이 문제지를 연결할까요?", parent=self):
                    return
                confirmed = True
            s["pdf"] = path
            s["pdf_확인됨"] = confirmed
            s.pop("pdf_warning", None)
            remember_set(s)
            startup_log(f"문제지 연결: 세트 '{s['name']}' ← {path}"
                        + (" (회차 불일치 — 사용자 확인)" if confirmed else ""))
            sel = self.listbox.curselection()
            self.refresh_sets()
            if sel:
                self.listbox.selection_clear(0, "end")
                self.listbox.selection_set(min(sel[0],
                                               self.listbox.size() - 1))
            self._show_info()

        def choose_direct(self):
            problem = filedialog.askopenfilename(
                parent=self, title="1/4 문제 파일 선택",
                filetypes=[("Excel 파일", "*.xlsx *.xlsm"), ("모든 파일", "*.*")])
            if not problem:
                return
            answer = filedialog.askopenfilename(
                parent=self, title="2/4 정답 파일 선택",
                initialdir=os.path.dirname(problem),
                filetypes=[("Excel 파일", "*.xlsx *.xlsm"), ("모든 파일", "*.*")])
            if not answer:
                return
            pdf = filedialog.askopenfilename(
                parent=self,
                title="3/4 문제지 PDF 선택 (없으면 [취소]로 건너뛰기)",
                initialdir=os.path.dirname(problem),
                filetypes=[("PDF 파일", "*.pdf"), ("모든 파일", "*.*")])
            s = build_direct_set(problem, answer, pdf=pdf or None)
            s["direct"] = True
            if messagebox.askyesno(
                    APP_TITLE, "4/4 이 구성을 세트로 저장할까요?\n"
                    "(저장하면 다음 실행부터 목록에 자동 표시)", parent=self):
                remember_set(s)
            self.sets.append(s)
            self.listbox.insert("end", f" [직접] {s['name']}")
            self.listbox.selection_clear(0, "end")
            self.listbox.selection_set("end")
            self._show_info()

        def _ensure_grade_py(self):
            self.grade_py = find_grade_py(self.grade_py)
            if self.grade_py:
                return True
            messagebox.showwarning(
                APP_TITLE, "grade.py(채점 프로그램)를 찾지 못했습니다.\n"
                "다음 창에서 grade.py 위치를 직접 선택해 주세요.", parent=self)
            path = filedialog.askopenfilename(
                parent=self, title="grade.py 선택",
                filetypes=[("Python 파일", "*.py")])
            if path and os.path.isfile(path):
                self.grade_py = os.path.abspath(path)
                return True
            return False

        def start_exam(self, practice=None):
            day_plan = getattr(self, "_pending_plan", None)
            self._pending_plan = None
            if self.exam_running:
                messagebox.showinfo(APP_TITLE, "이미 시험이 진행 중입니다.",
                                    parent=self)
                return
            s = self._selected_set()
            if not s:
                messagebox.showinfo(APP_TITLE, "먼저 세트를 선택하세요.",
                                    parent=self)
                return
            for label, p in (("문제", s["problem"]), ("정답", s["answer"])):
                if not os.path.isfile(p):
                    messagebox.showerror(
                        APP_TITLE, f"{label} 파일을 찾을 수 없습니다:\n{p}",
                        parent=self)
                    return
            if not self._ensure_grade_py():
                return
            if not practice and set_records(s["name"]) \
                    and not messagebox.askokcancel(
                    APP_TITLE,
                    "재응시: 원본 문제 파일에서 새 깨끗한 사본으로 "
                    "시작합니다.\n이전 풀이는 풀이_*.xlsx 파일로 그대로 "
                    "보관됩니다.\n계속할까요?", parent=self):
                return
            if s.get("norm") not in self._warned_dirty:
                traces = problem_has_formula_traces(s["problem"])
                if traces:
                    self._warned_dirty.add(s.get("norm"))
                    if not messagebox.askyesno(
                            APP_TITLE,
                            "원본 문제 파일에 풀이 흔적이 있습니다 "
                            f"(계산작업 시트에 수식 {traces}개).\n"
                            "깨끗한 원본이 아니면 채점 기준(diff)이 "
                            "왜곡됩니다 — 그대로 진행할까요?", parent=self):
                        return
            needs_macro = practice is None or any(
                "매크로" in str(sh) for sh in (practice.get("sheets") or []))
            if needs_macro and s.get("norm"):
                cfg = load_set_config()
                ent = cfg.setdefault(s["norm"], {})
                if not ent.get("매크로안내"):
                    ent["매크로안내"] = True
                    save_set_config(cfg)
                    messagebox.showinfo(
                        APP_TITLE,
                        "매크로 안내\n\n빨간 '매크로 차단' 배너가 보이면:\n"
                        "1) 파일 우클릭 → 속성 → '차단 해제' 체크, 또는\n"
                        "2) 시험장의 [Excel 신뢰 위치로 등록] 사용\n\n"
                        "저장할 때는 반드시 .xlsm 형식을 유지하세요 "
                        "(풀이 사본은 자동으로 .xlsm으로 만들어 드립니다).",
                        parent=self)
            mode = (practice.get("mode") or "부분연습") if practice else "시험"
            startup_log(f"세트 선택: '{s['name']}' · 모드={mode} · "
                        f"문제={s['problem']} · PDF={s.get('pdf') or '없음'}")
            try:
                if practice and practice.get("mode") == "오답재풀이":
                    student = make_retry_copy(s["problem"], s["name"])
                elif practice:
                    student = make_practice_copy(s["problem"], s["name"],
                                                 practice["label"])
                else:
                    student = make_attempt_copy(s["problem"], s["name"])
            except Exception as e:
                text = log_error(f"풀이 사본 생성: {s.get('name')}", e)
                CollapsibleErrorDialog(
                    self, APP_TITLE,
                    f"'{s.get('name')}' 세트를 시작하지 못했습니다: 풀이 사본을 "
                    f"만들 수 없습니다 ({e})", text)
                return
            if practice:
                minutes = int(practice.get("minutes")
                              or practice_minutes([]))
            else:
                minutes = int(self.minutes_var.get() or DEFAULT_MINUTES)
            exam = {
                "set": s,
                "student": student,
                "minutes": max(MIN_MINUTES, min(MAX_MINUTES, minutes)),
                "started": datetime.now(),
                "practice_info": practice,
                "plan": day_plan,
                "excel_proc": None,
                "closed": False,
            }
            self._current_exam = exam
            self.exam_running = True
            routine_touch()
            self.start_btn.configure(state="disabled", text="진행 중")
            layout = self._plan_monitor_layout(exam)
            delay = 0
            if s.get("pdf"):
                before = top_level_windows() if layout.get("pdf") else []
                ok_pdf, err_pdf = open_file(s["pdf"])
                startup_log(f"PDF 실행: {'성공' if ok_pdf else '실패 ' + err_pdf}"
                            f" · {s['pdf']}")
                delay = EXCEL_OPEN_DELAY_MS
                if ok_pdf and layout.get("pdf"):
                    self.after(PDF_PLACE_DELAY_MS,
                               lambda: self._place_pdf_window(exam, before, 0))
            self.after(delay, lambda: self._launch_workbook(exam))
            TimerWindow(self, exam)

        def _plan_monitor_layout(self, exam=None):
            layout = {}
            try:
                mons = list_monitors()
                if monitor_ui_visible(mons):
                    picks = resolve_monitor_assignment(mons)
                    layout = {"monitors": mons,
                              "pdf": pick_monitor(mons, picks.get("pdf")),
                              "excel": pick_monitor(mons, picks.get("excel")),
                              "timer": picks.get("timer")}
                    startup_log(
                        "모니터 배치: 문제지="
                        f"{(layout['pdf'] or {}).get('label')} · Excel="
                        f"{(layout['excel'] or {}).get('label')}")
            except Exception as e:
                log_error("모니터 배치 계산", e)
                layout = {}
            if exam is not None:
                exam["monitor_layout"] = layout
            return layout

        def _place_excel_window(self, exam, tries=0):
            if exam.get("closed"):
                return False
            mon = (exam.get("monitor_layout") or {}).get("excel")
            if mon is None:
                return False
            try:
                moved, why = place_excel_on_monitor(
                    mon, getattr(exam.get("excel_proc"), "pid", None))
            except Exception as e:
                log_error("Excel 창 배치", e)
                return False
            if moved:
                startup_log(f"Excel 창 배치: {mon.get('label')}")
                return True
            if tries + 1 < EXCEL_PLACE_TRIES:
                self.after(PLACE_RETRY_MS,
                           lambda: self._place_excel_window(exam, tries + 1))
                return False
            startup_log(f"Excel 창 배치 실패: {why}")
            return False

        def _place_pdf_window(self, exam, before, tries=0):
            if exam.get("closed"):
                return False
            mon = (exam.get("monitor_layout") or {}).get("pdf")
            if mon is None:
                return False
            try:
                moved, why = place_new_window_on_monitor(mon, before)
            except Exception as e:
                log_error("문제지 창 배치", e)
                return False
            if moved:
                startup_log(f"문제지 창 배치: {mon.get('label')}")
                return True
            if tries + 1 < PDF_PLACE_TRIES:
                self.after(PLACE_RETRY_MS,
                           lambda: self._place_pdf_window(exam, before,
                                                          tries + 1))
                return False
            startup_log(f"문제지 창 배치 실패: {why}")
            self.show_toast(
                f"문제지 창은 자동으로 옮기지 못했습니다 — 창을 직접 "
                f"{mon.get('label')} 쪽으로 끌어다 놓으세요.", seconds=8)
            return False

        def _launch_workbook(self, exam, prefer_excel=True, verify=True,
                             source="시험 시작"):
            student = exam.get("student") or ""
            if not student or not os.path.isfile(student):
                err = f"풀이 파일이 없습니다: {student or '(경로 없음)'}"
                startup_log(f"Excel 실행({source}): 실패 · {err}")
                messagebox.showwarning(APP_TITLE, err, parent=self)
                return False, "", err
            ok, method, err, proc = open_workbook(student,
                                                  prefer_excel=prefer_excel)
            exam["excel_proc"] = proc
            startup_log(f"Excel 실행({source}): {'성공' if ok else '실패'} · "
                        f"방법={method}" + (f" · 오류={err}" if err else "")
                        + f" · 파일={student}")
            if not ok:
                self._show_excel_warning(exam, f"실행 실패 — {err}")
                return False, method, err
            if verify:
                self.after(EXCEL_CHECK_DELAY_MS,
                           lambda: self._verify_excel(exam, 0))
            if (exam.get("monitor_layout") or {}).get("excel"):
                self.after(EXCEL_PLACE_DELAY_MS,
                           lambda: self._place_excel_window(exam, 0))
            return True, method, err

        def _verify_excel(self, exam, tries):
            if exam.get("closed") or not exam.get("student"):
                return
            try:
                status, detail = check_workbook_open(exam["student"],
                                                     exam.get("excel_proc"))
            except Exception as e:
                log_error("Excel 실행 확인", e)
                return
            startup_log(f"Excel 확인({tries + 1}차): {status} · {detail}")
            if status in ("open", "unknown"):
                return
            if tries < 1:
                self.after(3000, lambda: self._verify_excel(exam, tries + 1))
                return
            self._show_excel_warning(exam, detail)

        def _show_excel_warning(self, exam, detail):
            w = self._excel_warn
            if w is not None and w.winfo_exists():
                w.update_detail(detail)
                w.lift()
                return w
            self._excel_warn = ExcelWarnWindow(self, exam, detail)
            return self._excel_warn

        def reopen_student(self, exam, prefer_excel=True):
            return self._launch_workbook(
                exam, prefer_excel=prefer_excel, verify=True,
                source="다시 열기(Excel 직접)" if prefer_excel
                else "기본 프로그램으로 열기")

        def show_startup_log(self):
            return DiagnosisWindow(
                self, read_startup_log(),
                title=f"{APP_TITLE} - 시작 로그",
                hint=("프로그램 시작·세트 선택·사본 생성·Excel/PDF 실행 결과가 "
                      f"시간순으로 기록된 파일입니다 ({STARTUP_LOG_PATH}). "
                      "[클립보드 복사]해서 채팅에 붙여넣으면 진단해 드립니다."))

        def exam_closed(self):
            self.exam_running = False
            routine_touch()
            exam = self._current_exam
            if exam is not None:
                exam["closed"] = True
            self._current_exam = None
            try:
                TimerOverlay.close_all()
            except Exception:
                pass
            w = self._excel_warn
            if w is not None and w.winfo_exists():
                try:
                    w.destroy()
                except Exception:
                    pass
            self._excel_warn = None
            self.start_btn.configure(state="normal", text="시험 시작")
            self.refresh_records()
            self._show_info()
            if self.updater.pending:
                self.after(500, self._apply_update_now)

        def plan_for(self, no):
            try:
                return adaptive_day_plan(self.adaptive, no)
            except Exception as e:
                log_error("적응형 일정 조회", e)
                return plan_for_day(no)

        def recompute_plan(self):
            try:
                self.adaptive = build_adaptive_plan(
                    date.today(), load_records(), self.sets)
            except Exception as e:
                log_error("적응형 일정 계산", e)
                self.adaptive = None
            if getattr(self, "plan_title_lbl", None) is not None:
                self._render_plan_card()
            return self.adaptive

        def _render_plan_card(self):
            plan = self.plan_for(self.plan_no)
            names = None
            reasons = []
            if plan.get("세트"):
                tag = plan_day_tag(plan["no"])
                try:
                    saved = load_auto_picks(tag) if tag else []
                    slot_sets = resolve_day_sets(plan, self.sets,
                                                 load_records(), saved=saved)
                except Exception:
                    slot_sets = [(None, "") for _ in plan["세트"]]
                names = [(s["name"] if s else
                          (spec if spec != AUTO else "자동 선택(세트 없음)"))
                         for (s, _r), spec in zip(slot_sets, plan["세트"])]
                reasons = [(s["name"] if s else "(없음)") + " — " + r
                           for (s, r), spec in zip(slot_sets, plan["세트"])
                           if spec == AUTO]
            self.plan_title_lbl.configure(
                text=plan_title(plan, today=date.today(), set_names=names))
            todo = plan.get("할일", "")
            if plan.get("목표") and not plan.get("적응형"):
                todo += f"  [목표 {plan['목표']}점]"
            if reasons:
                todo += "\n자동 선택: " + " / ".join(reasons)
            if plan.get("사유"):
                todo += "\n진행: " + plan["사유"]
            if plan.get("안내"):
                todo += "\n안내: " + plan["안내"]
            steps = plan.get("스텝")
            if steps:
                tag = plan_day_tag(plan["no"])
                done = (load_step_progress(tag) if tag else set())
                req = mandatory_step_indexes(steps)
                n_done = len([i for i in req if i in done])
                todo += f"  [필수 {n_done}/{len(req)} 완료]"
            self.plan_todo_lbl.configure(text=todo)
            today_no = routine_day_no()
            self.plan_start_btn.configure(
                text="오늘 일정 시작" if self.plan_no == today_no
                else "이 일정 시작")
            if not exam_date_set():
                self.plan_todo_lbl.configure(text=EXAM_UNSET_TODO)
            elif not self.sets:
                self.plan_todo_lbl.configure(
                    text=NO_SETS_REASON + "\n" + set_intake_help(self.scan_root))
            try:
                if get_app_setting("신뢰위치등록"):
                    self.trust_hint_lbl.pack_forget()
                else:
                    self.trust_hint_lbl.pack(fill="x", pady=(2, 0))
            except Exception:
                pass

        def on_register_trust(self):
            if not messagebox.askyesno(
                    APP_TITLE,
                    "이 학습 폴더를 Excel '신뢰할 수 있는 위치'로 등록하면 "
                    "매크로 차단 배너가 사라집니다.\n\nExcel 보안 설정"
                    "(레지스트리 HKCU)을 수정합니다. 진행할까요?\n\n"
                    f"등록 폴더: {self.scan_root}", parent=self):
                return
            ok, msg = register_trusted_location(self.scan_root)
            if ok:
                set_app_setting("신뢰위치등록", True)
                self._render_plan_card()
                messagebox.showinfo(APP_TITLE, msg, parent=self)
            else:
                messagebox.showwarning(APP_TITLE, msg, parent=self)

        def _shift_plan(self, delta):
            order = PLAN_ORDER
            i = order.index(self.plan_no) if self.plan_no in order else 0
            self.plan_no = order[max(0, min(len(order) - 1, i + delta))]
            self._render_plan_card()

        def add_set_to_list(self, s):
            if not s or not s.get("norm"):
                return False
            for x in self.sets:
                if x.get("norm") == s.get("norm"):
                    return False
            self.sets.append(s)
            mark = "[직접] " if s.get("direct") else ""
            pdf_mark = "" if s.get("pdf") else "  (문제지 미연결)"
            self.listbox.insert("end", f" {mark}{s['name']}{pdf_mark}")
            return True

        def show_scan_diagnosis(self):
            try:
                text = scan_diagnosis_text(
                    self.scan_root, sets=self.sets,
                    hidden_sets=getattr(self, "hidden_sets", []))
            except Exception as e:
                log_error("세트 인식 진단", e)
                text = f"진단 생성 중 오류: {e}"
            DiagnosisWindow(self, text)

        def _select_set_in_list(self, s):
            for i, x in enumerate(self.sets):
                if x.get("norm") == s.get("norm"):
                    self.listbox.selection_clear(0, "end")
                    self.listbox.selection_set(i)
                    self.listbox.see(i)
                    self._show_info()
                    return True
            return False

        def start_today_plan(self):
            plan = self.plan_for(self.plan_no)
            if not plan.get("스텝"):
                messagebox.showinfo(f"{APP_TITLE} - {plan['제목']}",
                                    plan["할일"], parent=self)
                return
            g = getattr(self, "step_guide", None)
            if g is not None and g.winfo_exists():
                if g.plan["no"] == plan["no"]:
                    g.lift()
                    g.focus_force()
                    return
                g.destroy()
            self.step_guide = StepGuideWindow(self, plan)

        def open_practice_dialog(self):
            if self.exam_running:
                messagebox.showinfo(APP_TITLE, "이미 시험이 진행 중입니다.",
                                    parent=self)
                return
            s = self._selected_set()
            if not s:
                messagebox.showinfo(APP_TITLE, "먼저 세트를 선택하세요.",
                                    parent=self)
                return
            dlg = PracticeDialog(self)
            self.wait_window(dlg)
            if dlg.result:
                self.start_exam(practice=dlg.result)

        def open_review_mode(self, full_only=False, json_path=None,
                             set_info=None):
            s = set_info or self._selected_set()
            if not s:
                messagebox.showinfo(APP_TITLE, "먼저 세트를 선택하세요.",
                                    parent=self)
                return
            jp = json_path if json_path and os.path.isfile(json_path) \
                else find_latest_result_json(s, full_only=full_only)
            if not jp:
                messagebox.showinfo(
                    APP_TITLE, "이 세트의 채점 기록(채점결과 JSON)이 "
                    "없습니다.\n먼저 시험 모드로 응시해 채점을 받으세요.",
                    parent=self)
                return
            data, items = load_wrong_items(jp)
            if data is None:
                messagebox.showerror(
                    APP_TITLE, f"채점결과 파일을 읽을 수 없습니다:\n{jp}",
                    parent=self)
                return
            if not items:
                messagebox.showinfo(
                    APP_TITLE, f"최근 채점({data.get('total', '?')}점)에 "
                    "오답 항목이 없습니다. 복습할 내용이 없어요!",
                    parent=self)
                return
            note = None
            src = find_latest_attempt(s)
            if not src:
                cand = (data.get("files") or {}).get("student")
                src = cand if cand and os.path.isfile(cand) else None
            if src:
                try:
                    copy = make_review_copy(src, s["name"])
                    open_file(copy)
                except OSError as e:
                    note = f"이전 풀이 사본을 열지 못했습니다: {e}"
            else:
                note = ("이전 풀이 파일(풀이_*.xlsx)을 찾지 못해 해설만 "
                        "표시합니다.")
            record_review_state(
                s.get("norm") or "", os.path.basename(jp),
                load_review_state(s.get("norm") or "", os.path.basename(jp)),
                count_up=True)
            ReviewWindow(self, s, jp, items, note=note)

        def start_grading(self, exam, elapsed_seconds, practice, on_done):
            stamp = datetime.now().strftime("%Y%m%d_%H%M")
            out_dir = os.path.join(
                os.path.dirname(os.path.abspath(exam["student"])), "채점결과")
            try:
                os.makedirs(out_dir, exist_ok=True)
            except OSError:
                out_dir = exam["set"]["dir"]
            base = os.path.join(
                out_dir, f"채점결과_{exam['set']['name']}_{stamp}")
            html_path = base + ".html"
            json_path = base + ".json"
            state = {"done": False}

            pinfo = exam.get("practice_info")

            def do_run():
                return run_grading(
                    self.grade_py, exam["set"]["problem"],
                    exam["set"]["answer"], exam["student"],
                    key=exam["set"]["key"], html=html_path,
                    json_out=json_path, history=RECORDS_PATH,
                    sheets=pinfo.get("sheets") if pinfo else None)

            state["runner"] = do_run

            def worker():
                state["result"], state["stdout"], state["stderr"], \
                    state["rc"] = do_run()
                state["done"] = True

            threading.Thread(target=worker, daemon=True).start()

            def poll():
                if not state["done"]:
                    self.after(200, poll)
                    return
                self._grading_finished(exam, elapsed_seconds, practice,
                                       state, html_path)
                on_done()

            poll()

        def _copy_result_to_clipboard(self, result):
            try:
                payload = json.dumps(result, ensure_ascii=False,
                                     separators=(",", ":"))
                self.clipboard_clear()
                self.clipboard_append(payload)
                self.update_idletasks()
                return True
            except Exception:
                return False

        def _install_with_progress(self):
            prog = tk.Toplevel(self)
            prog.title(APP_TITLE)
            prog.attributes("-topmost", True)
            tk.Label(prog, text="openpyxl 설치 중입니다... 잠시 기다려 주세요.",
                     padx=26, pady=18, font=UI_FONT).pack()
            prog.update()
            ok, log = install_openpyxl()
            prog.destroy()
            return ok, log

        def _show_result_window(self, record, prev_records, show, html_path):
            prog = None
            try:
                sets = self.sets
                before = serialize_records(prev_records, sets)[0]
                after = serialize_records(prev_records + [record], sets)[0]
                prog = progress_delta(before, after)
            except Exception as e:
                log_error("결과 연출 계산", e)
            ResultWindow(self, show["result"], html_path,
                         folder=os.path.dirname(html_path),
                         copied=show["copied"], goal=show["goal"],
                         linked=self.routine_alive(), progress=prog)

        def _grading_finished(self, exam, elapsed, practice, state, html_path):
            result = state.get("result")
            combined = ((state.get("stderr") or "") + "\n"
                        + (state.get("stdout") or "")).strip()
            if result is None and classify_grading_error(
                    state.get("rc"), combined) == "module":
                if messagebox.askyesno(
                        APP_TITLE, "채점에 필요한 openpyxl이 없습니다.\n"
                        "지금 설치할까요? (인터넷 필요)", parent=self):
                    ok, log = self._install_with_progress()
                    if ok and state.get("runner"):
                        result, out2, err2, rc2 = state["runner"]()
                        state["rc"] = rc2
                        combined = ((err2 or "") + "\n" + (out2 or "")).strip()
                    elif not ok:
                        CollapsibleErrorDialog(
                            self, APP_TITLE,
                            "openpyxl 설치에 실패했습니다.\n\n" + MANUAL_PIP,
                            log)
            if result is None:
                kind = classify_grading_error(state.get("rc"), combined)
                if kind == "module":
                    headline = ("채점에 필요한 openpyxl이 설치되어 있지 않아 "
                                "채점하지 못했습니다.\n\n" + MANUAL_PIP)
                elif kind == "file":
                    headline = ("채점용 파일을 열 수 없습니다.\n"
                                "문제/정답/풀이 파일의 경로와 형식"
                                "(.xlsx/.xlsm)이 올바른지 확인하세요.")
                else:
                    headline = "채점에 실패했습니다."
                CollapsibleErrorDialog(
                    self, f"{APP_TITLE} - 채점 실패", headline, combined)
                score = None
                show = None
            else:
                score = result.get("total")
                show = {"result": result,
                        "copied": self._copy_result_to_clipboard(result),
                        "goal": (exam.get("plan") or {}).get("목표")}
            prev_records = load_records()
            pinfo = exam.get("practice_info")
            record = {
                "일시": exam["started"].strftime("%Y-%m-%d %H:%M"),
                "세트명": exam["set"]["name"] + (" (연습)" if practice else ""),
                "점수": score,
                "소요시간": format_elapsed(elapsed),
                "소요초": max(0, int(elapsed)),
                "리포트": html_path if os.path.isfile(html_path) else None,
                "mode": (pinfo.get("mode") or "부분연습") if pinfo else "시험",
            }
            try:
                record.update(grading_evidence(
                    problem=exam["set"].get("problem"),
                    answer=exam["set"].get("answer"),
                    student=exam.get("student"),
                    key=exam["set"].get("key"),
                    grade_py=self.grade_py))
            except Exception as e:
                log_error("채점 근거 저장", e)
            if pinfo:
                if pinfo.get("mode") == "오답재풀이":
                    record["영역"] = "오답재풀이(" + ",".join(
                        pinfo.get("sheets") or []) + ")"
                else:
                    record["영역"] = pinfo.get("label")
                if result:
                    record["만점"] = result.get("max_total")
            day_tag = (exam.get("plan") or {}).get("day")
            if day_tag:
                record["day"] = day_tag
                record["루틴"] = routine_tag()
            try:
                append_record(record)
            except OSError as e:
                messagebox.showwarning(
                    APP_TITLE, f"기록.json 저장에 실패했습니다: {e}",
                    parent=self)
            if show is not None:
                self._show_result_window(record, prev_records, show, html_path)
                if os.path.isfile(html_path):
                    open_file(html_path)
            plan_info = exam.get("plan") or {}
            step_idx = plan_info.get("step")
            done_steps = plan_info.get("done_steps") or (
                [step_idx] if step_idx is not None else [])
            if day_tag and done_steps and score is not None:
                try:
                    g = getattr(self, "step_guide", None)
                    if g is not None and g.winfo_exists() \
                            and g.day_tag == day_tag:
                        for idx in done_steps:
                            g.mark_step_done(int(idx))
                    else:
                        done = load_step_progress(day_tag)
                        done.update(int(idx) for idx in done_steps)
                        save_step_progress(day_tag, done)
                    self._render_plan_card()
                except Exception:
                    pass

def run_smoke():
    app = ExamApp(auto_update=False)
    app.update_idletasks()
    app.update()
    exam = {"set": {"name": "스모크테스트", "dir": BASE_DIR,
                    "problem": "", "answer": "", "key": None, "pdf": None},
            "student": "", "minutes": 40, "started": datetime.now()}
    timer = TimerWindow(app, exam)
    app.update_idletasks()
    app.update()
    assert timer.time_lbl.cget("text") == "40:00"
    assert app.listbox is not None
    _ov_before = timer_overlay_enabled()
    if _ov_before and timer.overlay is not None:
        app.update_idletasks()
        app.update()
        ov = timer.overlay
        assert ov.winfo_exists() and ov.time_lbl.cget("text") == "40:00"
        assert ov.head_lbl.cget("text") == "스모크테스트"
        _mons = list_monitors()
        _want = (overlay_position(_mons, resolve_monitor_assignment(
            _mons).get("timer"), size=ov._size(), saved=saved_overlay_pos())
            if _mons else overlay_position_for_screen(
                ov.winfo_screenwidth(), ov.winfo_screenheight(),
                size=ov._size(), saved=saved_overlay_pos()))
        assert (ov.winfo_x(), ov.winfo_y()) == _want, \
            f"오버레이 자리 {(ov.winfo_x(), ov.winfo_y())} != {_want}"
        assert ov.update_time("05:00", "#FF8B80", "남은 시간") == "05:00"
        timer.close_overlay()
        app.update_idletasks()
        app.update()
        assert timer.overlay is None and not ov.winfo_exists()
        assert TimerOverlay.close_all() == 0
        timer._open_overlay("스모크테스트")
        app.update_idletasks()
        app.update()
    items = [{"sheet": "계산작업", "label": "계산 문제 1", "lost": 8,
              "cells": [{"coord": "J3", "got": "#VALUE!", "expected": "본선",
                         "formula": "=IF(OR(...),\"본선\",\"\")",
                         "got_formula": "=OR(IF(...))"}],
              "diff_notes": ["중첩 순서가 반대입니다"],
              "explain": ["① SMALL(...): ..."], "hint": "확인"}]
    review = ReviewWindow(app, exam["set"],
                          os.path.join(BASE_DIR, "가짜.json"), items)
    app.update_idletasks()
    app.update()
    assert "계산 문제 1" in review.listbox.get(0)
    review.checks = set()
    review._refresh_list()
    review.listbox.selection_set(0)
    review.toggle_check()
    assert review.retake_btn.cget("state") == "normal"
    review.destroy()
    dlg = PracticeDialog(app)
    app.update_idletasks()
    app.update()
    dlg.apply_preset(["계산작업"])
    sheets, label = dlg.selection()
    assert sheets == ["계산작업"] and label == "계산작업"
    assert dlg.minutes_var.get() == 15
    dlg.apply_preset(["기본작업"])
    sheets, label = dlg.selection()
    assert sheets == ["기본작업-1", "기본작업-2", "기본작업-3"]
    dlg.destroy()
    smoke_exam = date.today() + timedelta(days=5)
    save_exam_dates(smoke_exam, None)
    assert exam_dates() == [smoke_exam], exam_dates()
    app.plan_no = routine_day_no()
    app._render_plan_card()
    assert app.plan_title_lbl.cget("text"), "오늘의 학습 제목 비어 있음"
    app.plan_no = 3
    app._render_plan_card()
    t3 = app.plan_title_lbl.cget("text")
    d3 = routine_date_for(3)
    assert "Day 3 (D-5)" in t3 and f"{d3.month}/{d3.day}(" in t3, t3
    assert any(k in t3 for k in ("오늘", "예정", "완료", "응시 없음",
                                 "첫 응시", "재응시", "자유 복습",
                                 "세트 없음", "Day 3")), t3
    assert app.recalc_btn.cget("text") == "일정 다시 계산"
    assert app.exam_date_btn.cget("text") == "시험일 설정"
    ad = app.recompute_plan()
    assert ad is None or ad.get("kind") in (
        "unset", "free", "week", "exam", "after"), ad
    if ad and ad.get("kind") == "week":
        today_plan = app.plan_for(routine_day_no())
        assert today_plan.get("적응형") is True
        if today_plan.get("스텝"):
            gA = StepGuideWindow(app, today_plan)
            app.update_idletasks()
            app.update()
            assert gA.listbox.size() == len(today_plan["스텝"])
            gA.destroy()
    assert "시험 D-" in t3, t3
    app.plan_no = 3
    app._render_plan_card()
    app._shift_plan(1)
    assert "Day 4" in app.plan_title_lbl.cget("text")
    app.plan_no = ROUTINE_DAYS
    app._shift_plan(1)
    assert app.plan_no == PLAN_EXAM1
    assert "시험일" in app.plan_title_lbl.cget("text")
    dlg_e = ExamDateDialog(app)
    app.update_idletasks()
    app.update()
    assert dlg_e.values() == [smoke_exam.isoformat()], dlg_e.values()
    second = smoke_exam + timedelta(days=3)
    assert dlg_e.add(second.isoformat()) is True
    assert dlg_e.add(second.isoformat()) is False, "중복은 안 들어간다"
    assert dlg_e.add("내일") is False, "형식 오류는 안 들어간다"
    assert dlg_e.values() == [smoke_exam.isoformat(), second.isoformat()], \
        dlg_e.values()
    assert dlg_e.remove(second.isoformat()) is True
    assert dlg_e.values() == [smoke_exam.isoformat()]
    assert dlg_e.clear_all() is True and dlg_e.values() == []
    assert dlg_e.validate() == ([], ""), "시험일 0개도 정상"
    dlg_e.add(smoke_exam.isoformat())
    assert dlg_e.save() is True and exam_dates() == [smoke_exam]
    assert UPDATE_REPO_URL in first_run_notice_text()
    assert "동의" in first_run_notice_text() and "거절" in first_run_notice_text()
    _au_saved = auto_update_enabled()
    dlg_c = UpdateConsentDialog(app)
    app.update_idletasks()
    app.update()
    assert dlg_c.agreed is False and dlg_c.answered is False, "기본값은 거절"
    dlg_c.decline()
    assert dlg_c.agreed is False and dlg_c.answered is True
    assert apply_update_consent(False) is False \
        and auto_update_enabled() is False and not first_run_notice_pending()
    assert apply_update_consent(True) is True and auto_update_enabled() is True
    set_app_setting(AUTO_UPDATE_SETTING, _au_saved)
    dlg_s = app.open_settings()
    app.update_idletasks()
    app.update()
    assert dlg_s is app.open_settings(), "설정 창은 하나만"
    assert app.scan_root in dlg_s.root_lbl.cget("text")
    assert dlg_s.reset_root() == default_scan_root()
    assert scan_root_is_default() is True
    dlg_s.destroy()
    save_step_progress("d03", set())
    guide = StepGuideWindow(app, plan_for_day(3))
    app.update_idletasks()
    app.update()
    assert guide.listbox.size() == 5, guide.listbox.size()
    assert guide._required_steps() == [0, 1], guide._required_steps()
    assert "필수 0/2" in guide.progress_lbl.cget("text")
    assert guide.current_index() == 0
    assert "▶" in guide.listbox.get(0)
    assert "세트:" in guide.pick_lbl.cget("text"), guide.pick_lbl.cget("text")
    guide.listbox.selection_clear(0, "end")
    guide.listbox.selection_set(0)
    guide.toggle_check()
    assert 0 in guide.done and "필수 1/2" in guide.progress_lbl.cget("text")
    assert guide.selected_index() == 1, "다음 미완료 스텝 자동 포커스"
    assert "[v]" in guide.listbox.get(0) and "▶" in guide.listbox.get(1)
    _orig_info = messagebox.showinfo
    _cleared = []
    messagebox.showinfo = lambda *a, **k: _cleared.append(a)
    try:
        guide.mark_step_done(1)
    finally:
        messagebox.showinfo = _orig_info
    assert "필수 2/2" in guide.progress_lbl.cget("text")
    assert guide._all_done() is True, "필수만 끝나도 클리어"
    assert _cleared and "클리어" in _cleared[0][1], _cleared
    assert load_step_progress("d03") == {0, 1}, "진행 상태 저장"
    guide.destroy()
    guide2 = StepGuideWindow(app, plan_for_day(3))
    app.update_idletasks()
    app.update()
    assert guide2.done == {0, 1} and guide2.current_index() == 2
    guide2.destroy()
    save_step_progress("d03", set())
    fake_s = {"name": "2024년 상시1회 2급", "norm": "smoke상시1", "dir": BASE_DIR,
              "problem": "p", "answer": "a", "key": None, "pdf": None}
    fake_s2 = {"name": "코코모의고사 2회", "norm": "smoke코코2", "dir": BASE_DIR,
               "problem": "p2", "answer": "a2", "key": None, "pdf": None}
    saved_sets, saved_start = app.sets, app.start_exam
    calls = []
    app.sets = [fake_s, fake_s2]
    app.listbox.delete(0, "end")
    for s in app.sets:
        app.listbox.insert("end", " " + s["name"])
    app.start_exam = lambda practice=None: calls.append(
        (practice, app._pending_plan))
    save_step_progress("d01", set())
    guide3 = StepGuideWindow(app, dict(plan_for_day(1),
                                       세트=[fake_s["name"]]))
    assert guide3.slot_sets[0][0] is fake_s, guide3.pick_text()
    assert "일정 지정" in guide3.pick_lbl.cget("text")
    guide3.listbox.selection_clear(0, "end")
    guide3.listbox.selection_set(0)
    guide3.start_step()
    assert len(calls) == 1, "모의 스텝 -> start_exam 호출"
    practice_arg, pend = calls[0]
    assert practice_arg is None and pend == {
        "day": "d01", "목표": 70, "step": 0, "done_steps": [0, 1]}, \
        (practice_arg, pend)
    assert app.minutes_var.get() == 40
    guide3.destroy()
    save_step_progress("d01", set())
    save_step_progress("d05", set())
    save_auto_picks("d05", [])
    guide4 = StepGuideWindow(app, plan_for_day(5))
    app.update_idletasks()
    app.update()
    ptxt = guide4.pick_lbl.cget("text")
    assert "자동 선택" in ptxt and "신규 세트라서" in ptxt, ptxt
    assert guide4.change_btn.winfo_manager(), "[다른 세트로 바꾸기] 표시"
    assert guide4.steps[0]["이름"].startswith(guide4.slot_sets[0][0]["name"])
    picked_before = guide4.slot_sets[0][0]["name"]
    other = fake_s2 if picked_before == fake_s["name"] else fake_s
    chooser = SetChooserDialog(guide4, app.sets, 1, guide4.apply_set_change)
    app.update_idletasks()
    app.update()
    chooser.listbox.selection_set(app.sets.index(other))
    chooser.apply()
    assert guide4.slot_sets[0][0] is other
    assert "직접 선택" in guide4.pick_lbl.cget("text")
    assert guide4.steps[0]["이름"].startswith(other["name"]), guide4.steps[0]
    assert load_auto_picks("d05")[0]["세트"] == other["name"], "선택 저장"
    guide4.destroy()
    save_auto_picks("d05", [])
    save_step_progress("d05", set())
    app.sets, app.start_exam = saved_sets, saved_start
    app._pending_plan = None
    fake_s2r = {"name": "2024년 상시2회 2급", "norm": "smoke상시2", "dir": BASE_DIR,
                "problem": "p3", "answer": "a3", "key": None, "pdf": None}
    fake_A = {"name": "컴활2급 A형", "norm": "smokeA", "dir": BASE_DIR,
              "problem": os.path.join(BASE_DIR, "시험장.py"),
              "answer": os.path.join(BASE_DIR, "시험장.py"), "key": None,
              "pdf": None, "direct": True}
    saved_sets, saved_start = app.sets, app.start_exam
    calls = []
    app.sets = [fake_s2r]
    app.listbox.delete(0, "end")
    app.listbox.insert("end", " " + fake_s2r["name"])
    app.start_exam = lambda practice=None: calls.append(
        (practice, app._pending_plan))
    cfg_before = load_set_config()
    cfg_before.pop("_슬롯매핑", None)
    save_set_config(cfg_before)
    save_step_progress("d03", set())
    _plan2 = dict(plan_for_day(3), 세트=["2024 A형", "2024 상시 2회"],
                  목표들=[70, 70])
    _plan2["스텝"] = build_day_steps(_plan2)
    assert len(_plan2["스텝"]) == 9, len(_plan2["스텝"])
    guide5 = StepGuideWindow(app, _plan2)
    app.update_idletasks()
    app.update()
    assert guide5.missing_slots() == [0], guide5.pick_text()
    assert guide5.slot_sets[1][0] is fake_s2r
    assert len(guide5.missing_rows) == 1 and \
        guide5.missing_rows[0][2].cget("text") == "직접 선택"
    assert "찾지 못함" in guide5.pick_text() and \
        guide5.pick_lbl.cget("fg") == RED
    guide5.listbox.selection_clear(0, "end")
    guide5.listbox.selection_set(0)
    _orig_info = messagebox.showinfo
    infos = []
    messagebox.showinfo = lambda *a, **k: infos.append(a)
    try:
        guide5.start_step()
    finally:
        messagebox.showinfo = _orig_info
    assert calls == [], "미발견 슬롯은 시작되지 않음"
    assert infos and "다른 세트로 바꾸기" in infos[0][1], infos
    guide5.listbox.selection_clear(0, "end")
    guide5.listbox.selection_set(4)
    guide5.start_step()
    assert len(calls) == 1 and calls[0][1]["step"] == 4, "정상 슬롯은 시작"
    assert app._selected_set() is fake_s2r
    guide5.apply_direct_slot(0, fake_A)
    assert guide5.missing_slots() == [] and not guide5.missing_rows
    assert guide5.steps[0]["이름"].startswith("컴활2급 A형")
    assert load_slot_mapping()["2024 A형"]["name"] == "컴활2급 A형", "매핑 저장"
    assert any(x["norm"] == "smokeA" for x in app.sets), "목록 편입"
    guide5.listbox.selection_clear(0, "end")
    guide5.listbox.selection_set(0)
    guide5.start_step()
    assert len(calls) == 2 and app._selected_set() is fake_A
    guide5.destroy()
    guide6 = StepGuideWindow(app, plan_for_day(3))
    _orig_msg = messagebox.showerror
    shown = []
    messagebox.showerror = lambda *a, **k: shown.append(a)
    guide6._start_step = lambda i, st: (_ for _ in ()).throw(RuntimeError("주입"))
    guide6.listbox.selection_set(0)
    guide6.start_step()
    messagebox.showerror = _orig_msg
    assert shown and "시작하지 못했습니다: 주입" in shown[0][1], shown
    assert guide6.winfo_exists()
    guide6.destroy()
    cfg_after = load_set_config()
    cfg_after.pop("_슬롯매핑", None)
    cfg_after.pop("smokeA", None)
    save_set_config(cfg_after)
    save_step_progress("d03", set())
    app.sets, app.start_exam = saved_sets, saved_start
    app._pending_plan = None
    log_before = os.path.getsize(ERROR_LOG_PATH) \
        if os.path.isfile(ERROR_LOG_PATH) else 0
    _LAST_ERROR_DIALOG[0] = None
    boom = tk.Button(app, command=lambda: 1 / 0)
    boom.invoke()
    app.update_idletasks()
    app.update()
    assert _LAST_ERROR_DIALOG[0] is not None, "오류 대화상자 생성"
    dlg = _LAST_ERROR_DIALOG[0]
    assert "ZeroDivisionError" in dlg.copy_text()
    assert dlg.copy_error() is True and dlg.copy_btn.cget("text") == "복사됨"
    dlg.destroy()
    boom.destroy()
    with open(ERROR_LOG_PATH, encoding="utf-8") as f:
        f.seek(log_before)
        tail = f.read()
    assert "ZeroDivisionError" in tail and "tk callback" in tail, tail[-200:]
    diag = DiagnosisWindow(app, scan_diagnosis_text(
        app.scan_root, sets=app.sets,
        hidden_sets=getattr(app, "hidden_sets", [])))
    app.update_idletasks()
    app.update()
    assert "일정 슬롯 매칭" in diag.text_value and diag.copy() is True
    diag.destroy()
    startup_log("스모크 테스트 줄")
    logw = app.show_startup_log()
    app.update_idletasks()
    app.update()
    assert "시작 로그" in logw.title() and "스모크 테스트 줄" in logw.text_value
    logw.destroy()
    assert timer.reopen_btn.cget("text") == "풀이 파일 열기"
    warn = app._show_excel_warning(exam, "테스트 상세")
    app.update_idletasks()
    app.update()
    assert warn.winfo_exists() and "테스트 상세" in warn.detail_lbl.cget("text")
    assert [b.cget("text") for b in warn.buttons] == [
        "다시 열기(Excel 직접)", "기본 프로그램으로 열기", "폴더 열기"]
    assert app._show_excel_warning(exam, "갱신") is warn
    assert "갱신" in warn.detail_lbl.cget("text")
    _orig_warn = messagebox.showwarning
    warned = []
    messagebox.showwarning = lambda *a, **k: warned.append(a)
    try:
        ok_re, _m, err_re = app.reopen_student(exam)
    finally:
        messagebox.showwarning = _orig_warn
    assert ok_re is False and "풀이 파일이 없습니다" in err_re and warned
    app._current_exam = exam
    app.exam_closed()
    assert exam["closed"] is True and not warn.winfo_exists()
    app._verify_excel(exam, 0)
    bad = {"name": "2024년 상시2회 2급", "norm": "smoke상시2회", "dir": BASE_DIR,
           "problem": "p", "answer": "a", "key": None,
           "pdf": os.path.join(BASE_DIR, "2024 상시1회 문제지.pdf")}
    saved_sets = app.sets
    app.sets = [bad]
    app.listbox.delete(0, "end")
    app.listbox.insert("end", " x")
    app.listbox.selection_set(0)
    app._show_info()
    assert app.pdf_warn_lbl.winfo_manager(), "충돌 경고 표시"
    assert PDF_MISMATCH_WARNING in app.pdf_warn_lbl.cget("text")
    bad["pdf"] = None
    bad["pdf_warning"] = "저장된 문제지 연결(x.pdf)의 회차·형·연도가 세트와 달라 무시했습니다"
    app._show_info()
    assert app.pdf_warn_lbl.winfo_manager() and \
        "무시했습니다" in app.pdf_warn_lbl.cget("text")
    bad["pdf_warning"] = None
    app._show_info()
    assert not app.pdf_warn_lbl.winfo_manager()
    app.sets = saved_sets
    app.listbox.delete(0, "end")
    for s_ in app.sets:
        app.listbox.insert("end", " " + s_["name"])
    assert not app.toast.winfo_manager()
    app.show_toast("스모크 안내", seconds=0)
    app.update_idletasks()
    app.update()
    assert app.toast.winfo_manager() and \
        app.toast_lbl.cget("text") == "스모크 안내"
    app.hide_toast()
    assert not app.toast.winfo_manager()
    _au_before = auto_update_enabled()
    app.auto_update_var.set(False)
    assert app._toggle_auto_update() is False and auto_update_enabled() is False
    app.auto_update_var.set(True)
    assert app._toggle_auto_update() is True and auto_update_enabled() is True
    set_app_setting(AUTO_UPDATE_SETTING, _au_before)
    assert app._apply_update_now() == "none"
    assert app._on_update_checked("latest") == "latest"
    app._on_update_checked("failed")
    assert "자동 업데이트 실패" in app.toast_lbl.cget("text")
    app.hide_toast()
    app.updater.data_synced = 2
    app._on_update_checked("data")
    assert "기대값 파일 2개" in app.toast_lbl.cget("text")
    app.hide_toast()
    assert app._busy_for_restart() is True
    timer.withdraw()
    app.update_idletasks()
    app.update()
    assert app._busy_for_restart() is False
    _dw = DiagnosisWindow(app, "x")
    app.update_idletasks()
    app.update()
    assert app._busy_for_restart() is True
    _dw.destroy()
    app.update_idletasks()
    app.update()
    assert app._busy_for_restart() is False
    timer.deiconify()
    r_args, r_kw = restart_command(argv=["시험장.py"])
    assert r_args[0] == sys.executable and r_args[1].endswith("시험장.py") \
        and r_kw["cwd"] == BASE_DIR, (r_args, r_kw)
    kind, _p = resolve_step_action(plan_for_day(1)["스텝"][0], saved_sets)
    assert kind in ("exam", "missing")
    kind, _p = resolve_step_action(plan_for_day(1)["스텝"][1], saved_sets)
    assert kind == "info"
    kind, _p = resolve_step_action(plan_for_day(0)["스텝"][0], saved_sets)
    assert kind == "info"
    fake = {"total": 87, "mode": "full", "graded_sheets": ["계산작업"]}
    assert app._copy_result_to_clipboard(fake) is True
    back = app.clipboard_get()
    import json as _json
    assert _json.loads(back) == fake, back[:80]
    import urllib.request as _ur
    assert app.routine is not None and app.routine.running, "루틴 서버 기동"
    assert app.routine_btn.cget("text") == "루틴 열기"
    assert app.routine_auto_open_chk.cget("text") == "루틴 자동 열기"
    _url = app.routine_url()
    assert _url and _url.startswith(f"http://127.0.0.1:{app.routine.port}/?t=")
    with _ur.urlopen(f"http://127.0.0.1:{app.routine.port}/api/state",
                     timeout=5) as _resp:
        _st = _json.loads(_resp.read().decode("utf-8"))
        assert _resp.headers.get("Cache-Control") == "no-store"
    assert _st["today"] == date.today().isoformat() and "plan" in _st \
        and len(_st["plan"]["days"]) == 16, list(_st)
    with _ur.urlopen(f"http://127.0.0.1:{app.routine.port}/",
                     timeout=5) as _resp:
        _page = _resp.read().decode("utf-8")
    assert _page.lstrip().lower().startswith("<!doctype html>"), _page[:40]
    _req = _ur.Request(f"http://127.0.0.1:{app.routine.port}/api/check",
                       data=b'{"key":"smoke|quiz","value":true}',
                       headers={"Content-Type": "application/json"})
    try:
        _ur.urlopen(_req, timeout=5)
        raise AssertionError("토큰 없는 POST 가 통과")
    except _ur.HTTPError as e:
        assert e.code == 401, e.code
    _box = {}

    def _post_check():
        r = _ur.Request(f"http://127.0.0.1:{app.routine.port}/api/check",
                        data=b'{"key":"smoke|quiz","value":true}',
                        headers={"Content-Type": "application/json",
                                 "X-Coco-Token": app.routine.token})
        try:
            with _ur.urlopen(r, timeout=10) as resp:
                _box["r"] = _json.loads(resp.read().decode("utf-8"))
        except Exception as e:
            _box["e"] = e
    _t = threading.Thread(target=_post_check, daemon=True)
    _t.start()
    _deadline = time.time() + 10
    while _t.is_alive() and time.time() < _deadline:
        app.update()
        time.sleep(0.05)
    assert _box.get("r", {}).get("ok") is True, _box
    assert _box["r"]["checks"].get("smoke|quiz") is True
    _cfg = load_set_config()
    assert _cfg.get(WEB_CHECKS_KEY, {}).get("smoke|quiz") is True, "_웹체크 저장"
    _cfg.get(WEB_CHECKS_KEY, {}).pop("smoke|quiz", None)
    save_set_config(_cfg)
    save_step_progress("d03", set())
    g7 = StepGuideWindow(app, plan_for_day(3))
    app.update_idletasks()
    app.update()
    assert not g7.web_btn.winfo_manager()
    g7.listbox.selection_clear(0, "end")
    g7.listbox.selection_set(g7.listbox.size() - 1)
    g7._show_detail()
    assert g7.web_btn.winfo_manager() and g7.web_tab_of() == "quiz"
    _opened = []
    app.open_routine_page = lambda tab=None: _opened.append(tab)
    g7.start_step()
    assert _opened == ["quiz"], _opened
    g7.destroy()
    save_step_progress("d03", set())
    del app.open_routine_page
    rw = ResultWindow(app, {"total": 80, "pass_line": 70, "sheets": []}, "",
                      linked=True)
    app.update_idletasks()
    app.update()
    assert "자동 반영" in rw.link_lbl.cget("text")
    rw.destroy()
    _ser_before = [{"id": "20260903100000", "date": "2026-09-03",
                    "time": "10:00", "mode": "시험", "total": 72,
                    "set": {"name": "스모크", "norm": "스모크"},
                    "sheets": [], "wrong_items": [], "elapsed_sec": 2300}]
    _ser_after = _ser_before + [{
        "id": "20260904100000", "date": "2026-09-04", "time": "10:00",
        "mode": "시험", "total": 84, "set": {"name": "스모크2", "norm": "스모크2"},
        "sheets": [], "wrong_items": [], "elapsed_sec": 2350}]
    _prog = progress_delta(_ser_before, _ser_after)
    assert _prog["prev_best"] == 72 and _prog["best"] == 84         and _prog["best_delta"] == 12 and _prog["tier_up"] == [80], _prog
    assert [a["id"] for a in _prog["new_badges"]] == ["tier80"],         _prog["new_badges"]
    rw2 = ResultWindow(app, {"total": 84, "pass_line": 70, "sheets": []}, "",
                       goal=80, progress=_prog)
    app.update_idletasks()
    app.update()
    _texts = [w.cget("text") for w in rw2.winfo_children()[0].winfo_children()
              if isinstance(w, tk.Label)]
    assert any("80점 달성! 숙련 단계" in t for t in _texts), _texts
    assert any("직전 최고 72점 대비 +12점" in t for t in _texts), _texts
    assert rw2._gauge_target == 0.84 and rw2.score_gauge.winfo_exists()
    for _ in range(ResultWindow.STEPS + 6):
        app.update()
        time.sleep(0.04)
    assert rw2._gauge_step == ResultWindow.STEPS, rw2._gauge_step
    rw2.destroy()
    app._render_status_card([], ser=_ser_after)
    app.update_idletasks()
    app.update()
    assert [lbl.cget("text") for _g, lbl in app.tier_lbls][:2]         == ["✓ 합격 70", "✓ 숙련 80"], [lbl.cget("text")
                                       for _g, lbl in app.tier_lbls]
    assert "최고 84점 / 목표 90점(고득점)" in app.gauge_lbl.cget("text")
    assert "도전 과제" in app.badge_count_lbl.cget("text")
    assert len(app.badge_chips) == len(ACHIEVEMENT_DEFS) == 12
    assert app.badge_chips[0][1].cget("text").startswith("✓")
    app.show_achievements()
    app.update_idletasks()
    app.update()
    _ach_win = [w for w in app.winfo_children()
                if isinstance(w, DiagnosisWindow)
                and "도전 과제" in w.title()]
    _txt = _ach_win[-1].text_value if _ach_win else ""
    assert _ach_win and "[도전 과제 " in _txt and "○ 미달성" in _txt \
        and all(a["name"] in _txt for a in ACHIEVEMENT_DEFS), _txt[:200]
    assert "점수대" in app.band_lbl.cget("text"), app.band_lbl.cget("text")
    app.show_score_board()
    app.update_idletasks()
    app.update()
    _bw = app.board_win
    assert _bw is not None and _bw.winfo_exists()
    assert str(_bw.start_btn["state"]) == "disabled"
    if _bw.rows:
        _bw._select(_bw.rows[0][0])
        app.update()
        assert str(_bw.start_btn["state"]) == "normal"
        assert "선택: " in _bw.pick_lbl.cget("text")
    assert band_summary_text([{"name": "미달", "count": 2},
                              {"name": "합격", "count": 0}]) == "미달 2"
    _bw.destroy()
    app.update()
    assert app.board_win is None
    app.exam_running = True
    assert app.routine_action("start_exam", None, None, None, None) == \
        (False, "시험 진행 중")
    app.exam_running = False
    assert app.routine_action("show", None, None, None, None) == (True, None)
    assert app.routine_action("open_pdf", {"name": "x", "pdf": None}, None,
                              "x", None)[0] is False
    _pos0, _pick0, _ov0 = (saved_overlay_pos(), monitor_assignment(),
                           timer_overlay_enabled())
    _sd = SettingsDialog(app)
    app.update_idletasks()
    app.update()
    _mons2 = _sd.monitors
    assert bool(_sd.mon_vars) is monitor_ui_visible(_mons2), \
        f"모니터 {len(_mons2)}대인데 배치 칸 {'있음' if _sd.mon_vars else '없음'}"
    if _sd.mon_vars:
        assert set(_sd.mon_vars) == {"pdf", "excel", "timer"}
        assert _sd.apply_monitors() == {k: _sd._monitor_choice(k)
                                        for k in ("pdf", "excel", "timer")}
        assert _sd.reset_monitors() == resolve_monitor_assignment(_mons2)
    _sd.overlay_var.set(not _ov0)
    assert _sd._toggle_overlay() is (not _ov0) \
        and timer_overlay_enabled() is (not _ov0)
    _sd.overlay_var.set(_ov0)
    assert _sd._toggle_overlay() is _ov0
    assert _sd.reset_overlay_pos() is None and saved_overlay_pos() is None
    _sd.destroy()
    save_overlay_pos(*(_pos0 or (None, None)))
    save_monitor_assignment(**_pick0)
    timer.finished = True
    timer.destroy()
    assert timer.overlay is None and TimerOverlay.close_all() == 0
    app.destroy()
    assert not app.routine.running if app.routine else True
    print("SMOKE OK: 창 생성/위젯 렌더/타이머(+우측 상단 오버레이)/설정 창"
          "(모니터 배치)/오답노트 패널/단계 가이드(세트 "
          "자동 선택·바꾸기·미발견 직접 선택)/오류 대화상자·로그/진단 창/"
          "시작 로그 창/Excel 확인 안내 창/PDF 회차 경고/안내 띠·자동 업데이트 "
          "토글/루틴 연동 서버(상태·쓰기·페이지·퀴즈 버튼·결과 문구)/"
          "게임형 상태 카드(티어·게이지·연속·도전 과제 창)/점수대 보드 창"
          "(밴드 묶음·세트 선택·이 세트 응시)/결과 연출"
          "(승급·새 배지·합격선 게이지)/파괴 정상")

def _notify_no_tk():
    msg = ("tkinter를 사용할 수 없습니다.\n"
           "Windows용 Python 설치 시 'tcl/tk and IDLE' 옵션을 포함해 주세요.\n"
           "(리눅스: sudo apt install python3-tk)")
    shown = False
    try:
        if sys.stdout is not None:
            print(msg)
            shown = True
    except Exception:
        pass
    if not shown:
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(None, msg, APP_TITLE, 0x10)
        except Exception:
            pass

def _show_fatal(title, msg):
    shown = False
    if HAS_TK:
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(title, msg)
            root.destroy()
            shown = True
        except Exception:
            shown = False
    if not shown:
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(None, str(msg), str(title), 0x10)
            shown = True
        except Exception:
            pass
    if not shown:
        try:
            sys.__stderr__.write(f"{title}: {msg}\n")
        except Exception:
            pass

def _platform_text():
    try:
        return platform.platform()
    except Exception:
        return sys.platform

def main():
    ap = argparse.ArgumentParser(description=APP_TITLE)
    ap.add_argument("--scan-root", help="모의고사 스캔 폴더 (기본: 상위 폴더)")
    ap.add_argument("--smoke", action="store_true",
                    help="GUI 스모크 테스트 후 종료 (개발용)")
    args = ap.parse_args()
    startup_log(f"프로그램 시작 v{__version__} · python={sys.executable} · "
                f"{_platform_text()} · 실행 파일={sys.argv[0] if sys.argv else ''}"
                f" · 옵션={sys.argv[1:]}")
    if not HAS_TK:
        startup_log("tkinter 없음 — 종료")
        _notify_no_tk()
        return 1
    if args.smoke:
        run_smoke()
        return 0
    try:
        app = ExamApp(scan_root=args.scan_root)
    except Exception as e:
        text = log_error("시작 화면 생성", e)
        _show_fatal(f"{APP_TITLE} - 시작 오류",
                    "시작 화면을 만들지 못했습니다.\n오류 내용은 "
                    f"{ERROR_LOG_PATH}에 기록되었습니다. 이 내용을 복사해 "
                    "채팅에 붙여넣어 주세요.\n\n" + text[-1500:])
        return 1
    try:
        app.mainloop()
    except Exception as e:
        text = log_error("메인 루프", e)
        _show_fatal(f"{APP_TITLE} - 오류",
                    "프로그램이 예기치 않게 중단되었습니다.\n오류 내용은 "
                    f"{ERROR_LOG_PATH}에 기록되었습니다.\n\n" + text[-1500:])
        return 1
    startup_log("정상 종료")
    return 0

if __name__ == "__main__":
    sys.exit(main())
