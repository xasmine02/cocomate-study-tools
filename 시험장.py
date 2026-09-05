#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
코코 시험장 — 컴활 2급 실기 모의고사 런처

시작 화면에서 모의고사 세트를 고르고 [시험 시작]을 누르면
문제 파일 사본이 Excel로 열리고 타이머가 시작됩니다.
제출하면 grade.py로 자동 채점해 점수와 리포트를 보여 줍니다.
오늘의 학습(일정 v4, 9/3 시작·시험 2회): 매일 모의고사 1세트 40분 완주
→ 채점 → 오답노트 → 오답 재풀이. 세트는 기록 기반으로 자동 선택됩니다.

의존성: Python 표준 라이브러리 + tkinter (채점은 grade.py/openpyxl 필요)
"""

__version__ = "2.1.1"

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import webbrowser
from datetime import date, datetime, timedelta

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    HAS_TK = True
except ImportError:
    HAS_TK = False

# ---------------------------------------------------------------------------
# 상수 / 브랜딩
# ---------------------------------------------------------------------------

APP_TITLE = "코코 시험장"
BRAND = "#107C41"        # 코코 그린
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


def _ensure_records_home():
    """기록.json 보관 폴더 생성 + 구버전 위치의 기록 자동 이전."""
    try:
        os.makedirs(RECORDS_DIR, exist_ok=True)
        if os.path.isfile(_OLD_RECORDS_PATH) \
                and not os.path.isfile(RECORDS_PATH):
            shutil.move(_OLD_RECORDS_PATH, RECORDS_PATH)
    except OSError:
        pass

def log_error(context, exc=None, path=None):
    """오류를 채점결과/시험장_오류.log에 append (pythonw에서도 흔적 보존).

    exc: 예외 객체 또는 (type, value, tb). 반환: 기록한 텍스트.
    """
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
    return text


UI_FONT = ("Malgun Gothic", 10)
UI_FONT_BOLD = ("Malgun Gothic", 10, "bold")
DIGIT_FONT = ("Consolas", 44, "bold")

PROBLEM_RE = re.compile(r"^(?P<name>.+)_문제\.(?P<ext>xlsx|xlsm)$", re.IGNORECASE)


# ---------------------------------------------------------------------------
# 헤드리스 로직 (GUI 없이 테스트 가능)
# ---------------------------------------------------------------------------


def default_scan_root():
    """기본 스캔 경로: 시험장.py가 있는 폴더의 상위 폴더."""
    parent = os.path.dirname(BASE_DIR)
    return parent if parent else BASE_DIR


def find_grade_py(user_path=None):
    """grade.py 탐색: 같은 폴더 -> ../채점/ -> 사용자 지정."""
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
_ROLE_TOKENS_STRIP = ("문제지", "기대값", "정답지", "답안지", "정답", "답안",
                      "문제")


def norm_set_key(name):
    """파일명 -> 세트 식별 키: 소문자화, 공백/언더스코어/괄호/구분자 제거,
    역할 토큰(문제/정답/답안/문제지/기대값) 제거."""
    s = os.path.splitext(os.path.basename(str(name)))[0].lower()
    s = re.sub(r"[\s_\-()\[\]{}.,·~]+", "", s)
    for tok in _ROLE_TOKENS_STRIP:
        s = s.replace(tok, "")
    return s


def file_role(name):
    """파일명에서 역할 판별: 'answer' / 'problem' / None."""
    base = os.path.splitext(os.path.basename(str(name)))[0]
    if "정답" in base or "답안" in base:
        return "answer"
    if "문제" in base:
        return "problem"
    return None


def display_name(problem_path):
    """문제 파일명에서 세트 표시명 (뒤쪽 '문제'/'(문제)' 토큰 제거)."""
    stem = os.path.splitext(os.path.basename(str(problem_path)))[0]
    stem = re.sub(r"[\s_\-]*[(\[]?\s*문제(지)?\s*[)\]]?[\s_\-]*$", "",
                  stem).strip(" _-")
    return stem or os.path.splitext(os.path.basename(str(problem_path)))[0]


def set_tokens(text):
    """세트 식별 토큰: 연도(20xx), N회, A/B/가/나형, N급, 키워드."""
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
    for m in re.finditer(r"(?<!\d)(2[0-9])(?!\d)", s):   # 2자리 연도 '24'
        toks.add(m.group(1))
    for word in ("상시", "코코", "모의", "복원", "기출", "실기", "필기",
                 "드릴", "계산", "컴활"):
        if word in s:
            toks.add(word)
    return toks


def _token_conflict(a, b):
    """연도/회차/형 토큰이 양쪽 모두에 있는데 서로 다르면 충돌."""
    for pat in (r"20\d{2}", r"\d{1,2}회", r"[ab가나]형"):
        ca = {t for t in a if re.fullmatch(pat, t)}
        cb = {t for t in b if re.fullmatch(pat, t)}
        if ca and cb and not (ca & cb):
            return True
    return False


def match_pdf_for_set(toks, pdf_paths):
    """토큰 우선순위 매칭으로 유일한 문제지 PDF 찾기. 복수/0개면 None."""
    scored = []
    for p in pdf_paths:
        pt = set_tokens(os.path.basename(p))
        if _token_conflict(toks, pt):
            continue
        shared = len(toks & pt)
        if shared >= 1:
            scored.append((shared, p))
    if not scored:
        return None
    best = max(s for s, _p in scored)
    matched = [p for s, p in scored if s == best]
    return matched[0] if len(matched) == 1 else None


def load_set_config(path=SET_CONFIG_PATH):
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_set_config(config, path=SET_CONFIG_PATH):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        return True
    except OSError:
        return False


def remember_set(s, path=SET_CONFIG_PATH):
    """세트 구성을 세트설정.json에 저장 (정규화 키 기준, 기존 항목과 병합)."""
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
    return save_set_config(cfg, path)


def apply_set_config(sets, config):
    """저장된 세트 구성 반영. 사라진 경로는 무시."""
    by_key = {s["norm"]: s for s in sets}
    for k, ent in (config or {}).items():
        if not isinstance(ent, dict):
            continue
        pdf, keyj = ent.get("pdf"), ent.get("key")
        if k in by_key:
            s = by_key[k]
            if pdf and os.path.isfile(pdf):
                s["pdf"] = pdf
            if keyj and os.path.isfile(keyj):
                s["key"] = keyj
            continue
        prob, ans = ent.get("problem"), ent.get("answer")
        if prob and ans and os.path.isfile(prob) and os.path.isfile(ans):
            sets.append({
                "name": ent.get("name") or display_name(prob),
                "norm": k, "dir": os.path.dirname(prob),
                "problem": prob, "answer": ans,
                "key": keyj if keyj and os.path.isfile(keyj) else None,
                "pdf": pdf if pdf and os.path.isfile(pdf) else None,
                "saved": True,
            })
    return sets


def scan_sets(root, config=None):
    """느슨한 세트 그룹핑 스캔.

    xlsx/xlsm/pdf/기대값 json을 정규화 키로 묶고, 파일명 토큰으로
    문제/정답 역할을 판별합니다. 같은 키에 PDF가 없으면 토큰 퍼지
    매칭으로 문제지 PDF를 연결합니다. '풀이_' 파일과 '채점결과' 폴더
    제외. 반환: [{"name","norm","dir","problem","answer","key","pdf"}]
    """
    sets = []
    all_pdfs = []
    groups = {}
    if root and os.path.isdir(root):
        for dirpath, dirnames, filenames in os.walk(root):
            dirnames[:] = [d for d in dirnames
                           if not d.startswith((".", "__"))
                           and d != "채점결과"]
            for fn in sorted(filenames):
                ext = os.path.splitext(fn)[1].lower()
                if ext not in (".xlsx", ".xlsm", ".pdf", ".json"):
                    continue
                if fn.startswith(("풀이_", "채점결과", "~$", ".")):
                    continue
                path = os.path.join(dirpath, fn)
                if ext == ".json" and "기대값" not in fn:
                    continue
                if ext == ".pdf":
                    all_pdfs.append(path)
                k = norm_set_key(fn)
                if not k:
                    continue
                g = groups.setdefault(k, {"excel": [], "pdf": [], "json": []})
                if ext in (".xlsx", ".xlsm"):
                    g["excel"].append(path)
                elif ext == ".pdf":
                    g["pdf"].append(path)
                else:
                    g["json"].append(path)
    for k, g in groups.items():
        answers = [p for p in g["excel"] if file_role(p) == "answer"]
        problems = [p for p in g["excel"] if file_role(p) == "problem"]
        answer = sorted(answers)[0] if answers else None
        problem = sorted(problems)[0] if problems else None
        if not problem and answer and len(g["excel"]) == 2:
            # 정답 토큰만 있는 2파일 세트 -> 나머지를 문제로 추정
            problem = next(p for p in g["excel"] if p != answer)
        if not (problem and answer) or problem == answer:
            continue
        sets.append({
            "name": display_name(problem), "norm": k,
            "dir": os.path.dirname(problem),
            "problem": problem, "answer": answer,
            "key": sorted(g["json"])[0] if g["json"] else None,
            "pdf": sorted(g["pdf"])[0] if g["pdf"] else None,
        })
    # 같은 키 PDF가 없는 세트: 토큰 퍼지 매칭 (유일할 때만)
    used = {s["pdf"] for s in sets if s["pdf"]}
    for s in sets:
        if s["pdf"]:
            continue
        toks = set_tokens(s["name"]) | set_tokens(os.path.basename(s["dir"]))
        cand = match_pdf_for_set(toks, [p for p in all_pdfs if p not in used])
        if cand:
            s["pdf"] = cand
            used.add(cand)
    apply_set_config(sets, config if config is not None else load_set_config())
    sets.sort(key=lambda s: s["name"])
    return sets


def build_direct_set(problem, answer, pdf=None):
    """직접 선택한 문제/정답(+문제지)으로 세트 구성 (기대값/문제지 자동 감지)."""
    d = os.path.dirname(os.path.abspath(problem))
    k = norm_set_key(problem)
    key = None
    auto_pdf = None
    try:
        for fn in os.listdir(d):
            if fn.startswith("풀이_"):
                continue
            ext = os.path.splitext(fn)[1].lower()
            if ext == ".json" and "기대값" in fn and norm_set_key(fn) == k:
                key = os.path.join(d, fn)
            elif ext == ".pdf" and norm_set_key(fn) == k:
                auto_pdf = os.path.join(d, fn)
    except OSError:
        pass
    return {
        "name": display_name(problem), "norm": k, "dir": d,
        "problem": os.path.abspath(problem),
        "answer": os.path.abspath(answer),
        "key": key,
        "pdf": os.path.abspath(pdf) if pdf else auto_pdf,
    }


def load_records(path=RECORDS_PATH):
    if path == RECORDS_PATH:
        _ensure_records_home()
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except Exception:
        return []


def append_record(record, path=RECORDS_PATH):
    """기록.json에 응시 기록 1건 추가. 전체 목록 반환."""
    if path == RECORDS_PATH:
        _ensure_records_home()
    records = load_records(path)
    records.append(record)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    return records


def records_summary(records):
    """세트별 최고점과 최근 3회 점수 (부분 연습·오답 재풀이 기록은 제외).

    {세트명: {"best":.., "recent":[..]}}"""
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


# --- 매크로 보존 사본 생성 (xlsx -> xlsm ZIP 변환 + MotW 제거) ---

XLSX_MAIN_CT = ("application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet.main+xml")
XLSM_MAIN_CT = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"


def convert_xlsx_to_xlsm(src, dst):
    """xlsx -> xlsm ZIP 수준 변환.

    [Content_Types].xml의 워크북 메인 파트 content-type만 교체하고 그 외
    모든 파트는 바이트 무손실 복사 (openpyxl 재저장 없음 — 서식·차트 보존).
    """
    import zipfile
    with zipfile.ZipFile(src) as zin, \
            zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(XLSX_MAIN_CT.encode("utf-8"),
                                    XLSM_MAIN_CT.encode("utf-8"))
            zout.writestr(item, data)
    return dst


def _strip_motw(path):
    """Mark-of-the-Web(Zone.Identifier ADS) 제거 시도 — 실패는 조용히 무시."""
    try:
        os.remove(path + ":Zone.Identifier")
    except OSError:
        pass


def _unique_stem(d, stem):
    """같은 이름의 .xlsm/.xlsx 사본과 겹치지 않는 경로 어간."""
    cand = os.path.join(d, stem)
    n = 2
    while os.path.exists(cand + ".xlsm") or os.path.exists(cand + ".xlsx"):
        cand = os.path.join(d, f"{stem}_{n}")
        n += 1
    return cand


def copy_as_macro_enabled(source, dst_stem):
    """사본을 매크로 저장 가능한 .xlsm으로 생성.

    원본이 .xlsx면 ZIP 수준 변환(무결성 스모크 실패 시 원본 확장자로
    폴백), .xlsm이면 그대로 복사. 사본의 MotW도 제거.
    """
    ext = os.path.splitext(source)[1].lower()
    if ext == ".xlsx":
        dst = dst_stem + ".xlsm"
        try:
            convert_xlsx_to_xlsm(source, dst)
            try:  # 무결성 스모크 (openpyxl 없으면 생략)
                import openpyxl
                openpyxl.load_workbook(dst).close()
            except ImportError:
                pass
        except Exception:
            try:
                if os.path.isfile(dst):
                    os.remove(dst)
            except OSError:
                pass
            dst = dst_stem + ".xlsx"
            shutil.copy2(source, dst)
    else:
        dst = dst_stem + (ext or ".xlsm")
        shutil.copy2(source, dst)
    _strip_motw(dst)
    return dst


def make_attempt_copy(problem, set_name, when=None):
    """문제 파일을 풀이_<세트명>_<일시>.xlsm 사본으로 (매크로 저장 가능)."""
    when = when or datetime.now()
    stamp = when.strftime("%Y%m%d_%H%M")
    d = os.path.dirname(os.path.abspath(problem))
    return copy_as_macro_enabled(
        problem, _unique_stem(d, f"풀이_{set_name}_{stamp}"))


# --- 앱 설정 (세트설정.json의 "_설정" 영역) ---

def get_app_setting(name, default=None, path=None):
    cfg = load_set_config(path or SET_CONFIG_PATH)
    return (cfg.get("_설정") or {}).get(name, default)


def set_app_setting(name, value, path=None):
    p = path or SET_CONFIG_PATH
    cfg = load_set_config(p)
    cfg.setdefault("_설정", {})[name] = value
    return save_set_config(cfg, p)


# --- Excel 신뢰 위치 등록 (매크로 차단 배너 해결) ---

TRUST_MANUAL_GUIDE = (
    "수동 설정: Excel → 파일 → 옵션 → 보안 센터 → 보안 센터 설정 → "
    "신뢰할 수 있는 위치 → '새 위치 추가'에서 학습 폴더를 추가하고 "
    "'이 위치의 하위 폴더도 신뢰할 수 있음'을 체크하세요.")


def build_trusted_location_values(folder, version="16.0", slot="LocationCC"):
    """등록할 레지스트리 키 경로·값 구성 (테스트 가능)."""
    path = (rf"Software\Microsoft\Office\{version}\Excel\Security"
            rf"\Trusted Locations\{slot}")
    folder = str(folder).rstrip("\\/") + "\\"
    return {"key": path,
            "values": {"Path": folder, "AllowSubfolders": 1,
                       "Description": "코코 시험장 학습 폴더"}}


def register_trusted_location(folder, version=None):
    """HKCU에 신뢰 위치 등록. (성공 여부, 메시지) 반환.

    Windows 외 환경/실패 시 수동 설정 경로 안내를 메시지에 포함.
    """
    try:
        import winreg
    except ImportError:
        return False, ("이 기능은 Windows에서만 동작합니다.\n\n"
                       + TRUST_MANUAL_GUIDE)
    versions = [version] if version else []
    try:  # 설치된 Office 버전 키 탐색
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
        versions.append("16.0")  # Office 2016+/365 기본
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
    """grade.py 실행 인자 목록. sheets 지정 시 '--sheets 시트1,시트2' 부분 채점.
    (오답 재풀이는 오답 시트 목록을 그대로 넘깁니다 — GUI 없이 테스트 가능)"""
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
    """grade.py를 서브프로세스로 실행. sheets 지정 시 부분 채점.

    반환: (결과 dict 또는 None, stdout, stderr, returncode)
    결과 dict는 --json 출력이 우선, 실패 시 콘솔 총점 파싱 폴백.
    """
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
    """파일을 기본 프로그램으로 열기. (성공 여부, 오류 메시지) 반환."""
    path = os.path.abspath(path)
    try:
        if sys.platform == "win32":
            os.startfile(path)  # noqa
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


MANUAL_PIP = ("수동 설치: 명령 프롬프트에서  py -m pip install openpyxl\n"
              "(위 명령이 안 되면: python -m pip install openpyxl)")


def _pip_python():
    """pip 실행용 인터프리터. pythonw.exe면 같은 폴더 python.exe로 치환."""
    exe = sys.executable
    base = os.path.basename(exe).lower()
    if base.startswith("pythonw"):
        cand = os.path.join(os.path.dirname(exe),
                            base.replace("pythonw", "python", 1))
        if os.path.isfile(cand):
            return cand
    return exe


def install_openpyxl(timeout=600):
    """pip로 openpyxl 설치 (창 숨김). (성공 여부, 출력 요약) 반환."""
    kwargs = {}
    if sys.platform == "win32":
        kwargs["creationflags"] = 0x08000000  # CREATE_NO_WINDOW
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
    """채점 실패 원인 분류: 'module' | 'file' | 'other'."""
    t = text or ""
    if rc == 3 or "ModuleNotFoundError" in t or "ImportError" in t \
            or "openpyxl 라이브러리가 설치되어" in t:
        return "module"
    if "FileNotFoundError" in t or "BadZipFile" in t \
            or "찾을 수 없습니다" in t or "열 수 없습니다" in t:
        return "file"
    return "other"


# ---------------------------------------------------------------------------
# 오늘의 학습 — 14일 루틴 일정 v4 (웹 루틴과 동일: 2026-09-03 시작, 시험 2회)
#   매일 모의고사 1세트 40분 완주 → 채점 → 오답노트 → 오답 재풀이
# ---------------------------------------------------------------------------

ROUTINE_START = date(2026, 9, 3)                     # Day 1 (9/3 목)
EXAM_DATES = [date(2026, 9, 11), date(2026, 9, 18)]  # 시험 1 · 시험 2
EXAM_DATE = EXAM_DATES[0]                            # (구 코드 호환)
ROUTINE_TAG = "2026-09"     # 기록.json 루틴 세대 표시 (구 루틴 기록과 구분)
PROGRESS_NS = "r0903"       # 세트설정.json '_진행' 키 접두 (구 루틴 진행과 분리)

PLAN_EXAM1, PLAN_EXAM2, PLAN_AFTER = 15, 16, 17     # 시험 1 / 시험 2 / 이후
PLAN_ORDER = ([0] + list(range(1, 9)) + [PLAN_EXAM1]
              + list(range(9, 15)) + [PLAN_EXAM2, PLAN_AFTER])  # 시간순
AUTO = "자동선택"           # 세트 슬롯 값: 기록 기반 자동 선택

ROUTINE_PLAN = {
    0: {"제목": "내일 시작", "종류": "안내",
        "할일": "내일 9/3(목)부터 Day 1이 시작됩니다. 프로그램 실행·채점 "
               "흐름 확인, 모의고사 세트·문제지 PDF 준비, 루틴 웹페이지 "
               "즐겨찾기까지 오늘 마쳐 두세요."},
    1: {"제목": "1차 완주", "종류": "모의", "세트": ["2024 상시 1회"],
        "목표": None},
    2: {"제목": "1차 완주", "종류": "모의", "세트": ["코코 1회"],
        "목표": None},
    3: {"제목": "1차 완주", "종류": "모의",
        "세트": ["2024 A형", "2024 상시 2회"], "목표": None},
    4: {"제목": "1차 완주", "종류": "모의", "세트": ["코코 2회", "2024 B형"],
        "목표": None},
    5: {"제목": "1차 완주", "종류": "모의", "세트": ["24 2급 상시"],
        "목표": None},
    6: {"제목": "1차 완주", "종류": "모의", "세트": ["컴활 2급 상시"],
        "목표": None},
    7: {"제목": "재도전", "종류": "모의", "세트": ["2026 1회"], "목표": 65},
    8: {"제목": "시험 1 전 마무리", "종류": "모의", "세트": [AUTO],
        "목표": 70, "특별": ["실수노트"]},
    9: {"제목": "2차 완주", "종류": "모의", "세트": [AUTO], "목표": 75,
        "특별": ["복기"]},
    10: {"제목": "2차 완주", "종류": "모의", "세트": [AUTO, AUTO],
         "목표": 75},
    11: {"제목": "2차 완주", "종류": "모의", "세트": [AUTO], "목표": 80,
         "메모": "신규 세트 우선"},
    12: {"제목": "2차 완주", "종류": "모의", "세트": [AUTO], "목표": 80},
    13: {"제목": "2차 완주", "종류": "모의", "세트": [AUTO], "목표": 85},
    14: {"제목": "시험 2 전 최종", "종류": "모의", "세트": [AUTO],
         "목표": 85, "특별": ["실수노트"]},
    PLAN_EXAM1: {"제목": "시험 1", "종류": "안내",
                 "할일": "시험 1 당일. 수험표·신분증 확인, 고사장 30분 전 "
                        "도착. 저장은 Ctrl+S 수시로, 계산작업은 한 문제 3분 "
                        "넘기면 다음으로. 준비한 만큼 충분합니다 — 화이팅!"},
    PLAN_EXAM2: {"제목": "시험 2", "종류": "안내",
                 "할일": "시험 2 당일. 시험 1의 복기 메모를 한 번 훑고 "
                        "출발하세요. 실수 노트의 항목만 지키면 됩니다 — "
                        "화이팅!"},
    PLAN_AFTER: {"제목": "루틴 완주", "종류": "안내",
                 "할일": "14일 루틴과 시험 2회를 완주했습니다. 수고 많았습니다! "
                        "결과와 관계없이 쌓은 실력은 남습니다."},
}

PREP_STEPS = [
    {"이름": "프로그램 실행·채점 확인", "형": "안내", "분": 15,
     "설명": "아무 세트나 [시험 시작]으로 열고 바로 제출해 채점까지 한 번 "
            "돌려보세요. openpyxl 설치 안내가 뜨면 설치합니다."},
    {"이름": "세트·문제지 PDF 준비 확인", "형": "안내", "분": 10,
     "설명": "아래 목록에 기출 세트와 코코 모의고사가 모두 보이고 "
            "(문제지 미연결) 표시가 없는지 확인하세요. 새 세트는 폴더에 "
            "넣기만 하면 자동으로 편입됩니다."},
    {"이름": "루틴 웹페이지 즐겨찾기", "형": "안내", "분": 5,
     "설명": "웹 루틴 페이지에서 실수 노트·함수 사전 위치를 확인하고 "
            "즐겨찾기에 추가하세요. 내일 9/3(목) Day 1: 2024 상시 1회 "
            "40분 완주로 시작합니다."},
]


def routine_date_for(no):
    """일정 번호 -> 날짜. Day 1~14는 시험일을 건너뛰어 배정(시험일 제외)."""
    no = int(no)
    if 1 <= no <= 14:
        d = ROUTINE_START + timedelta(days=no - 1)
        for ex in sorted(EXAM_DATES):   # 시험일에 닿으면 하루 밀림
            if d >= ex:
                d += timedelta(days=1)
        return d
    if no == PLAN_EXAM1:
        return EXAM_DATES[0]
    if no == PLAN_EXAM2:
        return EXAM_DATES[1]
    return None


def routine_day_no(today=None):
    """날짜 -> 일정 번호. 0=시작 전, 1~14=Day, 15=시험1, 16=시험2, 17=이후."""
    today = today or date.today()
    if today < ROUTINE_START:
        return 0
    for no in PLAN_ORDER:
        if routine_date_for(no) == today:
            return no
    return PLAN_AFTER


def plan_day_tag(no):
    """일정 번호 -> 진행 저장/기록 day 라벨 ('d00'~'d14', 그 외 None)."""
    return f"d{int(no):02d}" if 0 <= int(no) <= 14 else None


def dday_text(today=None):
    """시험 2회 D-day 병기 문구: '시험1 D-6 · 시험2 D-13'."""
    today = today or date.today()
    parts = []
    for i, ex in enumerate(EXAM_DATES, 1):
        n = (ex - today).days
        txt = "D-day" if n == 0 else (f"D-{n}" if n > 0 else f"D+{-n}")
        parts.append(f"시험{i} {txt}")
    return " · ".join(parts)


def plan_slot_names(plan):
    """세트 슬롯 표시명 목록 (자동 슬롯은 '자동 선택')."""
    return [spec if spec != AUTO else "자동 선택"
            for spec in (plan.get("세트") or [])]


def build_day_steps(plan, slot_names=None):
    """일정 -> 스텝 시퀀스 (공통 템플릿, 세트 슬롯마다 ①~④ 반복).

    ① 시험 모드 40분 완주 ② 채점·성적 복사(자동 체크) ③ 오답노트 모드
    ④ 오답 재풀이 15분 → (특별) 실수 노트 → ⑤ (선택) 함수 퀴즈.
    """
    no = plan.get("no", 0)
    if plan.get("종류") != "모의":
        return [dict(s) for s in PREP_STEPS] if no == 0 else []
    slots = list(plan.get("세트") or [])
    names = list(slot_names or plan_slot_names(plan))
    goal = plan.get("목표")
    g_txt = f" (목표 {goal}점)" if goal else ""
    steps = []
    if "복기" in (plan.get("특별") or []):
        steps.append({"이름": "시험 1 복기", "형": "안내", "분": 15,
                      "설명": "어제 시험 1을 복기하세요 — 막힌 문제·시간 "
                             "배분·실수를 웹 루틴 실수 노트에 3줄로 "
                             "적습니다. 시험 2까지 이 노트만 지키면 됩니다."})
    for k, spec in enumerate(slots):
        name = names[k] if k < len(names) else "자동 선택"
        n_txt = f" ({k + 1}/{len(slots)})" if len(slots) > 1 else ""
        auto_txt = (" 자동 선택된 세트는 창 위쪽에 이유와 함께 표시되며 "
                    "[다른 세트로 바꾸기]로 바꿀 수 있습니다."
                    if spec == AUTO else "")
        steps.extend([
            {"이름": f"{name} 시험 모드 40분 완주{n_txt}", "형": "모의",
             "세트": spec, "슬롯": k, "목표": goal, "분": 40,
             "설명": f"{name} 40분 실전 완주{g_txt}. 제출하면 자동 채점되고 "
                    "이 단계와 다음 '채점·성적 복사' 단계가 자동으로 "
                    "체크됩니다." + auto_txt},
            {"이름": f"채점·성적 복사{n_txt}", "형": "채점", "세트": spec,
             "슬롯": k, "분": 5,
             "설명": "채점이 끝나면 자동 체크됩니다. 성적 JSON은 클립보드에 "
                    "복사되어 있으니 웹 루틴에 [성적 붙여넣기]하세요."},
            {"이름": f"오답노트 모드{n_txt}", "형": "오답노트", "세트": spec,
             "슬롯": k, "분": 15,
             "설명": "틀린 항목의 해설을 하나씩 읽고 '이해했음'을 체크하세요. "
                    "이전 풀이 사본이 함께 열립니다."},
            {"이름": f"오답 재풀이 15분{n_txt}", "형": "오답재풀이",
             "세트": spec, "슬롯": k, "분": 15,
             "설명": "오답이 있던 시트만 새 사본(오답재풀이_*.xlsm)에서 15분 "
                    "안에 다시 풀고 제출하세요. 그 시트들만 채점됩니다."},
        ])
    if "실수노트" in (plan.get("특별") or []):
        steps.append({"이름": "실수 노트 정리", "형": "안내", "분": 10,
                      "설명": "오늘까지의 오답에서 반복된 실수를 웹 루틴 실수 "
                             "노트에 정리하세요. 시험장에서 볼 마지막 "
                             "체크리스트입니다."})
    steps.append({"이름": "(선택) 함수 퀴즈", "형": "안내", "분": 10,
                  "선택": True,
                  "설명": "(선택) 웹 루틴 함수 퀴즈 10문제 — 오늘 틀린 함수가 "
                         "있으면 그 함수부터. 건너뛰어도 됩니다."})
    return steps


def _default_todo(plan):
    names = plan_slot_names(plan)
    txt = ("모의고사 " + " + ".join(names) + " 40분 완주 → 채점 → 오답노트 "
           "→ 오답 재풀이 15분")
    if len(names) > 1:
        txt += f" (오늘은 {len(names)}세트)"
    if plan.get("메모"):
        txt += f"  [{plan['메모']}]"
    return txt


def plan_for_day(no, slot_names=None):
    """일정 번호 -> 일정 dict (no/날짜/할일/스텝 포함)."""
    no = int(no)
    if no not in ROUTINE_PLAN:
        no = max(0, min(PLAN_AFTER, no))
        if no not in ROUTINE_PLAN:
            no = 0
    plan = dict(ROUTINE_PLAN[no])
    plan["no"] = no
    plan["날짜"] = routine_date_for(no)
    plan["세트"] = list(plan.get("세트") or [])
    if not plan.get("할일"):
        plan["할일"] = _default_todo(plan)
    plan["스텝"] = build_day_steps(plan, slot_names)
    return plan


# 날짜별 기본 스텝 시퀀스 (세트명은 슬롯 표시명) — 무결성 테스트/조회용
ROUTINE_STEPS = {no: plan_for_day(no)["스텝"] for no in PLAN_ORDER}


def plan_title(plan, today=None, set_names=None):
    """카드 제목: 'Day 3 · 9/5(토) · 1차 완주 — 2024 A형 + 2024 상시 2회'.

    today를 주면 시험 2회 D-day를 병기합니다.
    """
    d = plan.get("날짜")
    d_txt = f"{d.month}/{d.day}({'월화수목금토일'[d.weekday()]})" if d else ""
    no = plan["no"]
    if 1 <= no <= 14:
        head = f"Day {no}"
        names = list(set_names or plan_slot_names(plan))
        body = f"{plan['제목']} — {' + '.join(names)}" if names \
            else plan["제목"]
    elif no in (PLAN_EXAM1, PLAN_EXAM2):
        head, body = "시험일", plan["제목"]
    elif no == PLAN_AFTER:
        head, body = "루틴", plan["제목"]
    else:
        s = ROUTINE_START
        head = "준비"
        body = (f"{plan['제목']} — {s.month}/{s.day}"
                f"({'월화수목금토일'[s.weekday()]}) Day 1")
    parts = [head]
    if d_txt:
        parts.append(d_txt)
    parts.append(body)
    if today:
        parts.append(dday_text(today))
    return " · ".join(parts)


# 일정 슬롯별 식별 토큰: 전 토큰 일치가 없을 때 이 토큰만으로 유일 매칭 허용
# (예: '컴활2급 A형 문제.xlsx'처럼 파일명에 '2024'가 없어도 A형 인식)
SLOT_IDENTITY = {
    "2024 상시 1회": ["상시", "1회"],
    "2024 상시 2회": ["상시", "2회"],
    "2024 A형": ["a형"],
    "2024 B형": ["b형"],
    "코코 1회": ["코코", "1회"],
    "코코 2회": ["코코", "2회"],
    "24 2급 상시": ["24", "2급", "상시"],
    "컴활 2급 상시": ["컴활", "상시"],
    "2026 1회": ["2026", "1회"],
}


def slot_identity_tokens(text):
    """슬롯 문구의 식별 토큰 (표에 없으면 연도(20xx)를 뺀 나머지 토큰)."""
    if text in SLOT_IDENTITY:
        return set(SLOT_IDENTITY[text])
    toks = set_tokens(text)
    rest = {t for t in toks if not re.fullmatch(r"20\d{2}", t)}
    return rest or toks


def _set_tokens_of(s):
    return set_tokens(s["name"]) | set_tokens(os.path.basename(s["dir"]))


def match_slot(sets, text):
    """슬롯 문구 -> (세트 or None, 판정 설명). GUI 없이 테스트 가능.

    ① 문구의 전 토큰이 세트에 있고 유일 → '전체 토큰 일치'
    ② 없으면 식별 토큰(SLOT_IDENTITY)만으로 유일 → '식별 토큰 일치'
    ③ 그래도 0개/복수면 미발견 (설명에 후보 나열).
    """
    toks = set_tokens(text)
    scored = []
    for s in sets:
        st = _set_tokens_of(s)
        if _token_conflict(toks, st) or (toks - st):
            continue
        shared = len(toks & st)
        if shared >= 1:
            scored.append((shared, s))
    if scored:
        best = max(sc for sc, _s in scored)
        matched = [s for sc, s in scored if sc == best]
        if len(matched) == 1:
            return matched[0], "전체 토큰 일치 (" + ", ".join(sorted(toks)) + ")"
        return None, "복수 후보 (전체 토큰): " + ", ".join(
            s["name"] for s in matched)
    idt = slot_identity_tokens(text)
    cands = [s for s in sets
             if idt <= _set_tokens_of(s)
             and not _token_conflict(idt, _set_tokens_of(s))]
    if len(cands) == 1:
        return cands[0], "식별 토큰 일치 (" + ", ".join(sorted(idt)) + ")"
    if not cands:
        return None, ("후보 없음 — 파일명에 " + "·".join(sorted(idt))
                      + " 토큰을 가진 문제/정답 짝이 없음")
    return None, "복수 후보 (식별 토큰 " + "·".join(sorted(idt)) + "): " + \
        ", ".join(s["name"] for s in cands)


def find_set_for_tokens(sets, text):
    """일정의 세트 지정 문구를 스캔된 세트에 매칭 (match_slot의 세트만)."""
    return match_slot(sets, text)[0]


def load_slot_mapping(path=None):
    """세트설정.json '_슬롯매핑': {슬롯 문구: {problem, answer, pdf}}."""
    cfg = load_set_config(path or SET_CONFIG_PATH)
    raw = cfg.get("_슬롯매핑") or {}
    return {k: v for k, v in raw.items() if isinstance(v, dict)}


def save_slot_mapping(spec, s, path=None):
    """슬롯 문구 -> 직접 선택한 세트 파일 매핑 저장 (이후 자동 해석)."""
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
    """매핑 항목 -> 세트 dict (파일이 없으면 None)."""
    try:
        prob, ans = ent.get("problem"), ent.get("answer")
        if not (prob and ans and os.path.isfile(prob) and os.path.isfile(ans)):
            return None
        pdf = ent.get("pdf")
        s = build_direct_set(prob, ans, pdf=pdf if pdf and
                             os.path.isfile(pdf) else None)
        if ent.get("name"):
            s["name"] = ent["name"]
        s["direct"] = True
        return s
    except Exception:
        return None


def scan_diagnosis_text(root, config=None, specs=None, sets=None):
    """[세트 인식 진단] 텍스트: 파일 → 키 → 역할 → 세트 → 슬롯 매칭 결과."""
    lines = [f"[코코 시험장 세트 인식 진단] {__version__} · "
             f"{datetime.now().strftime('%Y-%m-%d %H:%M')}",
             f"스캔 루트: {root}", ""]
    if config is None:
        config = load_set_config()
    if sets is None:
        sets = scan_sets(root, config) if root and os.path.isdir(root) else []
    owner = {}
    for s in sets:
        for k in ("problem", "answer", "pdf", "key"):
            if s.get(k):
                owner[os.path.abspath(s[k])] = s["name"]
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
    lines.append("== 파일 → 정규화 키 → 역할 → 소속 세트 ==")
    lines.append("파일명 | 정규화 키 | 역할 | 소속 세트 | 폴더")
    for p in files:
        fn = os.path.basename(p)
        ext = os.path.splitext(fn)[1].lower()
        if ext == ".pdf":
            role = "문제지(PDF)"
        else:
            role = {"problem": "문제", "answer": "정답"}.get(file_role(fn),
                                                        "미상(문제/정답 표기 없음)")
        own = owner.get(os.path.abspath(p), "미소속")
        rel = os.path.relpath(os.path.dirname(p), root) if root else ""
        lines.append(f"{fn} | {norm_set_key(fn)} | {role} | {own} | {rel}")
    if not files:
        lines.append("(xlsx/xlsm/pdf 파일 없음)")
    lines.append("")
    lines.append("== 인식된 세트 ==")
    lines.append("세트명 | 토큰 | 문제 | 정답 | PDF | 기대값")
    for s in sets:
        lines.append(f"{s['name']} | {','.join(sorted(_set_tokens_of(s)))} | "
                     f"{os.path.basename(s['problem'])} | "
                     f"{os.path.basename(s['answer'])} | "
                     f"{os.path.basename(s['pdf']) if s.get('pdf') else '-'} | "
                     f"{os.path.basename(s['key']) if s.get('key') else '-'}")
    if not sets:
        lines.append("(인식된 세트 없음 — 같은 폴더에 '…문제.xlsx'와 '…정답.xlsm' "
                     "짝이 있어야 합니다)")
    lines.append("")
    lines.append("== 일정 슬롯 매칭 ==")
    if specs is None:
        specs = []
        for no in range(1, 15):
            for spec in ROUTINE_PLAN.get(no, {}).get("세트") or []:
                if spec != AUTO and spec not in specs:
                    specs.append(spec)
    mapping = {k: v for k, v in (config.get("_슬롯매핑") or {}).items()
               if isinstance(v, dict)}
    for spec in specs:
        s, how = match_slot(sets, spec)
        line = f"슬롯 '{spec}' → " + (f"세트 '{s['name']}' ({how})" if s
                                     else f"미발견 ({how})")
        m = mapping.get(spec)
        if m:
            ms = set_from_mapping(m)
            line += (f" · 직접 선택 저장됨: {os.path.basename(m.get('problem') or '?')}"
                     + ("" if ms else " (파일 없음 — 무시)"))
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


# ---------------------------------------------------------------------------
# 자동 세트 선택 (재응시 세트 고르기) — GUI 없이 테스트 가능
# ---------------------------------------------------------------------------

def _record_date(r):
    try:
        return datetime.strptime(str(r.get("일시"))[:10], "%Y-%m-%d").date()
    except Exception:
        return None


def set_exam_records(set_name, records):
    """세트의 시험 모드(전체 응시) 기록 — 부분연습·오답재풀이 제외."""
    return [r for r in set_records(set_name, records)
            if r.get("mode") not in ("부분연습", "오답재풀이")]


def pick_set_for_retry(sets, records, count=1, today=None, exclude=()):
    """재응시 세트 자동 선택. [(세트, 이유)] count개 (서로 다른 세트).

    우선순위: ① 응시 기록이 없는 신규 세트(폴더에 새로 넣은 세트 편입)
    ② 기록상 최고점이 가장 낮은 세트 — 단, 최근 2일(오늘·어제) 응시한
    세트는 제외 ③ 전부 제외되면 최저점 세트. exclude: 제외할 세트/이름.
    """
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
    """세트설정.json '_자동선택'에 저장된 오늘 자동 선택 [{세트, 이유}]."""
    cfg = load_set_config(path or SET_CONFIG_PATH)
    raw = (cfg.get("_자동선택") or {}).get(day_tag) or []
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
    """오늘의 자동 선택 결과 저장 (picks: [{세트, 이유} 또는 None])."""
    p = path or SET_CONFIG_PATH
    cfg = load_set_config(p)
    cfg.setdefault("_자동선택", {})[day_tag] = [
        {"세트": x["세트"], "이유": x.get("이유", "")} if x else None
        for x in picks]
    return save_set_config(cfg, p)


def resolve_day_sets(plan, sets, records=None, today=None, saved=None,
                     mapping=None):
    """일정의 세트 슬롯 -> [(세트 or None, 이유)].

    고정 슬롯은 저장된 직접 선택(mapping) → 토큰 매칭(match_slot) 순,
    자동 슬롯은 저장된 선택(saved)을 우선 복원하고 없으면
    pick_set_for_retry로 고릅니다. 슬롯끼리는 서로 다른 세트.
    """
    slots = list(plan.get("세트") or [])
    out = [None] * len(slots)
    reasons = [""] * len(slots)
    taken = set()
    mapping = mapping or {}
    for k, spec in enumerate(slots):          # ① 고정 세트
        if spec != AUTO:
            s = set_from_mapping(mapping[spec]) if spec in mapping else None
            if s:
                how = "직접 선택 저장됨"
            else:
                s, how = match_slot(sets, spec)
            out[k], reasons[k] = s, (f"일정 지정 세트 · {how}" if s else
                                    f"세트를 찾지 못함 · {how}")
            if s:
                taken.add(s["norm"])
    saved = list(saved or [])
    for k, spec in enumerate(slots):          # ② 저장된 자동 선택 복원
        if spec == AUTO and k < len(saved) and saved[k]:
            s = next((x for x in sets if x["name"] == saved[k]["세트"]
                      and x["norm"] not in taken), None)
            if s:
                out[k] = s
                reasons[k] = (saved[k].get("이유") or "이전 선택") + " (유지)"
                taken.add(s["norm"])
    for k, spec in enumerate(slots):          # ③ 새로 자동 선택
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


# ---------------------------------------------------------------------------
# 오답 재풀이 (최신 채점 JSON의 오답 시트만 부분 채점) — GUI 없이 테스트 가능
# ---------------------------------------------------------------------------

def wrong_sheets_from_items(items):
    """오답 항목 -> 오답이 있는 시트 목록 (등장 순, 중복 제거)."""
    out = []
    for it in items or []:
        sh = str((it or {}).get("sheet") or "").strip()
        if sh and sh not in out:
            out.append(sh)
    return out


def make_retry_copy(problem, set_name, when=None):
    """문제 파일 -> 오답재풀이_<세트>_<일시>.xlsm 사본."""
    when = when or datetime.now()
    stamp = when.strftime("%Y%m%d_%H%M")
    d = os.path.dirname(os.path.abspath(problem))
    return copy_as_macro_enabled(
        problem, _unique_stem(d, f"오답재풀이_{set_name}_{stamp}"))


def retry_payload_for_set(s, minutes=15):
    """세트의 최신 전체 채점 JSON에서 오답 시트를 뽑아 재풀이 실행 정보로.

    반환 ('retry', payload) / ('missing', {이유}) / ('info', {메시지, 자동완료}).
    """
    jp = find_latest_result_json(s, full_only=True)
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


STEP_DONE_MESSAGE = ("오늘 완료! 웹 루틴에 성적 붙여넣기"
                     "(클립보드에 이미 복사됨)")


def sheet_names_of(path):
    """엑셀 파일의 시트 이름 목록 (openpyxl 없이 zip에서 직접)."""
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
    """드릴 워크북에서 n번 드릴 시트 이름 찾기.

    이름에 번호가 들어간 시트 우선(드릴1, 1.판정 등), 없으면 순서상
    n번째 시트. 못 찾으면 None.
    """
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


def _progress_key(day_tag):
    """'_진행' 저장 키 — 새 루틴 세대는 접두를 붙여 구 진행과 분리."""
    return f"{PROGRESS_NS}:{day_tag}"


def load_step_progress(day_tag, path=None):
    """세트설정.json '_진행'에서 완료 스텝 번호 집합 로드."""
    cfg = load_set_config(path or SET_CONFIG_PATH)
    raw = (cfg.get("_진행") or {}).get(_progress_key(day_tag)) or []
    out = set()
    for x in raw:
        try:
            out.add(int(x))
        except (TypeError, ValueError):
            pass
    return out


def save_step_progress(day_tag, done, path=None):
    """완료 스텝 번호 집합을 세트설정.json '_진행'에 저장."""
    p = path or SET_CONFIG_PATH
    cfg = load_set_config(p)
    cfg.setdefault("_진행", {})[_progress_key(day_tag)] = \
        sorted(int(i) for i in done)
    return save_set_config(cfg, p)


def resolve_step_action(step, sets, slot_sets=None):
    """스텝 -> 실행 방법. ('info'|'practice'|'exam'|'review'|'retry'|
    'missing', payload) 반환. GUI 없이 테스트 가능.

    slot_sets: resolve_day_sets() 결과 [(세트, 이유)] — 스텝의 '슬롯'
    번호로 세트를 찾습니다. 없으면 스텝의 '세트' 문구를 퍼지 매칭.
    """
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
    return "info", {"메시지": step.get("설명") or step.get("이름") or ""}


# ---------------------------------------------------------------------------
# 부분 연습 모드 (루틴 주차별 영역 연습)
# ---------------------------------------------------------------------------

PRACTICE_AREAS = [
    # (영역명, 포함 시트, 권장 시간(분))
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
    """영역명 목록 -> 시트명 목록 (순서 유지)."""
    out = []
    for name, sheet_list, _min in PRACTICE_AREAS:
        if name in area_names:
            out.extend(sheet_list)
    return out


def practice_minutes(area_names):
    """영역명 목록 -> 권장 시간 합(분)."""
    total = sum(m for name, _s, m in PRACTICE_AREAS if name in area_names)
    return max(MIN_MINUTES, min(MAX_MINUTES, total)) if total else \
        DEFAULT_MINUTES


def make_practice_copy(problem, set_name, area_label, when=None):
    """문제 파일 -> 부분연습_<세트>_<영역>_<일시>.xlsm 사본."""
    when = when or datetime.now()
    stamp = when.strftime("%Y%m%d_%H%M")
    safe_area = re.sub(r"[\\/:*?\"<>|,\s]+", "", str(area_label))[:20]
    d = os.path.dirname(os.path.abspath(problem))
    return copy_as_macro_enabled(
        problem, _unique_stem(d, f"부분연습_{set_name}_{safe_area}_{stamp}"))


# ---------------------------------------------------------------------------
# 재응시 / 오답노트 모드 로직 (GUI 없이 테스트 가능)
# ---------------------------------------------------------------------------


def set_records(set_name, records=None):
    """이 세트의 응시 기록 목록 ('(연습)' 접미 포함 매칭)."""
    records = records if records is not None else load_records()
    return [r for r in records
            if str(r.get("세트명", "")).startswith(str(set_name))]


def problem_has_formula_traces(problem_path, sheet_name="계산작업"):
    """원본 문제 파일의 계산작업 시트에 수식이 몇 개나 있는지 (오염 감지)."""
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
    """세트의 최신 채점결과 JSON 경로 (없으면 None).

    full_only=True면 부분 채점(mode: partial — 부분 연습·오답 재풀이) 결과는
    건너뛰고 전체 응시 결과만 찾습니다.
    """
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
    """채점결과 JSON 로드 -> (전체 dict 또는 None, 오답 항목 목록)."""
    try:
        with open(json_path, encoding="utf-8") as f:
            data = json.load(f)
        items = [i for i in (data.get("wrong_items") or [])
                 if isinstance(i, dict)]
        return data, items
    except Exception:
        return None, []


def find_latest_attempt(set_info):
    """가장 최근 풀이_ 사본 경로 (없으면 None)."""
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
    """이전 풀이를 오답연습_<세트>_<일시>.xlsm 사본으로 복사."""
    when = when or datetime.now()
    stamp = when.strftime("%Y%m%d_%H%M")
    d = os.path.dirname(os.path.abspath(source))
    return copy_as_macro_enabled(
        source, _unique_stem(d, f"오답연습_{set_name}_{stamp}"))


def review_item_text(item):
    """오답노트 항목 상세 텍스트 (패널 표시용)."""
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
    """저장된 이해 체크 인덱스 집합."""
    cfg = load_set_config(path)
    rv = (cfg.get(norm_key) or {}).get("오답연습") or {}
    try:
        return {int(i) for i in (rv.get("이해체크") or {}).get(json_name, [])}
    except Exception:
        return set()


def record_review_state(norm_key, json_name, checks,
                        path=SET_CONFIG_PATH, count_up=False):
    """오답연습 체크 상태(+횟수)를 세트설정.json에 저장.
    점수 기록(기록.json)에는 아무것도 남기지 않습니다."""
    cfg = load_set_config(path)
    ent = cfg.setdefault(norm_key, {})
    rv = ent.setdefault("오답연습", {})
    if count_up:
        rv["횟수"] = int(rv.get("횟수") or 0) + 1
    rv.setdefault("이해체크", {})[str(json_name)] = \
        sorted(int(i) for i in checks)
    return save_set_config(cfg, path)


# ---------------------------------------------------------------------------
# 자동 업데이트 (공개 저장소)
# ---------------------------------------------------------------------------

UPDATE_BASE_URL = ("https://raw.githubusercontent.com/"
                   "xasmine02/cocomate-study-tools/main/")
UPDATE_STAMP = os.path.join(RECORDS_DIR, ".update_check")


def _version_tuple(v):
    nums = re.findall(r"\d+", str(v or ""))
    return tuple(int(x) for x in nums[:3]) if nums else (0,)


def fetch_update_info(base_url=UPDATE_BASE_URL, timeout=3):
    """version.json 조회. 실패/404/오프라인이면 None (조용히 스킵)."""
    import urllib.request
    try:
        req = urllib.request.Request(
            base_url + "version.json",
            headers={"User-Agent": "cocomate-updater"})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            info = json.loads(r.read().decode("utf-8"))
        if isinstance(info, dict) and info.get("version"):
            return info
    except Exception:
        pass
    return None


def update_available(info, current=None):
    return bool(info) and _version_tuple(info.get("version")) > \
        _version_tuple(current if current is not None else __version__)


def should_check_update(stamp_path=None):
    """업데이트 확인 하루 1회 제한. 확인해도 될 때 True + 타임스탬프 갱신."""
    stamp_path = stamp_path or UPDATE_STAMP
    try:
        if time.time() - os.path.getmtime(stamp_path) < 86400:
            return False
    except OSError:
        pass
    try:
        os.makedirs(os.path.dirname(stamp_path), exist_ok=True)
        with open(stamp_path, "w", encoding="utf-8") as f:
            f.write(datetime.now().isoformat())
    except OSError:
        pass
    return True


def _update_target_path(rel_key, base_dir=None):
    """저장소 상대 경로('채점/grade.py') -> 내 설치 위치."""
    base_dir = base_dir or BASE_DIR
    root = os.path.dirname(base_dir)
    parts = [p for p in str(rel_key).split("/") if p not in ("", ".", "..")]
    cand = os.path.join(root, *parts)
    if os.path.isfile(cand):
        return cand
    flat = os.path.join(base_dir, parts[-1])
    if os.path.isfile(flat):
        return flat
    return cand if os.path.isdir(os.path.dirname(cand)) else flat


def apply_update(info, base_url=UPDATE_BASE_URL, base_dir=None, timeout=15):
    """모든 파일 다운로드 -> 검증(__version__ + 문법) -> 원자 교체(.bak 1개).

    반환: (성공 여부, 메시지). 어느 단계든 실패하면 기존 파일로 롤백.
    """
    import py_compile
    import urllib.parse
    import urllib.request
    files = (info or {}).get("files") or {}
    if not files:
        return False, "업데이트 파일 목록이 비어 있습니다."
    staged = []      # (target, tmp)
    replaced = []    # target
    try:
        for rel_key, repo_rel in files.items():
            target = _update_target_path(rel_key, base_dir)
            url = base_url + urllib.parse.quote(str(repo_rel))
            req = urllib.request.Request(
                url, headers={"User-Agent": "cocomate-updater"})
            with urllib.request.urlopen(req, timeout=timeout) as r:
                data = r.read()
            text = data.decode("utf-8")
            if "__version__" not in text:
                raise RuntimeError(f"{repo_rel}: __version__ 표식이 없어 "
                                   "배포 파일이 아닌 것으로 판단")
            tmp = target + ".new"
            os.makedirs(os.path.dirname(target) or ".", exist_ok=True)
            with open(tmp, "wb") as f:
                f.write(data)
            staged.append((target, tmp))  # 실패 시 정리 대상 등록 후 검증
            py_compile.compile(tmp, cfile=tmp + "c", doraise=True)
            try:
                os.remove(tmp + "c")
            except OSError:
                pass
        for target, tmp in staged:
            if os.path.isfile(target):
                shutil.copy2(target, target + ".bak")
            os.replace(tmp, target)
            replaced.append(target)
        return True, (f"{len(replaced)}개 파일을 "
                      f"v{info.get('version')}(으)로 업데이트했습니다.")
    except Exception as e:
        for _target, tmp in staged:
            for leftover in (tmp, tmp + "c"):
                try:
                    if os.path.isfile(leftover):
                        os.remove(leftover)
                except OSError:
                    pass
        for target in replaced:
            bak = target + ".bak"
            try:
                if os.path.isfile(bak):
                    shutil.copy2(bak, target)
            except OSError:
                pass
        return False, f"업데이트 실패(기존 버전 유지): {e}"


def format_elapsed(seconds):
    seconds = max(0, int(seconds))
    m, s = divmod(seconds, 60)
    return f"{m}분 {s:02d}초"


def beep(widget=None):
    """알림음: winsound 우선, 폴백 bell."""
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


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

if HAS_TK:

    class CollapsibleErrorDialog(tk.Toplevel):
        """오류 메시지 + 접을 수 있는 상세(stderr) 표시."""

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
            """메시지+상세를 클립보드로 (채팅에 붙여넣어 진단용)."""
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


    _LAST_ERROR_DIALOG = [None]   # 스모크/진단용: 마지막 오류 대화상자

    def _tk_report_callback_exception(root, exc, val, tb):
        """tkinter 콜백 예외 전역 처리: 로그 append + 접이식 오류 대화상자.

        pythonw에서는 stderr가 없어 예외가 조용히 사라지므로 반드시 흔적을
        남기고 사용자에게 보여 줍니다.
        """
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


    class ResultWindow(tk.Toplevel):
        """채점 결과 창: 큰 점수 + 시트별 점수 + 리포트 열기."""

        def __init__(self, master, result, html_path, folder=None,
                     copied=False, goal=None):
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
                tk.Label(frm, text=f"합격선 {PASS_LINE}점 기준: {verdict}",
                         bg=BG, fg=INK,
                         font=UI_FONT_BOLD).pack(pady=(0, 10))
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
                    g_txt = f"오늘 목표 {goal}점 달성! (+{diff}점)"
                    g_fg, g_bg = BRAND_DARK, BRAND_SOFT
                else:
                    g_txt = f"오늘 목표 {goal}점까지 {-diff}점 — 오답노트로 " \
                            "복습 후 재도전!"
                    g_fg, g_bg = "#8A5A00", "#FBF0DC"
                tk.Label(frm, text=g_txt, bg=g_bg, fg=g_fg,
                         font=("Malgun Gothic", 9, "bold"), padx=10, pady=5,
                         wraplength=360).pack(pady=(6, 0))
            if copied:
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


    class TimerWindow(tk.Toplevel):
        """항상 위 소형 타이머 창."""

        ALERTS = (600, 300, 60)  # 10분/5분/1분

        def __init__(self, app, exam):
            super().__init__(app)
            self.app = app
            self.exam = exam            # 진행 중 시험 정보 dict
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
            self._tick()

        def _fmt(self):
            m, s = divmod(max(0, self.remaining), 60)
            return f"{m:02d}:{s:02d}"

        def _color(self):
            if self.remaining <= 300:
                return "#FF8B80"      # 빨강 계열
            if self.remaining <= 600:
                return "#F2B24C"      # 호박 계열
            return "#7BD59A"          # 초록 계열

        def _tick(self):
            if self.finished:
                return
            if not self.paused:
                self.time_lbl.configure(text=self._fmt(), fg=self._color())
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

        def _time_up(self):
            self.finished = True
            self.remaining = 0
            self.time_lbl.configure(text="00:00", fg="#FF8B80")
            self.status_lbl.configure(text="시험 시간 종료")
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
            self.app.start_grading(self.exam, elapsed, self.practice,
                                   on_done=self._grading_done)

        def _grading_done(self):
            try:
                self.destroy()
            except Exception:
                pass
            self.app.exam_closed()


    class ReviewWindow(tk.Toplevel):
        """오답노트 모드 패널: 틀린 항목 리스트 + 상세 + 이해 체크.

        타이머 없음, 기록.json에 아무것도 남기지 않음.
        """

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
                                self.json_name, self.checks)
            self._refresh_list()

        def _retake(self):
            self.destroy()
            self.app.start_exam()


    class PracticeDialog(tk.Toplevel):
        """부분 연습 영역 선택 대화상자.

        확인 시 self.result = {"sheets": [...], "label": "...", "minutes": n}
        """

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
                v = tk.BooleanVar(value=False)
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
                    v = tk.BooleanVar(value=False)
                    self.sheet_vars[sh] = v
                    tk.Checkbutton(self.adv_frame, text=sh, variable=v,
                                   bg=BG, fg=INK, activebackground=BG,
                                   font=("Malgun Gothic", 9)
                                   ).pack(anchor="w", padx=16)
            tf = tk.Frame(frm, bg=BG)
            tf.pack(fill="x", pady=(10, 0))
            tk.Label(tf, text="연습 시간(분):", bg=BG, fg=INK,
                     font=UI_FONT).pack(side="left")
            self.minutes_var = tk.IntVar(value=DEFAULT_MINUTES)
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
            """(시트 목록, 영역 라벨). 개별 시트 체크가 있으면 우선."""
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
        """세트 인식 진단 결과 표시 + 클립보드 복사."""

        def __init__(self, app, text):
            super().__init__(app)
            self.text_value = text
            self.title(f"{APP_TITLE} - 세트 인식 진단")
            self.configure(bg=BG)
            self.geometry("760x520")
            frm = tk.Frame(self, bg=BG, padx=14, pady=10)
            frm.pack(fill="both", expand=True)
            tk.Label(frm, text="파일명 → 정규화 키 → 역할 → 세트 → 일정 슬롯 "
                     "매칭 결과입니다. [복사]해서 채팅에 붙여넣으면 진단해 "
                     "드립니다.", bg=BG, fg=INK, font=("Malgun Gothic", 9),
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


    class SetChooserDialog(tk.Toplevel):
        """[다른 세트로 바꾸기] — 슬롯을 고르고 목록에서 세트 선택."""

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
            self.slot_var = tk.IntVar(value=1)
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
        """오늘 일정 단계 가이드 — 매일 '완주 → 채점 → 오답노트 → 오답 재풀이'.

        세트 슬롯(고정/자동 선택)을 창을 열 때 확정해 위쪽에 이유와 함께
        표시하고, [다른 세트로 바꾸기]로 바꿀 수 있습니다. 진행 상태는
        세트설정.json '_진행'에 날짜 키(d01~)로 저장되어 이어집니다.
        """

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
            # 미발견 슬롯: 빨간 안내 + [직접 선택] [파일명 안내] (항상 살아 있는 버튼)
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
            tk.Button(bf, text="닫기", font=UI_FONT, relief="groove",
                      padx=12, pady=4, command=self.destroy).pack(
                side="right")
            self._render_header()
            self.refresh(select=self.current_index())

        # --- 세트 슬롯 ---

        def _resolve_slots(self, save=False):
            """세트 슬롯 확정 (직접 선택 매핑/자동 선택 복원 → 매칭 → 저장).

            해석 중 예외가 나도 창은 살아 있어야 하므로 실패 슬롯은
            None + 사유로 두고 오류 로그에 남깁니다.
            """
            plan = self.plan
            if plan.get("세트"):
                try:
                    saved = load_auto_picks(self.day_tag) if self.day_tag \
                        else []
                    self.slot_sets = resolve_day_sets(
                        plan, self.app.sets, load_records(), saved=saved,
                        mapping=load_slot_mapping())
                    for s, _r in self.slot_sets:   # 매핑으로 만든 직접 세트 편입
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
            """세트 선택 결과·이유 표시 문구 (한 줄에 슬롯 하나)."""
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
            """세트를 찾지 못한 슬롯 번호 목록."""
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
            """미발견 슬롯 -> 파일 선택 대화상자로 문제/정답 지정."""
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
            """직접 선택한 세트를 슬롯 k에 적용 + 매핑 저장 + 목록 편입."""
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
            """[다른 세트로 바꾸기] 대화상자."""
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
            """슬롯 seat를 세트 s로 교체 (같은 세트가 다른 슬롯에 있으면 맞교환)."""
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

        # --- 상태 ---

        def _all_done(self):
            return bool(self.steps) and \
                all(i in self.done for i in range(len(self.steps)))

        def current_index(self):
            """첫 미완료 스텝 (전부 완료면 마지막)."""
            for i in range(len(self.steps)):
                if i not in self.done:
                    return i
            return max(0, len(self.steps) - 1)

        def selected_index(self):
            sel = self.listbox.curselection()
            return sel[0] if sel and sel[0] < len(self.steps) else None

        def refresh(self, select=None):
            """목록·진행 라벨 갱신. select 지정 시 그 스텝 선택."""
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
            n_done = len([i for i in self.done
                          if 0 <= i < len(self.steps)])
            self.progress_lbl.configure(
                text=f"{n_done}/{len(self.steps)} 완료 · 체크는 자동 저장 "
                     "(껐다 켜도 이어짐)")
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
            lines = [f"{i + 1}. {st['이름']}  ·  {kind}  ·  예상 "
                     f"{st.get('분', '?')}분  ·  {state}",
                     "", st.get("설명", "")]
            lines.append("")
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
                lines.append("[이 단계 시작]을 누르면 오답노트 모드가 열립니다. "
                             "다 보고 나면 [완료 체크]를 누르세요.")
            elif kind == "오답재풀이":
                lines.append("[이 단계 시작]을 누르면 오답 시트 목록으로 새 "
                             "사본과 15분 타이머가 열리고, 제출하면 그 "
                             "시트들만 채점된 뒤 자동 체크됩니다.")
            elif kind in ("부분연습", "드릴"):
                lines.append("[이 단계 시작]을 누르면 풀이 사본과 타이머가 "
                             "열립니다. 채점까지 끝나면 자동으로 "
                             "체크됩니다.")
            self.detail_lbl.configure(text="\n".join(lines))

        # --- 동작 ---

        def _auto_steps_for(self, i):
            """채점 완료 시 함께 체크할 스텝: 자신 + 바로 뒤의 '채점' 스텝."""
            out = [i]
            if i + 1 < len(self.steps) and \
                    self.steps[i + 1].get("형") == "채점":
                out.append(i + 1)
            return out

        def start_step(self):
            """[이 단계 시작] — 어떤 예외도 조용히 죽지 않고 안내로 표시."""
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
            # 다음 미완료 스텝으로 자동 포커스
            self.refresh(select=self.current_index())
            self.maybe_celebrate()

        def mark_step_done(self, idx):
            """외부(채점 완료)에서 스텝 자동 체크."""
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
        """시작 화면."""

        def __init__(self, scan_root=None):
            super().__init__()
            self.title(APP_TITLE)
            self.configure(bg=BG)
            self.minsize(680, 520)
            self.scan_root = scan_root or default_scan_root()
            self.sets = []
            self.grade_py = find_grade_py()
            self.exam_running = False
            self._grade_state = None
            self._update_info = None
            self._update_prompted = False
            self._warned_dirty = set()   # 원본 오염 경고를 이미 띄운 세트
            self.plan_no = routine_day_no()   # 오늘의 학습 일정 번호
            self._pending_plan = None
            self.step_guide = None            # 단계 가이드 창 (열려 있으면)
            self._build_ui()
            self.refresh_sets()
            self.refresh_records()
            threading.Thread(target=self._bg_update_check,
                             daemon=True).start()
            self.after(1200, self._update_poll)

        # ---------------- UI 구성 ----------------

        def _build_ui(self):
            header = tk.Frame(self, bg=BRAND)
            header.pack(fill="x")
            tk.Label(header, text=APP_TITLE, bg=BRAND, fg="white",
                     font=("Malgun Gothic", 16, "bold")).pack(
                side="left", padx=18, pady=10)
            tk.Label(header, text="컴활 2급 실기 모의고사 런처", bg=BRAND,
                     fg="#CFE9DA", font=UI_FONT).pack(side="left")

            # 오늘의 학습 카드 (날짜 기반 루틴 일정)
            plan_card = tk.Frame(self, bg=BRAND_SOFT, padx=14, pady=8)
            plan_card.pack(fill="x", padx=16, pady=(10, 0))
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

            body = tk.Frame(self, bg=BG, padx=16, pady=12)
            body.pack(fill="both", expand=True)
            body.columnconfigure(0, weight=3)
            body.columnconfigure(1, weight=2)
            body.rowconfigure(1, weight=1)

            tk.Label(body, text="응시 가능한 모의고사", bg=BG, fg=INK,
                     font=UI_FONT_BOLD).grid(row=0, column=0, sticky="w")
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
            self.info_lbl = tk.Label(
                body, text="왼쪽 목록에서 세트를 선택하세요.", bg=CARD,
                fg=SUB, font=UI_FONT, justify="left", anchor="nw",
                padx=10, pady=8, highlightbackground=LINE,
                highlightthickness=1, wraplength=230)
            self.info_lbl.grid(row=1, column=1, sticky="nsew",
                               padx=(12, 0), pady=(4, 8))

            ctrl = tk.Frame(body, bg=BG)
            ctrl.grid(row=2, column=0, columnspan=2, sticky="ew")
            tk.Label(ctrl, text="시험 시간(분):", bg=BG, fg=INK,
                     font=UI_FONT).pack(side="left")
            self.minutes_var = tk.IntVar(value=DEFAULT_MINUTES)
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
            tk.Button(ctrl, text="업데이트 확인", font=UI_FONT, relief="groove",
                      padx=10, pady=4,
                      command=self.manual_update_check).pack(
                side="right", padx=4)
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

        # ---------------- 데이터 갱신 ----------------

        def refresh_all(self):
            self.refresh_sets()
            self.refresh_records()

        def refresh_sets(self):
            direct = [s for s in self.sets if s.get("direct")]
            scanned = scan_sets(self.scan_root)
            norms = {s["norm"] for s in scanned}
            self.sets = scanned + [s for s in direct
                                   if s.get("norm") not in norms]
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
                self._render_plan_card()   # 세트 편입/자동 선택 반영

        def refresh_records(self):
            records = load_records()
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
                               and mx else score_s)
                    score_s += f" [{r.get('mode')} · {r.get('영역', '?')}]"
                lines.append(f"{r.get('일시', '?')}  |  {r.get('세트명', '?')}"
                             f"  |  {score_s}  |  {r.get('소요시간', '-')}")
            self.records_lbl.configure(text="\n".join(lines))

        def _selected_set(self):
            sel = self.listbox.curselection()
            if not sel or sel[0] >= len(self.sets):
                return None
            return self.sets[sel[0]]

        def _show_info(self):
            s = self._selected_set()
            if not s:
                return
            recs = set_records(s["name"])
            self.start_btn.configure(
                text="재응시 (새로 시작)" if recs else "시험 시작")
            self.review_btn.configure(
                state="normal" if find_latest_result_json(s) else "disabled")
            summ = records_summary(load_records()).get(s["name"], {})
            best = summ.get("best")
            recent = summ.get("recent") or []
            lines = [
                f"세트: {s['name']}",
                f"폴더: {s['dir']}",
                f"정답 파일: 있음",
                f"기대값 JSON: {'있음' if s['key'] else '없음'}",
                "문제지 PDF: " + (os.path.basename(s["pdf"]) if s["pdf"]
                                else "미연결 ([문제지 연결]로 지정 가능)"),
                "",
                f"응시 기록: {len(recs)}회",
                f"최고 점수: {best if best is not None else '-'}",
                f"최근 3회: "
                + (" → ".join(str(x) for x in recent) if recent else "-"),
            ]
            self.info_lbl.configure(text="\n".join(lines), fg=INK)

        # ---------------- 자동 업데이트 ----------------

        def _bg_update_check(self):
            try:
                if should_check_update():
                    info = fetch_update_info()
                    if update_available(info):
                        self._update_info = info
            except Exception:
                pass

        def _update_poll(self, tries=0):
            if self._update_info and not self._update_prompted:
                self._update_prompted = True
                self._prompt_update(self._update_info)
                return
            if tries < 12:
                self.after(1000, lambda: self._update_poll(tries + 1))

        def _prompt_update(self, info):
            notes = info.get("notes") or "-"
            if not messagebox.askyesno(
                    APP_TITLE,
                    f"새 버전 v{info.get('version')}가 있습니다 — 지금 "
                    f"업데이트할까요?\n변경: {notes}", parent=self):
                return
            prog = tk.Toplevel(self)
            prog.title(APP_TITLE)
            prog.attributes("-topmost", True)
            tk.Label(prog, text="업데이트를 내려받는 중입니다...",
                     padx=26, pady=18, font=UI_FONT).pack()
            prog.update()
            ok, msg = apply_update(info)
            prog.destroy()
            if ok:
                messagebox.showinfo(
                    APP_TITLE, msg + "\n\n프로그램을 다시 시작해 주세요.",
                    parent=self)
                self._try_restart()
            else:
                CollapsibleErrorDialog(
                    self, APP_TITLE,
                    "업데이트에 실패해 기존 버전을 유지합니다.", msg)

        def _try_restart(self):
            if self.exam_running:
                return  # 시험 중에는 재시작하지 않음
            script = os.path.join(BASE_DIR, "시험장.py")
            try:
                self.destroy()
            except Exception:
                pass
            try:
                os.execl(sys.executable, sys.executable, script)
            except Exception:
                pass

        def manual_update_check(self):
            info = fetch_update_info()
            if update_available(info):
                self._update_prompted = True
                self._prompt_update(info)
            else:
                messagebox.showinfo(
                    APP_TITLE, f"현재 최신 버전입니다 (v{__version__}).",
                    parent=self)

        # ---------------- 시험 흐름 ----------------

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
            s["pdf"] = os.path.abspath(path)
            remember_set(s)
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
            """시험 시작. practice={"label","sheets","minutes"}면 부분 연습."""
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
            # 재응시: 새 깨끗한 사본으로 시작함을 명시 (시험 모드만)
            if not practice and set_records(s["name"]) \
                    and not messagebox.askokcancel(
                    APP_TITLE,
                    "재응시: 원본 문제 파일에서 새 깨끗한 사본으로 "
                    "시작합니다.\n이전 풀이는 풀이_*.xlsx 파일로 그대로 "
                    "보관됩니다.\n계속할까요?", parent=self):
                return
            # 원본 오염 감지 (세트당 경고 1회)
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
            # 매크로 안내 (세트당 1회, 매크로작업 포함 시)
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
            try:
                if practice and practice.get("mode") == "오답재풀이":
                    student = make_retry_copy(s["problem"], s["name"])
                elif practice:
                    student = make_practice_copy(s["problem"], s["name"],
                                                 practice["label"])
                else:
                    student = make_attempt_copy(s["problem"], s["name"])
            except Exception as e:
                log_error(f"풀이 사본 생성: {s.get('name')}", e)
                CollapsibleErrorDialog(
                    self, APP_TITLE,
                    f"'{s.get('name')}' 세트를 시작하지 못했습니다: 풀이 사본을 "
                    f"만들 수 없습니다 ({e})", log_error("", e))
                return
            ok, err = open_file(student)
            if not ok:
                CollapsibleErrorDialog(
                    self, APP_TITLE,
                    "풀이 파일을 Excel로 열지 못했습니다.\n"
                    "Excel(또는 호환 프로그램)이 설치되어 있는지 확인한 뒤,\n"
                    "아래 파일을 직접 열어 풀이를 진행하세요:\n\n" + student,
                    err)
            if s["pdf"]:
                open_file(s["pdf"])
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
            }
            self.exam_running = True
            self.start_btn.configure(state="disabled", text="진행 중")
            TimerWindow(self, exam)

        def exam_closed(self):
            self.exam_running = False
            self.start_btn.configure(state="normal", text="시험 시작")
            self.refresh_records()
            self._show_info()

        # ---------------- 오늘의 학습 ----------------

        def _render_plan_card(self):
            plan = plan_for_day(self.plan_no)
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
            if plan.get("목표"):
                todo += f"  [목표 {plan['목표']}점]"
            if reasons:
                todo += "\n자동 선택: " + " / ".join(reasons)
            steps = plan.get("스텝")
            if steps:
                tag = plan_day_tag(plan["no"])
                n_done = len([i for i in (load_step_progress(tag)
                                          if tag else set())
                              if 0 <= i < len(steps)])
                todo += f"  [스텝 {n_done}/{len(steps)} 완료]"
            self.plan_todo_lbl.configure(text=todo)
            today_no = routine_day_no()
            self.plan_start_btn.configure(
                text="오늘 일정 시작" if self.plan_no == today_no
                else "이 일정 시작")
            try:  # 신뢰 위치 미등록이면 1줄 안내
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
            """세트를 목록에 편입 (같은 norm이 있으면 무시). 직접 선택 세트용."""
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
            """[세트 인식 진단] — 텍스트 표 + 클립보드 복사."""
            try:
                text = scan_diagnosis_text(self.scan_root, sets=self.sets)
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
            """오늘(또는 미리보기 중인) 일정의 단계 가이드 창 열기."""
            plan = plan_for_day(self.plan_no)
            if not plan.get("스텝"):
                # 스텝이 없는 날(시험일/루틴 종료)은 안내만
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

        # ---------------- 부분 연습 모드 ----------------

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

        # ---------------- 오답노트 모드 ----------------

        def open_review_mode(self, full_only=False):
            s = self._selected_set()
            if not s:
                messagebox.showinfo(APP_TITLE, "먼저 세트를 선택하세요.",
                                    parent=self)
                return
            jp = find_latest_result_json(s, full_only=full_only)
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
            # 이전 풀이 사본 -> 오답연습 사본으로 열기
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

        # ---------------- 채점 ----------------

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
            """결과 JSON을 클립보드에 복사 — 루틴 웹 원클릭 연동용.

            실패(클립보드 잠김 등)는 조용히 무시하고 False 반환."""
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
                        # 설치 성공 -> 자동 재시도
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
            else:
                score = result.get("total")
                copied = self._copy_result_to_clipboard(result)
                goal = (exam.get("plan") or {}).get("목표")
                ResultWindow(self, result, html_path,
                             folder=os.path.dirname(html_path),
                             copied=copied, goal=goal)
                if os.path.isfile(html_path):
                    open_file(html_path)
            pinfo = exam.get("practice_info")
            record = {
                "일시": exam["started"].strftime("%Y-%m-%d %H:%M"),
                "세트명": exam["set"]["name"] + (" (연습)" if practice else ""),
                "점수": score,
                "소요시간": format_elapsed(elapsed),
                "리포트": html_path if os.path.isfile(html_path) else None,
                "mode": (pinfo.get("mode") or "부분연습") if pinfo else "시험",
            }
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
                record["day"] = day_tag        # 새 체계 d01~d14
                record["루틴"] = ROUTINE_TAG   # 구 루틴 기록과 구분
            try:
                append_record(record)
            except OSError as e:
                messagebox.showwarning(
                    APP_TITLE, f"기록.json 저장에 실패했습니다: {e}",
                    parent=self)
            # 단계 가이드에서 시작한 스텝: 채점 완료 시 자동 체크
            plan_info = exam.get("plan") or {}
            step_idx = plan_info.get("step")
            done_steps = plan_info.get("done_steps") or (
                [step_idx] if step_idx is not None else [])
            if day_tag and done_steps and score is not None:
                try:
                    g = getattr(self, "step_guide", None)
                    if g is not None and g.winfo_exists() \
                            and g.day_tag == day_tag:
                        for idx in done_steps:      # 시험 + 채점 스텝 자동 체크
                            g.mark_step_done(int(idx))
                    else:
                        done = load_step_progress(day_tag)
                        done.update(int(idx) for idx in done_steps)
                        save_step_progress(day_tag, done)
                    self._render_plan_card()
                except Exception:
                    pass


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------


def run_smoke():
    """GUI 스모크 테스트: 창 생성 -> 위젯 렌더 -> 타이머 창 -> destroy."""
    app = ExamApp()
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
    # 오답노트 패널 스모크
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
    review.checks = set()        # 이전 스모크가 저장한 체크 상태 무시(결정성)
    review._refresh_list()
    review.listbox.selection_set(0)
    review.toggle_check()
    assert review.retake_btn.cget("state") == "normal"  # 1/1 체크 완료
    review.destroy()
    # 부분 연습 대화상자 스모크 (부분 연습 모드는 일정 밖 기능으로 유지)
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
    # 오늘의 학습 카드 + 미리보기 화살표 (일정 v4: 9/3 시작, 시험 2회 D-day)
    assert app.plan_title_lbl.cget("text"), "오늘의 학습 제목 비어 있음"
    app.plan_no = 3
    app._render_plan_card()
    t3 = app.plan_title_lbl.cget("text")
    assert "Day 3" in t3 and "9/5(토)" in t3 and "1차 완주 — " in t3, t3
    assert "시험1 D" in t3 and "시험2 D" in t3, t3
    app._shift_plan(1)
    assert "Day 4" in app.plan_title_lbl.cget("text")
    app.plan_no = 8
    app._shift_plan(1)                       # d08 다음은 시험 1 (시간순)
    assert app.plan_no == PLAN_EXAM1
    assert "시험일" in app.plan_title_lbl.cget("text")
    app._shift_plan(1)
    assert app.plan_no == 9
    # 단계 가이드 창 — 2세트 날(Day 3): ①~④ ×2 + (선택) 퀴즈 = 9스텝
    save_step_progress("d03", set())
    guide = StepGuideWindow(app, plan_for_day(3))
    app.update_idletasks()
    app.update()
    assert guide.listbox.size() == 9, guide.listbox.size()
    assert "0/9" in guide.progress_lbl.cget("text")
    assert guide.current_index() == 0
    assert "▶" in guide.listbox.get(0)
    assert "세트 1:" in guide.pick_lbl.cget("text"), guide.pick_lbl.cget("text")
    guide.listbox.selection_clear(0, "end")
    guide.listbox.selection_set(0)
    guide.toggle_check()          # 1단계 완료 체크
    assert 0 in guide.done and "1/9" in guide.progress_lbl.cget("text")
    assert guide.selected_index() == 1, "다음 미완료 스텝 자동 포커스"
    assert "[v]" in guide.listbox.get(0) and "▶" in guide.listbox.get(1)
    guide.mark_step_done(1)       # 채점 완료 자동 체크 경로
    assert "2/9" in guide.progress_lbl.cget("text")
    assert load_step_progress("d03") == {0, 1}, "진행 상태 저장"
    guide.destroy()
    guide2 = StepGuideWindow(app, plan_for_day(3))   # 재시작 후 이어하기
    app.update_idletasks()
    app.update()
    assert guide2.done == {0, 1} and guide2.current_index() == 2
    guide2.destroy()
    save_step_progress("d03", set())
    # [이 단계 시작] -> start_exam 배선 (start_exam은 스텁으로 대체)
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
    guide3 = StepGuideWindow(app, plan_for_day(1))   # d01: 2024 상시 1회
    assert guide3.slot_sets[0][0] is fake_s, guide3.pick_text()
    assert "일정 지정" in guide3.pick_lbl.cget("text")
    guide3.listbox.selection_clear(0, "end")
    guide3.listbox.selection_set(0)          # 모의 스텝
    guide3.start_step()
    assert len(calls) == 1, "모의 스텝 -> start_exam 호출"
    practice_arg, pend = calls[0]
    assert practice_arg is None and pend == {
        "day": "d01", "목표": None, "step": 0, "done_steps": [0, 1]}, \
        (practice_arg, pend)
    assert app.minutes_var.get() == 40
    guide3.destroy()
    save_step_progress("d01", set())
    # 자동 선택 날(Day 8): 기록 없음 -> 신규 세트 우선 + 이유 표시 + 세트 바꾸기
    save_step_progress("d08", set())
    save_auto_picks("d08", [])
    guide4 = StepGuideWindow(app, plan_for_day(8))
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
    assert load_auto_picks("d08")[0]["세트"] == other["name"], "선택 저장"
    guide4.destroy()
    save_auto_picks("d08", [])
    save_step_progress("d08", set())
    app.sets, app.start_exam = saved_sets, saved_start
    app._pending_plan = None
    # v2.1.1: 2세트 날(Day 3) 한 슬롯만 미발견 -> 다른 슬롯 정상 시작 + 직접 선택 대체
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
    guide5 = StepGuideWindow(app, plan_for_day(3))
    app.update_idletasks()
    app.update()
    assert guide5.missing_slots() == [0], guide5.pick_text()
    assert guide5.slot_sets[1][0] is fake_s2r
    assert len(guide5.missing_rows) == 1 and \
        guide5.missing_rows[0][2].cget("text") == "직접 선택"
    assert "찾지 못함" in guide5.pick_text() and \
        guide5.pick_lbl.cget("fg") == RED
    guide5.listbox.selection_clear(0, "end")
    guide5.listbox.selection_set(0)          # 미발견 슬롯의 모의 스텝
    _orig_info = messagebox.showinfo
    infos = []
    messagebox.showinfo = lambda *a, **k: infos.append(a)   # 모달 차단 방지
    try:
        guide5.start_step()                  # -> missing 안내, 예외 없음
    finally:
        messagebox.showinfo = _orig_info
    assert calls == [], "미발견 슬롯은 시작되지 않음"
    assert infos and "다른 세트로 바꾸기" in infos[0][1], infos
    guide5.listbox.selection_clear(0, "end")
    guide5.listbox.selection_set(4)          # 슬롯 2(상시 2회) 모의 스텝
    guide5.start_step()
    assert len(calls) == 1 and calls[0][1]["step"] == 4, "정상 슬롯은 시작"
    assert app._selected_set() is fake_s2r
    guide5.apply_direct_slot(0, fake_A)      # 미발견 슬롯 직접 선택 대체
    assert guide5.missing_slots() == [] and not guide5.missing_rows
    assert guide5.steps[0]["이름"].startswith("컴활2급 A형")
    assert load_slot_mapping()["2024 A형"]["name"] == "컴활2급 A형", "매핑 저장"
    assert any(x["norm"] == "smokeA" for x in app.sets), "목록 편입"
    guide5.listbox.selection_clear(0, "end")
    guide5.listbox.selection_set(0)
    guide5.start_step()
    assert len(calls) == 2 and app._selected_set() is fake_A
    guide5.destroy()
    # 스텝 시작 중 예외 -> 사용자 메시지 + 로그, 창 생존
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
    cfg_after.pop("smokeA", None)            # remember_set 잔여 제거
    save_set_config(cfg_after)
    save_step_progress("d03", set())
    app.sets, app.start_exam = saved_sets, saved_start
    app._pending_plan = None
    # 전역 콜백 예외 처리: 로그 append + 오류 대화상자(복사 버튼)
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
    # 세트 인식 진단 창
    diag = DiagnosisWindow(app, scan_diagnosis_text(app.scan_root, sets=app.sets))
    app.update_idletasks()
    app.update()
    assert "일정 슬롯 매칭" in diag.text_value and diag.copy() is True
    diag.destroy()
    # 스텝 실행 매핑 (세트 없는 환경 -> 모의는 missing, 채점은 info)
    kind, _p = resolve_step_action(plan_for_day(1)["스텝"][0], saved_sets)
    assert kind in ("exam", "missing")
    kind, _p = resolve_step_action(plan_for_day(1)["스텝"][1], saved_sets)
    assert kind == "info"
    kind, _p = resolve_step_action(plan_for_day(0)["스텝"][0], saved_sets)
    assert kind == "info"
    # 클립보드 브리지 라운드트립
    fake = {"total": 87, "mode": "full", "graded_sheets": ["계산작업"]}
    assert app._copy_result_to_clipboard(fake) is True
    back = app.clipboard_get()
    import json as _json
    assert _json.loads(back) == fake, back[:80]
    timer.finished = True
    timer.destroy()
    app.destroy()
    print("SMOKE OK: 창 생성/위젯 렌더/타이머/오답노트 패널/단계 가이드(세트 "
          "자동 선택·바꾸기·미발견 직접 선택)/오류 대화상자·로그/진단 창/"
          "파괴 정상")


def _notify_no_tk():
    """tkinter가 없을 때 안내: 콘솔이 있으면 print, 없으면(pythonw) 메시지 창."""
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


def main():
    ap = argparse.ArgumentParser(description=APP_TITLE)
    ap.add_argument("--scan-root", help="모의고사 스캔 폴더 (기본: 상위 폴더)")
    ap.add_argument("--smoke", action="store_true",
                    help="GUI 스모크 테스트 후 종료 (개발용)")
    args = ap.parse_args()
    if not HAS_TK:
        _notify_no_tk()
        return 1
    if args.smoke:
        run_smoke()
        return 0
    app = ExamApp(scan_root=args.scan_root)
    app.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
