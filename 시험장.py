#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
코코 시험장 — 컴활 2급 실기 모의고사 런처

시작 화면에서 모의고사 세트를 고르고 [시험 시작]을 누르면
문제 파일 사본이 Excel로 열리고 타이머가 시작됩니다.
제출하면 grade.py로 자동 채점해 점수와 리포트를 보여 줍니다.
오늘의 학습(적응형 일정 v5, 9/3 시작·시험 2회): 매일 모의고사 1세트 40분 완주
→ 채점 → 오답노트 → 오답 재풀이. 기록.json 기준으로 매번 재계산해 밀린 날의
세트를 자동 재배치하고, 시험 전날까지 전 세트 완주를 보장하도록 용량을 올립니다.

v2.2.1: Excel 실행 강화 — Windows에서 EXCEL.EXE를 직접 찾아 실행하고 5초 뒤
프로세스를 확인해 안 떴으면 재열기 안내, 타이머 [풀이 파일 열기], 시작 로그
(채점결과/시험장_시작로그.txt), 문제지 PDF 회차 검증, 시작·재시작 안정화.
v2.2.2: 세트 기대값 JSON 퍼지 연결 — 파일명 키가 달라도 연도·회차 토큰이 같은
'*기대값*.json'을 문제지 PDF와 같은 규칙(유일할 때만)으로 세트에 연결.
v2.3.0: 완전 자동 업데이트 — 실행할 때마다 백그라운드로 version.json을 확인해
새 버전이면 묻지 않고 내려받아 검증(sha256·py_compile·json)·백업·적용하고
자동 재시작(시험 중이면 종료 후). 기대값 JSON(data_files)은 루트/기대값/에
자동 배포·갱신되고 세트에 자동 연결. 세트설정 `_설정.자동업데이트` 토글.
v2.3.1: 세트 인식 전면 수정 — 일정 슬롯↔세트를 전역 유일 배정(연도 충돌은 슬롯
전체 토큰 기준, 정확 > 식별 > 부분 일치, 한 세트는 한 슬롯, 동점은 미배정+사유),
문제 파일 내용(SHA-256)이 같은 세트 병합, 문제지 PDF·기대값 JSON 전역 유일 연결,
세트 토큰 없는 PDF는 진단에서 '무관'으로 접음. 코코 모의고사 1·2회 세트 파일
(xlsx·pdf)을 version.json `set_files`로 루트/모의고사/에 자동 배포.
v2.4.0: 루틴 웹 연동 — 프로그램 안에 로컬 HTTP 서버(127.0.0.1:8765, 사용 중이면
빈 포트)를 띄워 「2주 루틴」 페이지(시험장/루틴.html, 자동 업데이트로 갱신)를
서빙하고, 규약(문서/연동_API.md v2.4.1)대로 일정·세트·기록·오답노트·체크 상태를
/api/state 로 내보내며 체크·수동 점수·오답노트 완료·시험 시작 등 쓰기 요청을
받습니다(세션 토큰). [루틴 열기] 버튼·`_설정.루틴자동열기`·단계 가이드
[웹에서 퀴즈 풀기]. 채점 결과는 클립보드 복사 대신 루틴 페이지에 자동 반영.
v2.4.1: 시험일 변경 9/17(목)·9/18(금) 이틀 연속 — 학습일 Day 1~14 = 9/3~9/16(마감
9/16), 구간은 시험일 목록에서 도출(학습일 없는 구간은 건너뜀 → 단일 구간), 9/12
복기일 개념 제거(시험 1 저녁 복기 메모·실수 노트는 시험일 안내), 재응시 목표
사다리 70→75→80→85 는 재응시가 배정된 날짜 순으로 분배(같은 날 2세트는 같은
목표), 남는 용량은 재응시로 채움. 고정 일정표·안내문·D-day·루틴 페이지 동기화.

의존성: Python 표준 라이브러리 + tkinter (채점은 grade.py/openpyxl 필요)
"""

__version__ = "2.4.1"

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
STARTUP_LOG_PATH = os.path.join(RECORDS_DIR, "시험장_시작로그.txt")
STARTUP_LOG_KEEP = 200          # 시작 로그 최근 줄 수


def _ensure_records_home():
    """기록.json 보관 폴더 생성 + 구버전 위치의 기록 자동 이전."""
    try:
        os.makedirs(RECORDS_DIR, exist_ok=True)
        if os.path.isfile(_OLD_RECORDS_PATH) \
                and not os.path.isfile(RECORDS_PATH):
            shutil.move(_OLD_RECORDS_PATH, RECORDS_PATH)
    except OSError:
        pass

def startup_log(message, path=None, keep=STARTUP_LOG_KEEP):
    """시작 로그 1줄 append (채점결과/시험장_시작로그.txt, 최근 keep줄 유지).

    프로그램 시작·세트 선택·사본 생성·Excel/PDF 실행·예외처럼 "무슨 일이
    있었는지"를 pythonw(콘솔 없음)에서도 남깁니다. 실패는 조용히 무시.
    반환: 기록한 한 줄.
    """
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{stamp}] {__version__} " + str(message).replace("\n", " | ")
    p = path or STARTUP_LOG_PATH
    try:
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
    """시작 로그 전체 텍스트 (없으면 안내문)."""
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
    if path is None:   # 시작 로그에도 한 줄 요약 (흐름 추적용)
        last = detail.strip().splitlines()[-1] if detail.strip() else ""
        startup_log(f"예외: {context}" + (f" — {last}" if last else ""))
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
EXPECTED_DIR_NAME = "기대값"      # 자동 배포되는 세트별 기대값 JSON 폴더

# 루틴 웹 연동 서버 (v2.4.0, 규약: 문서/연동_API.md v2.4.1)
ROUTINE_API_VERSION = "2.4.1"          # /api/state 의 version (규약 버전)
ROUTINE_HTML_NAME = "루틴.html"        # 시험장 폴더의 루틴 페이지 파일
ROUTINE_DATA_KEY = "시험장/루틴.html"   # version.json set_files 키 (자동 업데이트)
ROUTINE_PORT_DEFAULT = 8765
ROUTINE_PORT_SETTING = "루틴포트"        # 세트설정 _설정: 실제로 쓴 포트
ROUTINE_AUTO_OPEN_SETTING = "루틴자동열기"   # 세트설정 _설정: 시작 시 브라우저 열기
WEB_CHECKS_KEY = "_웹체크"              # 세트설정: 웹 체크 상태 {키: true}


def expected_values_dir(base_dir=None):
    """자동 업데이트로 내려받는 기대값 JSON 폴더.

    설치 구조가 <루트>/시험장/시험장.py + <루트>/채점/grade.py 이면
    <루트>/기대값/, 한 폴더에 전부 있는 평면 구조면 <시험장 폴더>/기대값/.
    """
    base_dir = base_dir or BASE_DIR
    root = os.path.dirname(base_dir)
    if os.path.basename(base_dir) == "시험장" \
            or os.path.isdir(os.path.join(root, "채점")):
        return os.path.join(root, EXPECTED_DIR_NAME)
    return os.path.join(base_dir, EXPECTED_DIR_NAME)


def key_link_label(s):
    """세트 정보 패널용 기대값 표시: '파일명', '파일명 (자동 연결)', '없음'.

    세트 폴더 안의 정확 키(같은 정규화 키) 파일만 기본 연결로 보고, 그 밖의
    위치(루트/기대값 폴더·다른 폴더)나 토큰 일치로 연결된 파일은 '(자동 연결)'.
    """
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


_YEAR4_RE = re.compile(r"20\d{2}")
_YEAR2_RE = re.compile(r"2\d")
_ROUND_RE = re.compile(r"\d{1,2}회")
_FORM_RE = re.compile(r"[ab가나]형")


def _year_tokens(toks):
    """연도 토큰을 4자리로 정규화한 집합 ('24' → '2024', '2026' 그대로)."""
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
    """연도/회차/형 토큰이 양쪽 모두에 있는데 서로 다르면 충돌.

    연도는 2자리('24')와 4자리('2024')를 같은 해로 봅니다(v2.3.1) — '24 2급
    상시' 슬롯이나 '24_2급상시_문제지.pdf'가 2026 세트에 붙지 않도록.
    """
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
    """겹치는 토큰 수 — 연도는 정규화해 같은 해면 1개로 셉니다."""
    non_year = {t for t in a & b if not _is_year_token(t)}
    return len(non_year) + len(_year_tokens(a) & _year_tokens(b))


def _numeric_shared(a, b):
    """겹치는 연도(정규화)·회차·형 토큰 수 — 동점 가르기용."""
    n = len(_year_tokens(a) & _year_tokens(b))
    n += len({t for t in a
              if _ROUND_RE.fullmatch(t) or _FORM_RE.fullmatch(t)} & b)
    return n


def _pdf_score(toks, pt, name=""):
    """세트 토큰과 파일(문제지 PDF·기대값 JSON) 토큰의 일치 점수.

    회차·형·연도(2자리 포함) 충돌이면 0. 연도만 겹치는 파일은 파일명에
    '문제'/'기대값'이 있을 때만 후보(v2.3.1) — '아이모2026_발표.pdf' 같은
    무관한 문서가 연도 하나로 세트에 붙지 않도록.
    """
    if _token_conflict(toks, pt):
        return 0
    shared = _shared_count(toks, pt)
    if shared and not {t for t in toks & pt if not _is_year_token(t)} \
            and not any(w in str(name) for w in ("문제", "기대값")):
        return 0
    return shared


def match_pdf_for_set(toks, pdf_paths, others=()):
    """토큰 우선순위 매칭으로 유일한 문제지 PDF 찾기. 복수/0개면 None.

    others: 다른 세트들의 토큰 집합 목록. PDF가 다른 세트와 같은 점수로
    맞으면(예: '2024 상시 문제지.pdf'가 상시 1회·2회 모두에 맞음) 어느
    세트의 문제지인지 알 수 없으므로 연결하지 않습니다.
    (직접 선택 세트처럼 세트 목록이 없을 때 쓰는 단일 세트용 — 스캔 세트
    전체는 link_files_globally가 전역 유일 연결을 합니다.)
    """
    scored = []
    for p in pdf_paths:
        name = os.path.basename(p)
        pt = set_tokens(name)
        shared = _pdf_score(toks, pt, name)
        if shared < 1:
            continue
        if any(_pdf_score(ot, pt, name) >= shared for ot in others):
            continue           # 다른 세트에도 똑같이 맞는 애매한 PDF
        scored.append((shared, p))
    if not scored:
        return None
    best = max(s for s, _p in scored)
    matched = [p for s, p in scored if s == best]
    return matched[0] if len(matched) == 1 else None


def _file_link_score(st, path):
    """세트 토큰 ↔ 파일 점수 튜플 (겹침 수, 연도·회차·형 겹침 수, 파일 토큰이
    세트 토큰에 다 들어가는지). 후보가 아니면 None.

    겹침이 1개뿐이면 파일 토큰이 세트 토큰에 모두 포함될 때만 후보 —
    '코코모의고사1회_기대값.json'이 '1회' 하나로 2024 기출 1회 세트에 붙지 않도록
    ('2024_문제지.pdf' ↔ '2024 A형'처럼 파일 쪽 토큰이 더 적은 경우는 허용).
    """
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
    """세트 ↔ 파일(문제지 PDF 또는 기대값 JSON) 전역 유일 연결 (v2.3.1).

    각 파일은 최대 한 세트, 각 세트는 최대 한 파일. field가 이미 있는 세트와
    이미 쓰인 파일은 후보에서 빠집니다. 점수(겹치는 토큰 수 → 그중 연도·회차·형
    수 → 파일 토큰이 세트 토큰의 부분집합인지)가 높은 짝부터 배정하고,
    - 어떤 파일에 다른 세트(문제지가 이미 있는 세트 포함)가 더 높은 점수로
      맞으면 그 파일은 그 세트의 것으로 보고 낮은 세트에는 붙이지 않습니다.
    - 같은 점수로 두 후보가 남으면(동점) 연결하지 않고 세트에
      s[field + "_tie"] = "파일명 ↔ 경쟁 세트" 사유를 남깁니다(진단 표시).
    origin이 있으면 연결된 세트에 s[field + "_origin"] = origin.
    cands_of(세트) → 그 세트의 후보 경로 목록(기본: paths 전부) — 같은 이름의
    기대값이 여러 폴더에 있을 때 세트 폴더 사본을 고르는 데 씁니다.
    """
    tokens_of = tokens_of or _set_tokens_of
    used = {os.path.abspath(s[field]) for s in sets if s.get(field)}
    cands = [p for p in paths if os.path.abspath(p) not in used]
    if not cands:
        return
    toks = {id(s): tokens_of(s) for s in sets}
    # 파일별 최고 점수(문제지가 있는 세트 포함) — 더 잘 맞는 세트가 있으면 제외
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
    """세트설정.json 내용을 방어적으로 정규화: 최상위는 dict, 세트 항목과
    '_슬롯매핑/_진행/_자동선택/_설정'은 dict만 남김(깨진 항목은 버림)."""
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
    """설정의 하위 dict (없거나 dict가 아니면 빈 dict)."""
    v = (cfg or {}).get(key)
    return v if isinstance(v, dict) else {}


_FILE_LOCK = threading.RLock()   # 세트설정·기록 파일 읽기/쓰기 (루틴 서버 스레드와 공유)


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
        routine_touch()          # 루틴 페이지가 다시 그리도록 상태 표식 갱신
        return True
    except OSError:
        return False


# --- 상태 변경 표식 (루틴 웹 연동 /api/state 의 generated) ---
_ROUTINE_GEN = {"value": None, "lock": threading.Lock()}


def routine_touch():
    """상태가 바뀔 때마다 호출 — /api/state 의 `generated` 를 반드시 다른 값으로
    갱신합니다(마이크로초 ISO 일시, 같은 마이크로초면 1µs 올림). 세트설정·기록
    저장 함수와 시험 시작/종료에서 호출합니다. 반환: 새 표식."""
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
    """현재 상태 표식 (없으면 지금 만들어 반환)."""
    with _ROUTINE_GEN["lock"]:
        v = _ROUTINE_GEN["value"]
    return v if v is not None else routine_touch()


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
    if s.get("pdf") and s.get("pdf_확인됨"):
        ent["pdf_확인됨"] = True      # 회차가 달라도 사용자가 확인한 연결
    else:
        ent.pop("pdf_확인됨", None)
    return save_set_config(cfg, path)


PDF_MISMATCH_WARNING = ("⚠ 문제지 회차가 세트와 다릅니다 — [문제지 연결]로 "
                        "바로잡으세요")


def pdf_conflicts_with_set(s, pdf=None):
    """연결(저장)된 문제지 PDF의 회차·형·연도 토큰이 세트와 충돌하는지.

    예: '2024년 상시2회' 세트에 '…상시1회 문제지.pdf' → True.
    토큰이 없는 PDF(예: 'aaa.pdf')는 충돌로 보지 않습니다.
    """
    pdf = pdf or s.get("pdf")
    if not pdf:
        return False
    try:
        return _token_conflict(_set_tokens_of(s),
                               set_tokens(os.path.basename(str(pdf))))
    except Exception:
        return False


def _attach_saved_pdf(s, pdf, confirmed, origin):
    """저장된 PDF 연결을 세트에 반영 — 회차 충돌이면 무시하고 로그+경고 표시.
    사용자가 [문제지 연결]에서 확인한 연결(pdf_확인됨)은 그대로 존중."""
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
    """저장된 세트 구성 반영. 사라진 경로·깨진 항목·회차가 다른 PDF는 무시.
    병합된 중복 세트(s["중복"])의 키로 저장된 항목은 대표 세트에 반영합니다."""
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
    """path가 root 폴더(또는 그 하위)에 있는가."""
    try:
        return os.path.commonpath([os.path.abspath(path),
                                   os.path.abspath(root)]) == \
            os.path.abspath(root)
    except ValueError:
        return False


def _dedupe_keys_by_name(paths, set_dir, key_dirs):
    """같은 파일명의 기대값이 여러 폴더에 있으면(세트 폴더 사본 + 루트/기대값
    배포본 등) 우선순위가 높은 하나만 남깁니다 — 동점 후보로 취급되어
    연결이 막히지 않도록. 우선순위: 세트 폴더 → 기대값 폴더 → 그 밖(경로순)."""
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
    """같은 정규화 키의 기대값 후보 중 우선순위: 세트 폴더 → 루트/기대값
    폴더 → 그 밖(경로순). 반환: (경로, 출처 문구) 또는 (None, None)."""
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
    """폴더 깊이(루트 기준 경로 요소 수). 루트가 없거나 밖이면 절대 경로 기준."""
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
    """문제 파일 내용이 같은 세트를 하나로 병합 (v2.3.1).

    사용자가 세트를 복사·개명해 두 폴더에 둔 경우(예: '상시기출2회(문제).xlsm'
    과 '2024년 기출문제 유형 2회(문제).xlsm')를 한 세트로 봅니다. 크기가 같은
    문제 파일끼리만 SHA-256을 계산해 같으면 중복. 대표: 같은 키의 문제지
    PDF·기대값이 붙은 쪽 → 폴더가 얕은 쪽 → 이름순. 나머지는 대표의
    s["중복"] 목록({name, norm, dir, problem, answer, pdf, key})에 남아
    토큰(슬롯·PDF·기대값 매칭)과 진단 표시에 쓰이고, 중복 쪽에만 있던
    문제지·기대값은 대표가 넘겨받습니다. 반환: 병합된 세트 목록.
    """
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
    """내용은 다른데 이름이 비슷한 세트 짝 [(세트A, 세트B, 공통 토큰)] —
    회차/형 토큰을 공유하고 충돌 없이 토큰이 2개 이상 겹치는 경우(진단 안내용)."""
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
    """느슨한 세트 그룹핑 스캔.

    xlsx/xlsm/pdf/기대값 json을 정규화 키로 묶고, 파일명 토큰으로
    문제/정답 역할을 판별합니다. 같은 키에 PDF가 없으면 토큰 퍼지
    매칭으로 문제지 PDF를 연결합니다. '풀이_' 파일과 '채점결과' 폴더
    제외. 반환: [{"name","norm","dir","problem","answer","key","pdf"}]

    기대값 JSON 연결 우선순위(v2.3.0): ① 세트 폴더 안의 정확 키 파일
    ② 루트/기대값/(key_dirs, 자동 업데이트로 배포됨)의 정확 키 파일
    ③ 스캔 루트 어디든 정확 키 파일 ④ 연도·회차·형 토큰이 충돌 없이
    유일하게 맞는 파일(문제지 PDF와 같은 규칙). 출처는 s["key_origin"].
    key_dirs가 스캔 루트 밖이어도 그 폴더의 기대값 JSON은 후보에 넣습니다.

    v2.3.1: 문제 파일 내용이 같은 세트는 하나로 병합(merge_duplicate_sets,
    s["중복"]), 같은 키가 없는 문제지 PDF·기대값 JSON은 link_files_globally로
    전역 유일 연결(각 파일 한 세트, 각 세트 한 파일, 동점은 s["pdf_tie"]/
    s["key_tie"] 사유만 남김).
    """
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
    for kd in key_dirs:      # 스캔 루트 밖의 기대값 폴더도 후보에 포함
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
            # 정답 토큰만 있는 2파일 세트 -> 나머지를 문제로 추정
            problem = next(p for p in g["excel"] if p != answer)
        if not (problem and answer) or problem == answer:
            continue
        set_dir = os.path.dirname(problem)
        key, origin = _pick_exact_key(g["json"], set_dir, key_dirs)
        # 같은 키 PDF가 여럿이면 세트 폴더 안의 것 우선
        pdfs = sorted(g["pdf"], key=lambda p: (os.path.dirname(p) != set_dir, p))
        sets.append({
            "name": display_name(problem), "norm": k,
            "dir": set_dir,
            "problem": problem, "answer": answer,
            "key": key, "key_origin": origin,
            "pdf": pdfs[0] if pdfs else None,
            "pdf_사본": pdfs[1:],        # 같은 키의 다른 사본(진단 표시용)
        })
    # 문제 파일 내용이 같은 세트 병합 (대표가 토큰·문제지·기대값을 넘겨받음)
    sets = merge_duplicate_sets(sets, root)
    # 같은 키 PDF가 없는 세트: 토큰 전역 유일 연결 (이미 연결된 파일과 같은
    # 이름의 다른 사본은 후보에서 제외)
    used_pdf_names = {os.path.basename(s["pdf"]) for s in sets if s.get("pdf")}
    link_files_globally(
        sets, [p for p in all_pdfs if os.path.basename(p) not in used_pdf_names],
        "pdf")
    # 같은 키 기대값 JSON이 없는 세트: 같은 규칙 — 예: '2026_1회_기대값.json'
    # ↔ '2026 … 1회_문제.xlsm'. 같은 이름의 기대값이 여러 폴더에 있으면 세트
    # 폴더 → 기대값 폴더 순으로 하나만 후보에 둡니다.
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
    """직접 선택한 문제/정답(+문제지)으로 세트 구성 (기대값/문제지 자동 감지).

    기대값 JSON: 문제 파일 폴더의 정확 키 → 루트/기대값 폴더(key_dirs)의
    정확 키 → 두 곳의 후보 중 토큰(연도·회차·형)이 2개 이상 겹치고 유일한
    파일 순으로 연결합니다.
    """
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
            elif fn not in have:      # 세트 폴더에 같은 이름이 있으면 그쪽 우선
                key_cands.append(os.path.join(kd, fn))
    if key is None and key_cands:
        # 같은 키가 없으면 토큰(연도·회차·형) 퍼지 매칭 — 유일하고 토큰이
        # 2개 이상(예: 2026·1회) 겹칠 때만 (직접 선택은 다른 세트 정보가 없음)
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
    """점수 값 정규화: 숫자 → 그대로, 숫자 문자열("85", "85.0", "85점") →
    숫자, 그 외(None, "채점 실패", bool …) → None."""
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
    """기록 1건을 현재 스키마로 방어적 정규화. dict가 아니면 None.

    구버전(v1.x) 기록에는 mode/day/루틴 키가 없고(=시험 모드로 취급),
    손으로 고친 파일에는 점수가 문자열이거나 세트명이 빠져 있을 수
    있습니다 — 어느 경우에도 죽지 않고 읽을 수 있는 만큼만 살립니다.
    """
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
    """기록.json 내용(어떤 형태든) → 정규화된 기록 리스트.

    리스트가 정상 형태. dict면 {"기록": [...]} 또는 {세트명: 기록dict}로
    간주해 최대한 살립니다.
    """
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
    """기록.json 로드 (없음/깨짐/옛 스키마여도 절대 예외 없이 리스트 반환)."""
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
    """파싱 불가한 JSON을 덮어쓰기 전에 .corrupt-<일시> 사본으로 보존."""
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
    """기록.json에 응시 기록 1건 추가. 전체 목록 반환."""
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


CONTENT_TYPES_NAME = "[Content_Types].xml"
ZIP_FLAG_DATA_DESCRIPTOR = 0x08


def convert_xlsx_to_xlsm(src, dst):
    """xlsx -> xlsm ZIP 수준 변환.

    [Content_Types].xml의 워크북 메인 파트 content-type만 교체하고 그 외
    모든 파트는 바이트 무손실 복사 (openpyxl 재저장 없음 — 서식·차트 보존).
    항목마다 새 ZipInfo(이름·날짜·DEFLATED)로 깨끗하게 다시 씁니다 —
    원본 ZipInfo를 재사용하면 data descriptor 비트(0x08)·extra 필드 같은
    Office가 싫어하는 흔적이 그대로 전파될 수 있습니다. [Content_Types].xml은
    항상 첫 항목, 디렉터리·중복 항목은 버립니다.
    """
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
    """변환된 사본 검증. (통과 여부, 설명) 반환.

    zipfile.testzip 통과 + [Content_Types].xml이 첫 항목 + 어떤 항목에도
    data descriptor 비트(0x08) 없음 + 매크로 content-type 포함 + (있으면)
    openpyxl 로드.
    """
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

    원본이 .xlsx면 ZIP 수준 변환 후 verify_xlsm_copy 검증(실패 시 원본
    확장자 .xlsx로 그대로 복사해 폴백), .xlsm이면 그대로 복사. 사본의
    MotW도 제거. 결과(경로·크기·변환 성공 여부)를 시작 로그에 남깁니다.
    """
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
    """문제 파일을 풀이_<세트명>_<일시>.xlsm 사본으로 (매크로 저장 가능)."""
    when = when or datetime.now()
    stamp = when.strftime("%Y%m%d_%H%M")
    d = os.path.dirname(os.path.abspath(problem))
    return copy_as_macro_enabled(
        problem, _unique_stem(d, f"풀이_{set_name}_{stamp}"))


# --- 앱 설정 (세트설정.json의 "_설정" 영역) ---

def get_app_setting(name, default=None, path=None):
    cfg = load_set_config(path or SET_CONFIG_PATH)
    return _cfg_section(cfg, "_설정").get(name, default)


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


# --- Excel 직접 실행 · 실행 확인 (Windows) ---

CREATE_NO_WINDOW = 0x08000000
DETACHED_PROCESS = 0x00000008
CREATE_NEW_PROCESS_GROUP = 0x00000200
EXCEL_OPEN_DELAY_MS = 1000      # PDF를 먼저 열고 Excel은 이만큼 뒤에 (맨 앞에 오게)
EXCEL_CHECK_DELAY_MS = 5000     # Excel 실행 후 프로세스 확인까지 대기
_EXCEL_EXE_CACHE = [None, False]   # [경로, 탐색했는지]


def _exe_from_command(cmd):
    """레지스트리 shell\\Open\\command 값 → 실행 파일 경로.
    예: '"C:\\...\\EXCEL.EXE" /dde' 또는 'C:\\...\\EXCEL.EXE "%1"'."""
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
    """Windows 레지스트리에서 EXCEL.EXE 후보 경로 (winreg는 Windows 전용)."""
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
    """흔한 설치 경로 후보 (Program Files[ (x86)]\\Microsoft Office\\…)."""
    out = []
    subs = (r"Microsoft Office\root\Office16", r"Microsoft Office\Office16",
            r"Microsoft Office\root\Office15", r"Microsoft Office\Office15",
            r"Microsoft Office 15\root\Office15", r"Microsoft Office\Office14")
    for var in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432"):
        root = os.environ.get(var)
        if root:
            out.extend(os.path.join(root, sub, "EXCEL.EXE") for sub in subs)
    la = os.environ.get("LOCALAPPDATA")
    if la:   # 스토어판 Office의 앱 실행 별칭
        out.append(os.path.join(la, r"Microsoft\WindowsApps\excel.exe"))
    return out


def find_excel_exe(refresh=False):
    """Windows에서 EXCEL.EXE 경로 탐색: App Paths → ProgID command →
    InstallRoot → 흔한 경로 → PATH. 못 찾으면(또는 Windows 아님) None."""
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
    """풀이 파일 열기. (성공, 방법, 오류, Popen 또는 None) 반환.

    Windows에서 EXCEL.EXE를 찾으면 subprocess.Popen([excel, path])로 직접
    실행(어느 프로그램이 열었는지·즉시 죽었는지 알 수 있음), 못 찾거나 실행
    실패면 os.startfile(기본 프로그램) 폴백. 다른 OS는 open_file.
    """
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
    """EXCEL.EXE 프로세스가 있는지 (Windows: tasklist). 알 수 없으면 None."""
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
    """Excel이 파일을 열면 같은 폴더에 만드는 '~$이름' 잠금 파일 경로.
    (긴 이름은 앞 글자가 잘리므로 뒷부분 일치로 찾음) 없으면 None."""
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
    """Excel이 파일을 열었는지 판정. ('open'|'closed'|'unknown', 설명) 반환.

    잠금 파일(~$…)이 있으면 열림. 없으면 EXCEL.EXE 프로세스 존재 여부와
    직접 실행한 프로세스의 종료 코드로 판단합니다. (이미 떠 있던 Excel에
    파일을 넘기면 새 프로세스는 곧바로 끝나므로 종료 코드 0만으로는
    실패로 보지 않습니다.)
    """
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
    """재시작 명령 (args, Popen kwargs). 래퍼(코코시험장.pyw)로 시작했으면
    래퍼를, 아니면 시험장.py를 같은 인터프리터로 실행합니다.

    Windows: 콘솔이 튀지 않게 pythonw면 DETACHED_PROCESS, python.exe면
    CREATE_NO_WINDOW (+ CREATE_NEW_PROCESS_GROUP, 부모 종료와 분리).
    """
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
# 오늘의 학습 — 14일 루틴 일정 (웹 루틴과 동일: 2026-09-03 시작, 시험 2회)
#   학습일 Day 1~14 = 9/3(목)~9/16(수), 시험 1 = 9/17(목), 시험 2 = 9/18(금).
#   매일 모의고사 1세트 40분 완주 → 채점 → 오답노트 → 오답 재풀이
# ---------------------------------------------------------------------------

ROUTINE_START = date(2026, 9, 3)                     # Day 1 (9/3 목)
EXAM_DATES = [date(2026, 9, 17), date(2026, 9, 18)]  # 시험 1 · 시험 2 (이틀 연속)
EXAM_DATE = EXAM_DATES[0]                            # (구 코드 호환)
ROUTINE_TAG = "2026-09"     # 기록.json 루틴 세대 표시 (구 루틴 기록과 구분)
PROGRESS_NS = "r0903"       # 세트설정.json '_진행' 키 접두 (구 루틴 진행과 분리)

PLAN_EXAM1, PLAN_EXAM2, PLAN_AFTER = 15, 16, 17     # 시험 1 / 시험 2 / 이후
PLAN_ORDER = ([0] + list(range(1, 15))
              + [PLAN_EXAM1, PLAN_EXAM2, PLAN_AFTER])  # 시간순 (시험 2회는 학습일 뒤)
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
    # 9/10(목)~9/16(수): 재응시 사다리 70·70·75·75·80·80·85 (주말 9/12·13 2세트)
    8: {"제목": "재응시", "종류": "모의", "세트": [AUTO], "목표": 70},
    9: {"제목": "재응시", "종류": "모의", "세트": [AUTO], "목표": 70},
    10: {"제목": "재응시", "종류": "모의", "세트": [AUTO, AUTO], "목표": 75},
    11: {"제목": "재응시", "종류": "모의", "세트": [AUTO, AUTO], "목표": 75},
    12: {"제목": "재응시", "종류": "모의", "세트": [AUTO], "목표": 80},
    13: {"제목": "재응시", "종류": "모의", "세트": [AUTO], "목표": 80},
    14: {"제목": "시험 전 최종", "종류": "모의", "세트": [AUTO],
         "목표": 85, "특별": ["실수노트"]},
    PLAN_EXAM1: {"제목": "시험 1", "종류": "안내",
                 "할일": "시험 1 당일(9/17 목). 수험표·신분증 확인, 고사장 "
                        "30분 전 도착. 저장은 Ctrl+S 수시로, 계산작업은 한 "
                        "문제 3분 넘기면 다음으로. 시험 후 저녁에는 세트 "
                        "없이 복기 메모(막힌 유형·시간 부족 구간 3줄)와 실수 "
                        "노트 1회독만 가볍게 — 내일 2차가 있습니다."},
    PLAN_EXAM2: {"제목": "시험 2", "종류": "안내",
                 "할일": "시험 2 당일(9/18 금). 아침에 실수 노트와 어제의 "
                        "복기 메모만 한 번 훑고 출발하세요. 그 항목만 지키면 "
                        "됩니다 — 화이팅!"},
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
    """일정 번호 -> 날짜. Day 1~14 = 9/3~9/16 연속(시험일이 사이에 오면 건너뜀)."""
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
    ④ 오답 재풀이 15분 → (특별: 마감일) 실수 노트 → ⑤ (선택) 함수 퀴즈.
    """
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
        g_txt = f" (목표 {goal}점)" if goal else ""
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
             "설명": "채점이 끝나면 자동 체크되고, 결과는 루틴 페이지"
                    "([루틴 열기])에 자동으로 반영됩니다. 성적 JSON은 "
                    "클립보드에도 복사됩니다(아티팩트 페이지용)."},
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
                             "노트에 정리하세요. 내일 1차·모레 2차 시험장에서 "
                             "볼 마지막 체크리스트입니다."})
    steps.append({"이름": "(선택) 함수 퀴즈", "형": "안내", "분": 10,
                  "선택": True, "웹탭": "quiz",
                  "설명": "(선택) 루틴 페이지 함수 퀴즈 10문제 — 오늘 틀린 "
                         "함수가 있으면 그 함수부터. [웹에서 퀴즈 풀기]로 "
                         "퀴즈 탭이 열리고, 페이지에서 ⑤ 체크하면 이 단계도 "
                         "자동 체크됩니다. 건너뛰어도 됩니다."})
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
        names = list(set_names or plan.get("세트표시") or plan_slot_names(plan))
        if plan.get("적응형"):
            body = f"{plan['제목']}: {' + '.join(names)}" if names \
                else plan["제목"]
            if plan.get("남은세트") is not None and plan.get("종류") == "모의":
                body += f" · 남은 세트 {plan['남은세트']}"
        else:
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


# 일정 슬롯별 식별 규칙 (v2.3.1: 연도+회차/형 중심).
#   "핵심": 반드시 있어야 하는 토큰 묶음 목록 — 묶음 안은 대안(어느 하나면 됨,
#           연도는 2자리/4자리를 같은 해로 봄)
#   "보조": 식별 일치에 더 필요한 묶음 — 없으면 '부분 일치'로 강등
#   "연도제외": True면 연도 토큰(20xx·2자리)이 있는 세트는 후보에서 제외
# 어느 등급이든 슬롯 문구 전체 토큰과 세트 토큰의 연도·회차·형 충돌은 후보 제외
# (예: '2024 상시 1회'는 2026 세트에 붙지 않음).
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
    """슬롯 문구의 식별 규칙 dict (표에 없으면 연도(20xx)를 뺀 토큰 각각이 핵심)."""
    if text in SLOT_IDENTITY:
        return SLOT_IDENTITY[text]
    toks = set_tokens(text)
    rest = sorted(t for t in toks if not _YEAR4_RE.fullmatch(t)) or sorted(toks)
    return {"핵심": [[t] for t in rest]}


def slot_identity_tokens(text):
    """슬롯 문구의 식별 토큰 표시용 집합 — 대안 묶음은 '2024/24'처럼 합쳐 표시."""
    spec = slot_identity_spec(text)
    return {"/".join(g) for g in spec["핵심"]} | \
        {"/".join(g) for g in spec.get("보조", [])}


def _set_tokens_of(s):
    """세트 토큰: 세트명 + 폴더명 (+ 병합된 중복 세트의 이름·폴더명)."""
    toks = set_tokens(s["name"]) | set_tokens(os.path.basename(s["dir"] or ""))
    for d in s.get("중복") or []:
        toks |= set_tokens(d.get("name") or "")
        toks |= set_tokens(os.path.basename(d.get("dir") or ""))
    return toks


def _has_token(st, t):
    """세트 토큰에 t가 있는가 (연도는 2자리/4자리 동일시)."""
    if _is_year_token(t):
        return bool(_year_tokens({t}) & _year_tokens(st))
    return t in st


def _group_hit(st, group):
    """대안 묶음 중 세트에 있는 토큰 (없으면 None)."""
    for t in group:
        if _has_token(st, t):
            return t
    return None


def slot_candidate_score(spec, s, st=None):
    """(슬롯, 세트) 점수 (등급, 연도·회차·형 겹침, 전체 겹침) 또는 None(후보 아님).

    등급 3 전체 토큰 일치(슬롯 문구 토큰이 세트에 다 있음) > 2 식별 토큰 일치
    (핵심+보조) > 1 부분 일치(핵심만). 슬롯 전체 토큰과 연도·회차·형이 충돌하거나
    핵심 묶음이 하나라도 없으면 후보 아님. 반환 튜플 뒤에 사유 문구가 붙습니다:
    ((등급, 숫자겹침, 겹침), 사유).
    """
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
    """알려진 일정 슬롯 문구 전부 (식별 표 + 일정표 + 우선순위 표) — 순서 유지."""
    specs = list(SLOT_IDENTITY)
    for no in sorted(ROUTINE_PLAN):
        for spec in ROUTINE_PLAN[no].get("세트") or []:
            if spec != AUTO and spec not in specs:
                specs.append(spec)
    for spec in list(PRIORITY_SPECS) + [REDO_SPEC]:
        if spec not in specs:
            specs.append(spec)
    return specs


def assign_slots(sets, specs=None, mapping=None):
    """일정 슬롯 ↔ 세트 전역 유일 배정 (v2.3.1). 반환 {슬롯: (세트|None, 사유)}.

    요청한 specs에 알려진 슬롯 전부(all_slot_specs)를 더해 함께 배정합니다 —
    '2026 1회'처럼 확실한 슬롯이 그 세트를 먼저 가져가야 '2024 상시 1회'나
    '컴활 2급 상시'가 같은 세트를 차지하지 못하기 때문입니다.
    ① 세트설정 `_슬롯매핑`(mapping, 직접 선택 저장)이 최우선 — 그 세트(같은 키의
       스캔 세트 포함)는 다른 슬롯 후보에서 제외.
    ② 모든 (슬롯, 세트) 쌍을 slot_candidate_score로 채점, 점수 높은 짝부터
       탐욕적으로 배정(한 세트는 한 슬롯). 같은 점수의 짝이 슬롯 또는 세트를
       공유하면(동점) 그 슬롯들은 미배정으로 두고 사유에 후보를 나열합니다.
    ③ 남은 슬롯: 후보가 다른 슬롯에 배정됐으면 그 사실을, 아예 없으면
       '후보 없음'을 사유로.
    """
    specs = list(dict.fromkeys(list(specs or []) + all_slot_specs()))
    result = {}
    taken = set()          # 배정된 세트 norm
    mapping = mapping or {}
    for spec in specs:                           # ① 직접 선택 저장
        ent = mapping.get(spec)
        if not isinstance(ent, dict):
            continue
        ms = set_from_mapping(ent)
        if ms is None:
            continue
        result[spec] = (ms, "직접 선택 저장됨 (" +
                        os.path.basename(ms["problem"]) + ")")
        taken.add(ms["norm"])
        for s in sets:                           # 중복 병합 대표도 같은 세트
            if any(d.get("norm") == ms["norm"] for d in s.get("중복") or []):
                taken.add(s["norm"])
    toks = {id(s): _set_tokens_of(s) for s in sets}
    pairs = []                                   # ② 채점
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
    tied = {}                                    # spec -> 사유
    blocked = set()                              # 동점으로 묶인 세트 norm
    winner = {}                                  # norm -> spec
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
    for spec in specs:                           # ③ 미배정 사유
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
    """슬롯 문구 -> (세트 or None, 판정 설명). GUI 없이 테스트 가능.

    v2.3.1: 알려진 슬롯 전부와 함께 전역 유일 배정(assign_slots)한 결과 중
    text의 항목 — 다른 슬롯이 더 잘 맞는 세트는 이 슬롯에 오지 않습니다.
    """
    return assign_slots(sets, [text], mapping)[text]


def find_set_for_tokens(sets, text):
    """일정의 세트 지정 문구를 스캔된 세트에 매칭 (match_slot의 세트만)."""
    return match_slot(sets, text)[0]


def load_slot_mapping(path=None):
    """세트설정.json '_슬롯매핑': {슬롯 문구: {problem, answer, pdf}}.
    (저장된 pdf의 회차 충돌 검증은 set_from_mapping에서)"""
    cfg = load_set_config(path or SET_CONFIG_PATH)
    raw = _cfg_section(cfg, "_슬롯매핑")
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
        if not (isinstance(prob, str) and isinstance(ans, str) and prob
                and ans and os.path.isfile(prob) and os.path.isfile(ans)):
            return None
        s = build_direct_set(prob, ans)
        if ent.get("name"):
            s["name"] = str(ent["name"])
        s["direct"] = True
        pdf = ent.get("pdf")
        if isinstance(pdf, str) and pdf and os.path.isfile(pdf):
            # 저장된 PDF가 회차·형·연도 충돌이면 무시 (같은 키 자동 감지 유지)
            _attach_saved_pdf(s, pdf, bool(ent.get("pdf_확인됨")), "슬롯매핑")
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
    tie_files = {}          # 동점으로 연결 못 한 파일명 -> 사유
    for s in sets:
        for k in ("problem", "answer", "pdf", "key"):
            if s.get(k):
                owner[os.path.abspath(s[k])] = s["name"]
        for d in s.get("중복") or []:          # 병합된 중복 세트의 파일
            for k in ("problem", "answer", "pdf", "key"):
                if d.get(k):
                    owner.setdefault(os.path.abspath(d[k]),
                                     f"{s['name']} (중복·동일 내용)")
        for p in s.get("pdf_사본") or []:      # 같은 키 PDF의 다른 사본
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
    # 세트 토큰이 하나도 없는 PDF(대학 서류 등)는 연결 후보가 아니므로 접어 표시
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
        specs = []
        for no in range(1, 15):
            for spec in ROUTINE_PLAN.get(no, {}).get("세트") or []:
                if spec != AUTO and spec not in specs:
                    specs.append(spec)
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

    고정 슬롯은 저장된 직접 선택(mapping) → 전역 유일 배정(assign_slots) 순,
    자동 슬롯은 저장된 선택(saved)을 우선 복원하고 없으면
    pick_set_for_retry로 고릅니다. 슬롯끼리는 서로 다른 세트.
    """
    slots = list(plan.get("세트") or [])
    out = [None] * len(slots)
    reasons = [""] * len(slots)
    taken = set()
    mapping = mapping or {}
    objs = list(plan.get("세트객체") or [])
    whys = list(plan.get("슬롯사유") or [])
    assigned = None
    for k, spec in enumerate(slots):          # ① 고정 세트
        if spec != AUTO:
            s, how = None, ""
            if k < len(objs) and objs[k]:            # 적응형 배정 세트
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


def retry_payload_for_set(s, minutes=15, json_path=None):
    """세트의 최신 전체 채점 JSON(또는 json_path 로 지정한 채점 JSON)에서
    오답 시트를 뽑아 재풀이 실행 정보로.

    반환 ('retry', payload) / ('missing', {이유}) / ('info', {메시지, 자동완료}).
    """
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


# ---------------------------------------------------------------------------
# 적응형 일정 엔진 v5 — 기록(기록.json) 기준으로 매번 재계산 (웹 루틴과 동일 규격)
#   전날·그전날 못 한 세트는 자동으로 풀에 남아 오늘 이후로 재배치되고,
#   시험 전날(구간 마감)까지 전 세트 완주를 보장하도록 용량을 올린다.
#   구간은 시험일 목록에서 도출: 이전 시험일 다음 날 ~ 다음 시험일 전날,
#   학습일이 하나도 없는 구간(9/18 앞)은 건너뛰므로 이 일정은 단일 구간(9/3~9/16).
# ---------------------------------------------------------------------------

PRIORITY_SPECS = ["2024 상시 1회", "코코 1회", "2024 A형", "2024 상시 2회",
                  "코코 2회", "2024 B형", "24 2급 상시", "컴활 2급 상시",
                  "2026 1회"]          # 첫응시 고정 우선순위 (그 외 신규는 이름순)
REDO_SPEC = "2026 1회"                # 필수 재도전 (루틴 중 미응시면)
REDO_GOAL = 65
RETRY_MIN = 2                         # 재응시 최소(재도전 포함) — 부족하면 평일 승격
RETRY_MIN_SEG1 = RETRY_MIN            # (구 이름 호환)
RETRY_GOALS = [70, 75, 80, 85]        # 재응시 목표 사다리 (배정된 날짜 순으로 분배)
WEEKEND_CAP, WEEKDAY_CAP = 2, 1


def study_segments():
    """시험일 목록에서 학습 구간 도출: [{seg, seg_start, end, exam}] (학습일 0인
    구간은 건너뜀). 9/17·9/18 시험이면 1구간 9/3~9/16 하나뿐."""
    segs = []
    prev = ROUTINE_START
    for ex in sorted(EXAM_DATES):
        start, end = prev, ex - timedelta(days=1)
        if end >= start:
            segs.append({"seg": len(segs) + 1, "seg_start": start, "end": end,
                         "exam": ex})
        prev = ex + timedelta(days=1)
    return segs


def next_segment_after(seg_no):
    """seg_no 다음 학습 구간 (없으면 None) — 미리보기(preview)용."""
    return next((sg for sg in study_segments() if sg["seg"] == seg_no + 1), None)


def segment_for(today):
    """오늘 -> 구간 정보. kind: study(학습 구간) / exam / after."""
    if today in EXAM_DATES:
        return {"kind": "exam", "exam_no": EXAM_DATES.index(today) + 1}
    if today > max(EXAM_DATES):
        return {"kind": "after"}
    for sg in study_segments():
        if today <= sg["end"]:
            return {"kind": "study", "seg": sg["seg"], "seg_start": sg["seg_start"],
                    "start": max(today, sg["seg_start"]), "end": sg["end"]}
    return {"kind": "after"}      # 시험일 사이에 학습일이 없는 날 (이 일정엔 없음)


def base_capacity(d, seg_end):
    """기본 용량: 주말 2 / 평일 1 / 마감일(시험 전날) 1(+실수노트)."""
    if d == seg_end:
        return 1
    return WEEKEND_CAP if d.weekday() >= 5 else WEEKDAY_CAP


def exam_records_by_set(sets, records):
    """{norm: [시험 모드 기록(날짜 오름차순)]} — 부분연습·오답재풀이 제외."""
    out = {}
    for s in sets:
        recs = [r for r in set_exam_records(s["name"], records)
                if _record_date(r)]
        recs.sort(key=lambda r: (_record_date(r), str(r.get("일시"))))
        out[s["norm"]] = recs
    return out


def prioritized_sets(sets):
    """세트를 첫응시 우선순위로 정렬: PRIORITY_SPECS 순 → 그 외 이름순.
    반환 [(세트, 우선순위 라벨)]."""
    ranked = {}
    for i, spec in enumerate(PRIORITY_SPECS):
        s, _how = match_slot(sets, spec)
        if s and s["norm"] not in ranked:
            ranked[s["norm"]] = (i, spec)
    rest = sorted((s for s in sets if s["norm"] not in ranked),
                  key=lambda s: s["name"])
    out = [(s, ranked[s["norm"]][1]) for s in
           sorted((s for s in sets if s["norm"] in ranked),
                  key=lambda s: ranked[s["norm"]][0])]
    out.extend((s, "신규") for s in rest)
    return out


def _best(recs):
    scores = [r["점수"] for r in recs
              if isinstance(r.get("점수"), (int, float))]
    return max(scores) if scores else None


def _slot(kind, s, goal=None, why="", spec=None):
    return {"kind": kind, "set": s, "name": s["name"] if s else "자동 선택",
            "goal": goal, "auto": s is None, "why": why,
            "spec": spec or (s["name"] if s else AUTO)}


def build_adaptive_plan(today, records, sets):
    """적응형 일정 계산 (결정적: 같은 입력이면 같은 출력).

    반환 dict: kind, seg, start, end, days{날짜: [슬롯...]}, boost_days,
    demoted, missed_days, reason, remaining, warning, retry_target.
    슬롯: {kind: first|redo|retry, set, name, goal, auto, why, spec}
    """
    seg = segment_for(today)
    plan = {"today": today, "kind": seg["kind"], "days": {}, "boost_days": [],
            "demoted": [], "missed_days": [], "reason": "", "remaining": 0,
            "warning": None, "retry_target": 0, "seg": seg.get("seg")}
    if seg["kind"] != "study":
        plan.update(seg)
        return plan
    start, end, seg_start = seg["start"], seg["end"], seg["seg_start"]
    plan.update({"start": start, "end": end, "seg_start": seg_start})
    by_set = exam_records_by_set(sets, records)
    recs_all = [r for r in records
                if r.get("mode") not in ("부분연습", "오답재풀이")
                and _record_date(r)]
    done_today = len([r for r in recs_all if _record_date(r) == today])
    by_date = {}
    for r in recs_all:
        by_date.setdefault(_record_date(r), []).append(r)
    plan["_recs_by_date"] = by_date

    # 학습일·기본 용량 (승격 후보 = 평일, 이른 날부터 → 마감일은 최후)
    days = []
    d = start
    while d <= end:
        if d not in EXAM_DATES:
            days.append(d)
        d += timedelta(days=1)
    cap = {dd: base_capacity(dd, end) for dd in days}
    if today in cap:
        cap[today] = max(0, cap[today] - done_today)
    boostable = [dd for dd in days
                 if dd.weekday() < 5 and not (dd == today and done_today >= 2)]

    # 풀
    ordered = prioritized_sets(sets)
    first_pool = [(s, lab) for s, lab in ordered if not by_set.get(s["norm"])]
    redo_set = None
    rs, _how = match_slot(sets, REDO_SPEC)
    if rs and by_set.get(rs["norm"]) and not any(
            _record_date(r) >= ROUTINE_START for r in by_set[rs["norm"]]):
        redo_set = rs                           # 루틴 전 기록만 있음 → 재도전 필수
    mandatory = []
    for s, lab in first_pool:
        goal = REDO_GOAL if lab == REDO_SPEC else None
        sl = _slot("first", s, goal, f"첫 응시 · 우선순위 {lab}", spec=s["name"])
        sl["counts_as_retry"] = (lab == REDO_SPEC)   # 2026 1회는 재도전 몫
        mandatory.append(sl)
        if redo_set is not None and lab == REDO_SPEC:
            redo_set = None
    if redo_set is not None:
        pos = len(mandatory)
        for i, sl in enumerate(mandatory):
            if sl["why"].endswith("신규"):
                pos = i
                break
        redo_slot = _slot("redo", redo_set, REDO_GOAL,
                          "2026 1회 재도전 (루틴 중 미응시)",
                          spec=redo_set["name"])
        redo_slot["counts_as_retry"] = True
        mandatory.insert(pos, redo_slot)
    attempted = [(s, lab) for s, lab in ordered if by_set.get(s["norm"])]

    def recent(s):
        last = max(_record_date(r) for r in by_set[s["norm"]])
        return (today - last).days < 2

    def retry_key(s):
        b = _best(by_set[s["norm"]])
        return (b if b is not None else -1, s["name"])

    def retried_in_routine(s):
        """루틴 중 재응시(첫 기록 이후의 기록이 구간 안) 여부."""
        return any(_record_date(r) >= seg_start and i > 0
                   for i, r in enumerate(by_set[s["norm"]]))

    # 재응시 풀: 아직 재응시 안 한 세트 먼저, 각 그룹은 최고점 오름차순(동점 이름순),
    # 오늘·어제 응시한 세트는 그룹 뒤로
    groups = ([s for s, _l in attempted if not retried_in_routine(s)],
              [s for s, _l in attempted if retried_in_routine(s)])
    retry_pool = []
    for g in groups:
        retry_pool.extend(sorted((s for s in g if not recent(s)), key=retry_key))
        retry_pool.extend(sorted((s for s in g if recent(s)), key=retry_key))
    retries_done = 0
    redo_norm = rs["norm"] if rs else None
    for s, _l in attempted:
        recs = by_set[s["norm"]]
        if s["norm"] == redo_norm:              # 2026 1회: 루틴 중 응시 = 재도전 완료
            if any(_record_date(r) >= seg_start for r in recs):
                retries_done += 1
            continue
        for i, r in enumerate(recs):
            if _record_date(r) >= seg_start and i > 0:
                retries_done += 1

    # 필요량(첫응시 전부 + 재도전 + 재응시 최소) vs 용량 — 단계적 확장/축소
    #   ① 평일 승격(이른 날부터, 마감일 최후) ② 재응시 최소 축소(최소 1)
    #   ③ 첫응시 뒤에서부터 '선택' 강등
    retry_target = RETRY_MIN
    demoted = []
    boost_days = []
    while True:
        redo_in = any(sl.get("counts_as_retry") for sl in mandatory)
        extra = max(0, retry_target - retries_done - (1 if redo_in else 0))
        required = len(mandatory) + extra
        total_cap = sum(cap.values())
        if required <= total_cap:
            break
        nxt = next((dd for dd in boostable if dd not in boost_days), None)
        if nxt is not None:                       # ① 평일 용량 2로
            boost_days.append(nxt)
            cap[nxt] += 1
            continue
        if retry_target > 1:                      # ② 재응시 축소 (최소 1)
            retry_target -= 1
            continue
        if mandatory:                             # ③ 첫응시 뒤에서부터 강등
            sl = mandatory.pop()
            demoted.append(sl)
            continue
        break
    plan["retry_target"] = retry_target

    # 배정 1: 첫응시·재도전(우선순위 순)을 가장 이른 슬롯부터, 하루 안 중복 없음
    queue = list(mandatory)
    per_day = {dd: [] for dd in days}
    for dd in days:
        while len(per_day[dd]) < cap[dd] and queue:
            pick = None
            for i, sl in enumerate(queue):      # 같은 날 같은 세트 금지
                if all(x["set"]["norm"] != sl["set"]["norm"] for x in per_day[dd]):
                    pick = queue.pop(i)
                    break
            if pick is None:
                break
            per_day[dd].append(dict(pick, date=dd))
    for sl in queue:                            # 배정 못 한 필수 슬롯은 강등 처리
        demoted.append(sl)
    for sl in demoted:
        sl["kind"] = "optional"

    # 배정 2: 남는 용량은 전부 재응시 — 풀을 순환(최고점 낮은 세트부터), 같은 날
    # 이미 있는 세트는 건너뛰고 풀이 비었거나 전부 겹치면 '자동 선택'(당일 기록 기준)
    def _retry_slot(s):
        why = ("재응시 · 최고점 " + (f"{_best(by_set[s['norm']]):g}점"
                                  if _best(by_set[s["norm"]]) is not None
                                  else "점수 없음")) if s else "재응시 · 자동 선택"
        return _slot("retry", s, None, why)

    cursor = 0
    for dd in days:
        while len(per_day[dd]) < cap[dd]:
            s = None
            for _try in range(len(retry_pool)):
                cand = retry_pool[cursor % len(retry_pool)]
                cursor += 1
                if all(x["set"] is None or x["set"]["norm"] != cand["norm"]
                       for x in per_day[dd]):
                    s = cand
                    break
            per_day[dd].append(dict(_retry_slot(s), date=dd))
    # 재응시 목표 사다리 70→75→80→85: 재응시가 배정된 날짜 순으로 4단계 분배
    # (같은 날 2세트는 같은 목표). 예: 7일이면 70·70·75·75·80·80·85.
    retry_days = [dd for dd in days
                  if any(sl["kind"] == "retry" for sl in per_day[dd])]
    for rank, dd in enumerate(retry_days):
        level = min(len(RETRY_GOALS) - 1,
                    (rank * len(RETRY_GOALS)) // max(1, len(retry_days)))
        for sl in per_day[dd]:
            if sl["kind"] == "retry":
                sl["goal"] = RETRY_GOALS[level]
    plan["days"] = {dd: per_day[dd] for dd in days}
    plan["boost_days"] = boost_days
    plan["demoted"] = [sl["name"] for sl in demoted]
    plan["remaining"] = sum(len(v) for v in per_day.values())

    # 재배치 사유
    missed = []
    dd = seg_start
    while dd < min(today, end + timedelta(days=1)):
        if dd not in EXAM_DATES and not any(
                _record_date(r) == dd for r in recs_all):
            missed.append(dd)
        dd += timedelta(days=1)
    plan["missed_days"] = missed
    n_days = len([dd for dd in days if cap[dd] > 0])
    reason = ""
    if missed or boost_days or demoted:
        if missed:
            reason = "·".join(f"{m.month}/{m.day}" for m in missed) + " 미완주 → "
        reason += f"남은 세트 {plan['remaining']}개를 {n_days}일에 재배치"
        if boost_days:
            reason += "(평일도 2세트)"
    plan["reason"] = reason
    if demoted:
        plan["warning"] = (f"용량 부족: {len(demoted)}세트는 '선택'으로 강등 — "
                           + ", ".join(plan["demoted"])
                           + " (남는 시간에 추가 응시하세요)")
    return plan


def adaptive_day_plan(adaptive, no):
    """적응형 결과 -> 일정 번호 no의 카드/단계 창용 plan dict.

    adaptive가 없거나 해당 날짜가 계산 범위 밖이면 고정 일정표로 폴백.
    """
    base = plan_for_day(no)
    if not adaptive or adaptive.get("kind") != "study":
        return base
    d = base.get("날짜")
    if d is None or not (1 <= base["no"] <= 14):
        return base
    today = adaptive["today"]
    plan = dict(base)
    plan["적응형"] = True
    plan["특별"] = []
    if d < adaptive["start"]:                          # 지난 날: 기록 요약
        recs = adaptive.get("_recs_by_date", {}).get(d, [])
        plan.update({"종류": "지난", "세트": [], "세트객체": [], "스텝": [],
                     "목표들": []})
        if recs:
            plan["제목"] = "완료"
            plan["세트표시"] = [f"{r.get('세트명', '?')}"
                            + (f"({r['점수']}점)" if r.get("점수") is not None
                               else "") for r in recs]
            plan["할일"] = "완주한 날입니다."
        else:
            plan["제목"] = "미완주"
            plan["세트표시"] = []
            plan["할일"] = "이 날 완주 기록이 없어 남은 세트를 오늘 이후로 재배치했습니다."
        return plan
    if d not in adaptive["days"]:                      # 다음 구간 등 범위 밖
        plan.update({"종류": "안내", "세트": [], "세트객체": [], "스텝": [],
                     "제목": "다음 구간", "세트표시": [],
                     "할일": "시험 후 기록을 반영해 다시 계산됩니다."})
        return plan
    slots = adaptive["days"][d]
    plan["세트"] = [sl["spec"] if sl["set"] else AUTO for sl in slots]
    plan["세트객체"] = [sl["set"] for sl in slots]
    plan["슬롯사유"] = [sl["why"] for sl in slots]
    plan["목표들"] = [sl["goal"] for sl in slots]
    plan["목표"] = next((g for g in plan["목표들"] if g), None)
    plan["종류"] = "모의"
    plan["제목"] = "오늘" if d == today else "예정"
    plan["남은세트"] = adaptive["remaining"]
    if d == adaptive["end"]:
        plan["특별"].append("실수노트")
    if d == today:
        plan["사유"] = adaptive.get("reason") or ""
        plan["경고"] = adaptive.get("warning")
    names = [sl["name"] for sl in slots]
    if not slots:
        plan["제목"] = "오늘 완료" if d == today else "예정 없음"
        plan["할일"] = ("오늘 몫을 완주했습니다 — 오답노트 모드로 복습하거나 "
                      "쉬세요." if d == today else
                      "이 날은 배정된 세트가 없습니다 (여유일).")
        plan["스텝"] = []
        return plan
    goal_txt = " ".join(f"[{n} 목표 {g}점]" for n, g in zip(names, plan["목표들"])
                        if g)
    plan["할일"] = ("모의고사 " + " + ".join(names)
                  + " 40분 완주 → 채점 → 오답노트 → 오답 재풀이 15분"
                  + (f" (오늘은 {len(names)}세트)" if len(names) > 1 else "")
                  + (" " + goal_txt if goal_txt else ""))
    plan["스텝"] = build_day_steps(plan, names)
    return plan


STEP_DONE_MESSAGE = ("오늘 완료! 채점 결과는 루틴 페이지([루틴 열기])에 "
                     "자동 반영되어 있습니다.")


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
    payload = {"메시지": step.get("설명") or step.get("이름") or ""}
    if step.get("웹탭"):
        payload["웹탭"] = str(step["웹탭"])     # 루틴 페이지의 탭(#quiz 등)
    return "info", payload


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
                        path=SET_CONFIG_PATH, count_up=False, total=None,
                        done=None):
    """오답연습 체크 상태(+횟수)를 세트설정.json에 저장.
    점수 기록(기록.json)에는 아무것도 남기지 않습니다.

    total(오답 항목 수)을 주면 전부 체크됐을 때 '완료'[json_name]에 완료 일시를
    남기고, 하나라도 풀리면 지웁니다(루틴 웹 ③ 오답노트 판정과 같은 저장소).
    done=True/False 는 웹 토글처럼 완료 표식을 직접 켜고 끕니다."""
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
    """오답노트 완료 여부 → (done, when). 완료 표식이 있거나(웹 토글·전체 체크)
    total 개 항목이 모두 체크되어 있으면 완료."""
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


# ---------------------------------------------------------------------------
# 자동 업데이트 (공개 저장소) — v2.3.0 완전 자동
# ---------------------------------------------------------------------------
#
# 흐름: 실행할 때마다(하루 1회 제한 없음, 같은 실행에서는 1회) 백그라운드
# 스레드가 version.json(타임아웃 3초, 오프라인이면 조용히 건너뜀)을 확인하고
#   - 새 버전이면 files(프로그램)+data_files(기대값 JSON)를 전부 .new로
#     내려받아 검증(sha256 맵이 있으면 해시, .py/.pyw는 __version__ 표식 +
#     py_compile, .json은 json.loads, 시험장.py의 __version__ == 새 버전)한 뒤
#     메인 스레드에서 .bak 백업 → 교체(하나라도 실패하면 전체 롤백) →
#     자동 재시작(새 프로세스를 먼저 띄우고 종료). 시험 진행 중이면 적용을
#     미루고 시험 종료 후 적용합니다.
#   - 버전이 같아도 data_files/set_files 중 로컬에 없거나 sha256이 다른
#     파일은 내려받아 갱신(재시작 없음). 원격에서 사라진 파일은 지우지 않습니다.
# 재시작된 새 인스턴스는 .update_applied.json을 읽어 "자동 업데이트됨" 띠를
# 띄웁니다. 세트설정.json `_설정.자동업데이트`(기본 true)로 끌 수 있고,
# [업데이트 확인] 버튼은 토글과 무관하게 즉시 확인·적용합니다.
#
# version.json 형식:
#   {"version": "2.3.1", "notes": "...",
#    "files":      {"시험장/시험장.py": "시험장.py", ...},      # 프로그램
#    "data_files": {"기대값/코코모의고사1회_기대값.json": "기대값/..."},  # 기대값 JSON
#    "set_files":  {"모의고사/코코모의고사1회_문제.xlsx": "모의고사/...",  # 세트(xlsx·pdf)
#                   "시험장/루틴.html": "루틴.html"},                  # 루틴 페이지(2.4.0)
#    "sha256":     {"시험장.py": "<hex>", "기대값/...json": "<hex>", ...}}
# 루틴 페이지(HTML)도 set_files에 둡니다 — 2.3.0의 data_files 검증은 json.loads를
# 요구해 HTML이 섞이면 2.3.0 사용자의 업데이트가 실패하지만, set_files는 2.3.0
# 이하가 무시하고 2.3.1+는 sha256(+UTF-8 텍스트)만 검사해 <루트>/시험장/루틴.html
# 에 내려받습니다(서버가 매 요청마다 파일을 읽어 재시작 없이 반영).
# 2.2.3 이하의 apply_update는 files 항목에 __version__ 표식과 py_compile을
# 요구하므로 JSON은 반드시 data_files에만 둡니다(구버전은 그 키를 무시).
# 2.3.0의 data_files 검증은 UTF-8 디코드 + json.loads를 요구해 xlsx·pdf가
# 섞이면 업데이트 전체가 실패하므로, 바이너리 세트 파일은 set_files(2.3.1+에서
# 처리, 2.3.0 이하는 무시)에만 둡니다. set_files는 sha256(필수)만 검사하고
# 설치 루트(기대값/의 상위)/모의고사/… 에 내려받습니다.

UPDATE_BASE_URL = ("https://raw.githubusercontent.com/"
                   "xasmine02/cocomate-study-tools/main/")
SET_FILES_KEY = "set_files"      # version.json: 세트 파일(xlsx·pdf) — 2.3.1+
_BINARY_MAGIC = {".xlsx": b"PK\x03\x04", ".xlsm": b"PK\x03\x04", ".pdf": b"%PDF"}
UPDATE_NOTICE_PATH = os.path.join(RECORDS_DIR, ".update_applied.json")
AUTO_UPDATE_SETTING = "자동업데이트"
_VERSION_MARK_RE = re.compile(r'^__version__\s*=\s*["\']([^"\']+)["\']',
                              re.MULTILINE)


def _version_tuple(v):
    nums = re.findall(r"\d+", str(v or ""))
    return tuple(int(x) for x in nums[:3]) if nums else (0,)


def auto_update_enabled(path=None):
    """세트설정.json `_설정.자동업데이트` (기본 true)."""
    return bool(get_app_setting(AUTO_UPDATE_SETTING, True, path=path))


def _http_get(url, timeout):
    import urllib.request
    req = urllib.request.Request(url, headers={"User-Agent": "cocomate-updater"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def fetch_update_info(base_url=UPDATE_BASE_URL, timeout=3):
    """version.json 조회. 실패/404/오프라인이면 None (조용히 스킵)."""
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
    """소스 텍스트의 `__version__ = "x.y.z"` 값 (없으면 None)."""
    m = _VERSION_MARK_RE.search(text or "")
    return m.group(1) if m else None


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


def _data_target_path(rel_key, base_dir=None):
    """데이터 상대 경로 -> 설치 위치 (폴더는 적용 시 생성).

    '기대값/x_기대값.json' → 루트/기대값/x_기대값.json (expected_values_dir),
    '모의고사/코코모의고사1회_문제.xlsx' → 루트/모의고사/… — 루트는 설치 구조에
    따른 기대값 폴더의 상위(시험장/ 구조면 <루트>, 평면 구조면 그 폴더).
    폴더 없는 키('x.json')는 기대값/으로.
    """
    parts = [p for p in str(rel_key).split("/") if p not in ("", ".", "..")]
    if not parts:
        raise RuntimeError(f"잘못된 데이터 경로: {rel_key!r}")
    if len(parts) == 1:
        parts = [EXPECTED_DIR_NAME] + parts
    root = os.path.dirname(expected_values_dir(base_dir))
    return os.path.join(root, *parts)


def _verify_download(repo_rel, data, kind, info, expect_version=None):
    """내려받은 내용 검증 (실패 시 RuntimeError). 반환: 디코드된 텍스트
    (바이너리 세트 파일은 None).

    .py/.pyw: sha256(있으면) + __version__ 표식 + (호출 쪽에서) py_compile.
    .json: sha256(있으면) + json.loads. 그 밖의 확장자(xlsx·xlsm·pdf, v2.3.1
    set_files): sha256 필수 + 파일 머리(PK/%PDF)만 확인 — 텍스트 검사 없음.
    """
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
        # 루틴 페이지(v2.4.0, set_files): sha256 필수 + UTF-8 텍스트면 통과
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
    """items: [(rel_key, repo_rel, target, kind)] 다운로드 → 검증 → target.new.
    실패하면 만든 .new를 모두 지우고 예외."""
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
                           "rel": repo_rel})   # 실패 시 정리 대상 등록 후 검증
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
    """새 버전의 files(프로그램)+data_files(기대값)+set_files(세트 파일) 전부
    다운로드·검증·스테이징.

    반환: (성공 여부, 메시지, staged). 실패하면 .new를 정리하고 기존 파일은
    건드리지 않습니다. 적용은 commit_staged().
    """
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
    """스테이징된 .new를 .bak 백업 후 교체. 하나라도 실패하면 전체 롤백.
    반환: (성공 여부, 메시지)."""
    replaced = []     # (target, 기존 파일 있었는지)
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
    """모든 파일 다운로드 -> 검증 -> 원자 교체(.bak 1개). (구버전 호환 API)

    반환: (성공 여부, 메시지). 어느 단계든 실패하면 기존 파일로 롤백.
    """
    ok, msg, staged = stage_update(info, base_url, base_dir, timeout)
    if not ok:
        return False, msg
    return commit_staged(staged, (info or {}).get("version"))


def _all_data_files(info):
    """data_files(기대값 JSON) + set_files(세트 xlsx·pdf) 합친 {설치 키: 저장소 경로}."""
    out = {}
    for k in ("data_files", SET_FILES_KEY):
        v = (info or {}).get(k)
        if isinstance(v, dict):
            out.update(v)
    return out


def data_files_to_sync(info, base_dir=None):
    """로컬에 없거나 sha256이 다른 데이터 항목 [(rel_key, repo_rel, target)]
    (기대값 JSON + 세트 파일). sha256 맵에 없는 항목은 로컬 파일이 없을 때만
    내려받습니다."""
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
    """버전이 같아도 기대값(data_files)·세트 파일(set_files)만 갱신.
    반환: (성공, 개수, 메시지)."""
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
    """적용 직후 기록 — 재시작된 새 인스턴스가 읽어 안내 띠를 띄웁니다."""
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
    """아직 안 보여 준 업데이트 안내가 있고 현재 버전이 그 버전이면 반환
    (+ shown 표시). 없으면 None."""
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
    """자동 업데이트 상태 기계 (GUI 없이 테스트 가능).

    check()는 워커 스레드에서: 토글·중복 확인 → version.json → 새 버전이면
    stage_update(다운로드·검증)까지 하고 pending에 보관, 같은 버전이면
    기대값(data_files)만 동기화. apply_pending()은 메인 스레드에서:
    시험 중이면 미루고('deferred'), 아니면 commit_staged + 안내 기록.
    """

    def __init__(self, base_url=UPDATE_BASE_URL, base_dir=None, enabled=None,
                 current=None, log=None, notice_path=None):
        self.base_url = base_url
        self.base_dir = base_dir or BASE_DIR
        self._enabled = enabled if enabled is not None else auto_update_enabled
        self.current = current or __version__
        self.log = log or startup_log
        self.notice_path = notice_path or UPDATE_NOTICE_PATH
        self.checked = False        # 같은 실행에서 중복 확인 방지
        self.checking = False       # 워커 실행 중
        self.info = None
        self.pending = None         # 검증 완료·적용 대기 staged 목록
        self.state = None           # 마지막 check 결과
        self.detail = ""            # 마지막 메시지(오류 상세 등)
        self.data_synced = 0        # 마지막 check에서 갱신한 기대값 수
        self.applied_version = None

    def enabled(self):
        try:
            return bool(self._enabled())
        except Exception:
            return True

    def check(self, force=False, timeout=3, dl_timeout=15):
        """확인 + (새 버전이면) 다운로드·검증. 상태 문자열 반환:
        disabled / already / offline / latest / staged / data / failed / stuck
        force: 수동 [업데이트 확인] — 토글과 중복 확인 제한 무시."""
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
        """스테이징된 업데이트 적용. 반환: applied / deferred / none / failed."""
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


# ---------------------------------------------------------------------------
# 루틴 웹 연동 서버 (v2.4.0) — 규약: 문서/연동_API.md v2.4.1
# ---------------------------------------------------------------------------
#
# 시험장이 시작될 때 127.0.0.1:<port>(기본 8765, 사용 중이면 빈 포트)에
# ThreadingHTTPServer 를 데몬 스레드로 띄우고 시험장 폴더의 루틴.html 을
# 서빙합니다. 프로그램이 상태의 유일한 원본이며, 루틴 페이지는 /api/state 를
# 5초마다 읽어 그립니다(`generated` 가 바뀔 때만 다시 그림).
#   - 서버 스레드는 파일(기록.json·세트설정.json·채점결과 JSON)과 메모리
#     스냅샷(세트 목록·시험 진행 여부)만 읽습니다.
#   - 쓰기(POST)와 프로그램 동작(시험 시작 등)은 큐를 통해 Tk 스레드에서
#     실행되고, 핸들러는 결과를 짧게 기다렸다가 최신 상태로 응답합니다.
#   - POST 는 세션 토큰(X-Coco-Token 헤더 또는 ?t=) 필수, GET 은 자유.
#     Host 가 루프백이 아니면 거부(DNS 리바인딩 방지). 외부 접속 불가.
#
#   GET  /                    루틴 페이지 (doctype/html/head/body 골격으로 감쌈)
#   GET  /api/state           전체 상태 JSON (5초 폴링)
#   GET  /api/report/<id>     채점결과 HTML
#   POST /api/check           {"key","value"}              → _웹체크 저장
#   POST /api/score           {"date","set","total"}       → mode=수동 기록 추가
#   POST /api/review          {"record_id","done"}         → 오답노트 완료 상태
#   POST /api/action          {"action","set","record_id"} → 시험 시작·오답노트 등
# 응답: 성공 200 {"ok": true, ...최신 state} / 논리 오류 200 {"ok": false,
# "error"} / 토큰 401 / 본문·필드 400. JSON 은 ensure_ascii=False,
# Cache-Control: no-store. 시작 로그에는 쓰기·페이지 요청 요약만 남깁니다
# (/api/state 폴링은 건수만 셈).

ROUTINE_PAGE_HEAD = ('<!doctype html><html lang="ko"><head><meta charset="utf-8">'
                     '<meta name="viewport" content="width=device-width, '
                     'initial-scale=1"></head><body>')
# 골격 끝에 넣는 1줄 스크립트: URL 해시(#quiz 등)가 탭 id(tab-quiz)와 맞으면
# 그 탭을 연다 — 단계 가이드 [웹에서 퀴즈 풀기]가 /?t=…#quiz 로 열기 위함.
# 페이지가 자체 해시 처리를 갖춰도 같은 탭을 한 번 더 클릭할 뿐 무해.
ROUTINE_PAGE_TAIL = ('<script>(function(){try{var h=(location.hash||"")'
                     '.replace(/^#/,"");if(!h){return;}var t=document.'
                     'getElementById("tab-"+h);if(t&&t.getAttribute("role")'
                     '==="tab"){t.click();}}catch(e){}})();</script>'
                     '</body></html>')
ROUTINE_ACTIONS = ("start_exam", "open_review", "start_retry", "open_report",
                   "open_pdf", "show")
ROUTINE_BODY_LIMIT = 1024 * 1024        # POST 본문 상한 (1MB)
ROUTINE_UI_TIMEOUT = 8.0                # 쓰기 작업의 Tk 스레드 응답 대기(초)
ROUTINE_ACTION_TIMEOUT = 1.5            # 동작 요청 대기(초) — 대화상자가 뜨면 먼저 응답
_RECORD_STAMP_RE = re.compile(
    r"(\d{4})-?(\d{2})-?(\d{2})[ T]?(\d{2}):?(\d{2})(?::?(\d{2}))?")


def routine_html_path(base_dir=None):
    """루틴 페이지 파일 경로: <시험장 폴더>/루틴.html → 자동 업데이트 설치 위치
    (_data_target_path('시험장/루틴.html')) 순으로 있는 것. 없으면 첫 후보."""
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
    """루틴.html 내용 → 브라우저에 보낼 HTML. 아티팩트 규격 파일(doctype 없음)은
    골격으로 감싸고, 이미 <!doctype 으로 시작하면 그대로."""
    if text.lstrip()[:9].lower().startswith("<!doctype"):
        return text
    return ROUTINE_PAGE_HEAD + text + ROUTINE_PAGE_TAIL


def routine_missing_page(path):
    """루틴.html 이 없을 때의 안내 페이지 (15초마다 다시 확인)."""
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


# --- 기록 → 규약 records[] ---

def record_stamp_id(when):
    """'2026-09-06 16:12(:36)' → '20260906161236' (초가 없으면 '00'). 실패 None."""
    m = _RECORD_STAMP_RE.search(str(when or ""))
    if not m:
        return None
    y, mo, d, h, mi, sec = m.groups()
    return f"{y}{mo}{d}{h}{mi}{sec or '00'}"


def records_with_ids(records):
    """[(id, 기록)] — 같은 기록은 항상 같은 id(일시 14자리). 같은 분의 두 번째
    기록부터는 초 자리에 순번(01, 02…)을 넣어 유일하게 만듭니다."""
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
    """기록의 세트명 (타이머 일시정지 표시 ' (연습)' 접미 제거)."""
    name = str((r or {}).get("세트명") or "?")
    return name[:-5] if name.endswith(" (연습)") else name


def record_report_paths(r):
    """기록의 리포트 경로 → (존재하는 HTML 경로 or None, 존재하는 JSON 경로 or None)."""
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


_RESULT_CACHE = {}      # 채점결과 JSON 경로 → ((mtime, size), 요약)


def load_result_summary(json_path):
    """채점결과 JSON → {"sheets", "wrong_items", "count", "total"} (mtime 캐시).
    sheets 는 {name, alloc, earned}, wrong_items 는 {label, category, lost, sheet}
    만 남깁니다(웹은 category 문자열·lost 숫자가 있어야 오답으로 셈)."""
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


def serialize_records(records, sets, cfg=None):
    """기록.json → (규약 records[], review{}, 색인{id: 정보}).

    색인 정보: record(원본)·set(세트 dict or None)·norm·name·html·json·count(오답
    수)·date·time·mode — 쓰기 요청(review/action)이 id 로 기록을 찾는 데 씁니다.
    """
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


def serialize_sets(sets, records, index, review):
    """세트 목록 → 규약 sets[] (attempts·best 는 시험·수동 기록, review_done·
    retry_done 은 최근 시험 기록 기준)."""
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
        out.append({
            "name": s["name"], "norm": s.get("norm") or "",
            "pdf": bool(s.get("pdf")) and os.path.isfile(str(s.get("pdf"))),
            "key": bool(s.get("key")),
            "attempts": len(recs), "best": _best(recs),
            "review_done": review_done, "retry_done": retry_done,
        })
    return out


# --- 적응형 일정 → 규약 plan ---

def _iso(d):
    return d.isoformat() if isinstance(d, date) else None


def _seg_end_for(d):
    """날짜가 속한 학습 구간의 마감일 (이 일정은 9/16 하나). 시험일이면 그 날짜."""
    for sg in study_segments():
        if sg["seg_start"] <= d <= sg["end"]:
            return sg["end"]
    return d


def serialize_slot(sl, by_set):
    """엔진 슬롯 → {kind, set{name,norm,pdf}|null, goal, why, best, counts_as_retry}."""
    s = sl.get("set")
    kind = sl.get("kind") or "first"
    if kind not in ("first", "redo", "retry"):
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
        "done": False, "record_id": None,          # (선택 — 웹은 무시)
    }


def _day_title(d, today, kind, slots, promoted, deadline, recs_by_date):
    if kind == "exam":
        return f"시험 {EXAM_DATES.index(d) + 1}"
    if kind == "next":
        return "다음 구간"
    if kind == "past":
        recs = recs_by_date.get(d) or []
        return "완료" if recs else "미완주"
    parts = []
    if slots:
        parts.append(f"{len(slots)}세트" + (" (승격)" if promoted else ""))
    elif d == today and recs_by_date.get(d):
        parts.append("오늘 완료")
    else:
        parts.append("여유일")
    if deadline:
        parts.append("실수 노트")
    return " + ".join(parts)


def serialize_days(adaptive, today, recs_by_date, by_set, dates=None):
    """9/3~9/18 모든 날짜를 규약 days[] 로 (dates 를 주면 그 날짜만).

    kind: exam(시험일 9/17·9/18) > next(현재 구간 end 이후 학습일 — 이 일정엔
    없음) > past(오늘 이전) > mock(슬롯 있음) / rest(슬롯 없음). past·exam·next
    의 slots 는 []이고 오늘·예정일은 엔진의 남은 배정만 담습니다. review_day 는
    규약 호환용으로 남겨 두되 이 일정(시험 후 학습일 없음)에서는 항상 false.
    """
    study = bool(adaptive) and adaptive.get("kind") == "study"
    end = adaptive.get("end") if study else None
    boost = set(adaptive.get("boost_days") or []) if study else set()
    days_map = (adaptive.get("days") or {}) if study else {}
    deadlines = {sg["end"] for sg in study_segments()}
    if dates is None:
        dates = []
        d = ROUTINE_START
        while d <= max(EXAM_DATES):
            dates.append(d)
            d += timedelta(days=1)
    out = []
    for d in dates:
        slots = []
        if d in EXAM_DATES:
            kind, capacity = "exam", 0
        elif study and d > end:
            kind, capacity = "next", base_capacity(d, _seg_end_for(d))
        elif d < today:
            kind, capacity = "past", base_capacity(d, _seg_end_for(d))
        elif not study:
            kind, capacity = "next", base_capacity(d, _seg_end_for(d))
        else:
            slots = [serialize_slot(sl, by_set) for sl in days_map.get(d, [])]
            kind = "mock" if slots else "rest"
            capacity = base_capacity(d, end) + (1 if d in boost else 0)
        promoted = d in boost
        out.append({
            "date": d.isoformat(), "no": routine_day_no(d), "kind": kind,
            "title": _day_title(d, today, kind, slots, promoted, d in deadlines,
                                recs_by_date),
            "capacity": capacity, "promoted": promoted,
            "deadline": d in deadlines, "review_day": False,
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


def serialize_plan(adaptive, today, records, sets, preview=None):
    """build_adaptive_plan 결과 → 규약 plan (date → ISO, 내부 키 제거, days[] 전체,
    다음 학습 구간이 있으면 preview{days, boost_days} — 이 일정은 단일 구간이라 null)."""
    study = adaptive.get("kind") == "study"
    by_set = exam_records_by_set(sets or [], records or [])
    recs_by_date = adaptive.get("_recs_by_date") if study else None
    if not recs_by_date:
        recs_by_date = _recs_by_date(records)
    plan = {
        "kind": adaptive.get("kind"),
        "seg": adaptive.get("seg") if study else None,
        "today": today.isoformat(),
        "start": _iso(adaptive.get("start")) if study else None,
        "end": _iso(adaptive.get("end")) if study else None,
        "seg_start": _iso(adaptive.get("seg_start")) if study else None,
        "reason": str(adaptive.get("reason") or ""),
        "warning": adaptive.get("warning"),
        "remaining": int(adaptive.get("remaining") or 0),
        "retry_target": int(adaptive.get("retry_target") or 0),
        "missed_days": [_iso(x) for x in adaptive.get("missed_days") or []],
        "boost_days": [_iso(x) for x in adaptive.get("boost_days") or []],
        "demoted": [str(x) for x in adaptive.get("demoted") or []],
        "days": serialize_days(adaptive, today, recs_by_date, by_set),
        "preview": None,
    }
    if study and preview and preview.get("kind") == "study":
        p_today = preview.get("today") or preview["start"]
        dates = []
        d = preview["start"]
        while d <= preview["end"]:
            dates.append(d)
            d += timedelta(days=1)
        plan["preview"] = {
            "days": serialize_days(preview, p_today, recs_by_date, by_set,
                                   dates=dates),
            "boost_days": [_iso(x) for x in preview.get("boost_days") or []],
        }
    return plan


def build_state(sets, records=None, cfg=None, today=None, exam_running=False,
                generated=None):
    """규약 GET /api/state 전체 (순수 함수 — 파일은 인자로 안 주면 읽음).

    sets: 스캔된 세트 목록(스냅샷). records/cfg 를 생략하면 기록.json·세트설정.json
    을 읽습니다. today 를 주면 그 날짜 기준(테스트용)."""
    today = today or date.today()
    if records is None:
        records = load_records(RECORDS_PATH)
    if cfg is None:
        cfg = load_set_config(SET_CONFIG_PATH)
    sets = list(sets or [])
    try:
        adaptive = build_adaptive_plan(today, records, sets)
    except Exception as e:
        log_error("루틴 연동 일정 계산", e)
        adaptive = {"kind": segment_for(today)["kind"], "today": today,
                    "days": {}, "reason": "", "warning": f"일정 계산 오류: {e}"}
    preview = None
    nxt = next_segment_after(adaptive.get("seg") or 0) \
        if adaptive.get("kind") == "study" else None
    if nxt:                                   # 다음 학습 구간이 있을 때만 (지금은 없음)
        try:
            preview = build_adaptive_plan(nxt["seg_start"], records, sets)
        except Exception as e:
            log_error("루틴 연동 다음 구간 미리보기", e)
    recs, review, index = serialize_records(records, sets, cfg)
    checks = {str(k): bool(v) for k, v in
              _cfg_section(cfg, WEB_CHECKS_KEY).items()}
    return {
        "version": ROUTINE_API_VERSION,
        "app_version": __version__,
        "today": today.isoformat(),
        "generated": generated or routine_generated(),
        "exam_running": bool(exam_running),
        "plan": serialize_plan(adaptive, today, records, sets, preview),
        "sets": serialize_sets(sets, records, index, review),
        "records": recs,
        "review": review,
        "checks": checks,
        "settings": {"auto_open_routine": bool(
            _cfg_section(cfg, "_설정").get(ROUTINE_AUTO_OPEN_SETTING, False))},
    }


def record_index(sets, records=None, cfg=None):
    """{id: 정보} 색인만 (쓰기 요청이 record_id 로 기록을 찾을 때)."""
    if records is None:
        records = load_records(RECORDS_PATH)
    if cfg is None:
        cfg = load_set_config(SET_CONFIG_PATH)
    return serialize_records(records, sets, cfg)[2]


class RoutineRequestError(Exception):
    """규약의 오류 응답: status(200 논리 오류 / 400 본문·필드 / 401 토큰)."""

    def __init__(self, status, message):
        super().__init__(message)
        self.status = status
        self.message = message


class _RoutineHTTPServer(http.server.ThreadingHTTPServer):
    daemon_threads = True
    # Windows 의 SO_REUSEADDR 는 다른 프로세스가 듣고 있는 포트에도 bind 를 허용해
    # '사용 중 → 대체 포트' 판정이 깨지므로 끕니다(리눅스는 TIME_WAIT 재사용용으로 켬).
    allow_reuse_address = sys.platform != "win32"
    routine = None


class _RoutineHandler(http.server.BaseHTTPRequestHandler):
    """경로 → RoutineServer 메서드 연결. 응답은 항상 JSON(페이지·리포트 제외)."""
    server_version = "CocoRoutine/" + __version__
    sys_version = ""

    def log_message(self, fmt, *args):      # stderr 없는 pythonw 에서도 안전
        pass

    # --- 응답 도우미 ---
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

    # --- 요청 ---
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
    """루틴 웹 연동 로컬 서버 (시험장 프로그램이 상태의 원본).

    sets_fn(): 현재 세트 목록(읽기 전용 스냅샷)을 돌려주는 함수.
    ui_call(fn): fn 을 Tk 스레드에서 실행하도록 넘기는 함수 — 없으면(테스트)
        서버 스레드에서 바로 실행.
    action_fn(action, set, record, norm, record_id): Tk 스레드에서 프로그램 동작
        → (ok, 오류 문구). 없으면 동작은 받기만 하고 ok.
    on_change(kind, info): 쓰기 반영 뒤 Tk 스레드에서 호출(화면 갱신용).
    exam_running_fn(): 시험 진행 중 여부. today: 기준 날짜(date 또는 함수, 테스트용).
    port: 우선 포트(기본 8765) — 사용 중이면 빈 포트. html_path: 루틴 페이지 파일.
    """

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

    # --- 수명 ---
    def start(self):
        """포트 bind(우선 포트 → 빈 포트) 후 데몬 스레드로 serve_forever."""
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
        """페이지 주소 (토큰 포함). tab='quiz' 면 #quiz 로 그 탭을 엽니다."""
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
        """시작 로그 요약 (폴링 제외). 오류 응답은 문구까지."""
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

    # --- 읽기 ---
    def page(self):
        """GET / → (status, HTML). 파일은 매 요청마다 읽어 갱신을 바로 반영."""
        hp = self.html_path()
        try:
            with open(hp, encoding="utf-8-sig") as f:
                return 200, render_routine_page(f.read())
        except OSError:
            return 200, routine_missing_page(hp)

    def _touch_if_files_changed(self):
        """GUI 밖에서(grade.py·손 편집) 기록·설정 파일이 바뀌어도 generated 갱신."""
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
        """record_id → 존재하는 채점결과 HTML 경로 (기록에 있는 것만, 경로 탈출 불가)."""
        rid = str(rid or "").strip()
        if not rid or any(c in rid for c in "/\\") or ".." in rid:
            return None
        info = record_index(self.sets_fn()).get(rid)
        return info["html"] if info and info.get("html") else None

    # --- Tk 스레드 실행 ---
    def _run_on_ui(self, fn, timeout):
        """fn 을 Tk 스레드에서 실행하고 결과를 기다림 → (끝났는지, 결과)."""
        if self.ui_call is None:
            return True, fn()
        box = {}
        ev = threading.Event()

        def job():
            try:
                box["r"] = fn()
            except Exception as e:      # 로그는 호출 쪽에서
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
        """쓰기 작업을 Tk 스레드에서 실행(요청끼리는 직렬화). 응답 없으면 논리 오류로."""
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

    # --- 쓰기 ---
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
        if 1 <= no <= 14:
            record["day"] = plan_day_tag(no)
            record["루틴"] = ROUTINE_TAG

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
                     copied=False, goal=None, linked=False):
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
            self.reopen_btn = tk.Button(
                btns, text="풀이 파일 열기", width=12, font=UI_FONT,
                relief="flat", bg="#2C4A38", fg="white",
                activebackground="#3A5F49", command=self.reopen_file)
            self.reopen_btn.pack(side="left", padx=5)
            self._tick()

        def reopen_file(self):
            """풀이 파일을 Excel로 다시 열기 (창이 사라졌을 때 언제든)."""
            self.app.reopen_student(self.exam)

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


    class ExcelWarnWindow(tk.Toplevel):
        """Excel 창이 확인되지 않을 때 뜨는 비모달 안내 창 (타이머 위).

        사용자의 흐름을 막지 않도록 grab 없이 항상 위에 띄우고, [다시
        열기(Excel 직접)] [기본 프로그램으로 열기] [폴더 열기]를 제공합니다.
        """

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
                                self.json_name, self.checks,
                                total=len(self.items))
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
            # ⑤ 함수 퀴즈 단계에서만 보이는 [웹에서 퀴즈 풀기] (루틴 페이지 퀴즈 탭)
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
            """스텝의 루틴 페이지 탭(#quiz 등). 없으면 None."""
            i = self.selected_index() if i is None else i
            if i is None or i >= len(self.steps):
                return None
            return self.steps[i].get("웹탭") or None

        def open_web_step(self):
            """[웹에서 퀴즈 풀기] — 루틴 페이지의 해당 탭을 브라우저로 엽니다."""
            tab = self.web_tab_of()
            if not tab:
                return None
            return self.app.open_routine_page(tab=tab)

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
            if st.get("웹탭"):
                lines.append("[웹에서 퀴즈 풀기]를 누르면 루틴 페이지의 퀴즈 "
                             "탭이 브라우저로 열립니다. 페이지에서 ⑤를 "
                             "체크하면 이 단계도 자동으로 체크됩니다.")
                if not self.web_btn.winfo_manager():
                    self.web_btn.pack(side="left", padx=6)
            elif self.web_btn.winfo_manager():
                self.web_btn.pack_forget()
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
                if (payload or {}).get("웹탭"):        # 함수 퀴즈 → 루틴 페이지 탭
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

        def __init__(self, scan_root=None, auto_update=True):
            super().__init__()
            self.title(APP_TITLE)
            self.configure(bg=BG)
            self.minsize(680, 520)
            self.scan_root = scan_root or default_scan_root()
            self.sets = []
            self.grade_py = find_grade_py()
            self.exam_running = False
            self._grade_state = None
            self.auto_update = auto_update      # False: 시작 시 확인 안 함(테스트)
            self.updater = UpdateCoordinator()
            self._update_state = None           # 워커 결과 (메인 스레드 폴링)
            self._update_progress = None        # 수동 확인 진행 창
            self._poll_token = 0                # 폴링 세대 (중복 폴링 방지)
            self._toast_after = None
            self._warned_dirty = set()   # 원본 오염 경고를 이미 띄운 세트
            self.plan_no = routine_day_no()   # 오늘의 학습 일정 번호
            self.adaptive = None              # 적응형 일정 (recompute_plan)
            self._pending_plan = None
            self.step_guide = None            # 단계 가이드 창 (열려 있으면)
            self._current_exam = None         # 진행 중 시험 dict
            self._excel_warn = None           # Excel 확인 안내 창
            self.routine = None               # 루틴 웹 연동 서버 (RoutineServer)
            self._routine_jobs = queue.Queue()   # 서버 스레드 → Tk 스레드 작업 큐
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
            if auto_update:
                startup_log("업데이트 확인 시작 (백그라운드, 타임아웃 3초)")
                self._poll_token += 1
                tok = self._poll_token
                threading.Thread(target=self._bg_update_check,
                                 daemon=True).start()
                self.after(1200, lambda: self._update_poll(tok))
            if auto_update and self.routine_auto_open_var.get() \
                    and self.routine is not None:
                self.after(900, self.open_routine_page)   # 시작 시 자동 열기

        def destroy(self):
            """창 종료 시 루틴 서버도 내림 (재시작하는 새 프로세스가 포트를 쓰도록)."""
            srv, self.routine = self.routine, None
            if srv is not None:
                try:
                    srv.stop()
                except Exception:
                    pass
            super().destroy()

        # ---------------- UI 구성 ----------------

        def _build_ui(self):
            header = tk.Frame(self, bg=BRAND)
            header.pack(fill="x")
            tk.Label(header, text=APP_TITLE, bg=BRAND, fg="white",
                     font=("Malgun Gothic", 16, "bold")).pack(
                side="left", padx=18, pady=10)
            tk.Label(header, text="컴활 2급 실기 모의고사 런처", bg=BRAND,
                     fg="#CFE9DA", font=UI_FONT).pack(side="left")

            # 비모달 안내 띠 (자동 업데이트 결과 등) — show_toast()가 pack
            self.toast = tk.Frame(self, bg="#FFF4D6", padx=12, pady=6,
                                  highlightbackground="#E5C87A",
                                  highlightthickness=1, cursor="hand2")
            self.toast_lbl = tk.Label(
                self.toast, text="", bg="#FFF4D6", fg="#5A4300",
                font=("Malgun Gothic", 9), justify="left", anchor="w",
                wraplength=600)
            self.toast_lbl.pack(side="left", fill="x", expand=True)
            tk.Button(self.toast, text="✕", bg="#FFF4D6", fg="#5A4300",
                      relief="flat", font=("Malgun Gothic", 9), padx=4,
                      command=self.hide_toast).pack(side="right")
            self.toast_lbl.bind("<Button-1>", self.hide_toast)
            self.toast.bind("<Button-1>", self.hide_toast)

            # 오늘의 학습 카드 (날짜 기반 루틴 일정)
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
            infofrm = tk.Frame(body, bg=CARD, highlightbackground=LINE,
                               highlightthickness=1)
            infofrm.grid(row=1, column=1, sticky="nsew", padx=(12, 0),
                         pady=(4, 8))
            self.info_lbl = tk.Label(
                infofrm, text="왼쪽 목록에서 세트를 선택하세요.", bg=CARD,
                fg=SUB, font=UI_FONT, justify="left", anchor="nw",
                padx=10, pady=8, wraplength=230)
            self.info_lbl.pack(fill="both", expand=True)
            # 문제지 PDF 회차 충돌 경고 (충돌 시에만 표시)
            self.pdf_warn_lbl = tk.Label(
                infofrm, text="", bg="#FDECEA", fg=RED,
                font=("Malgun Gothic", 9, "bold"), justify="left",
                anchor="w", padx=10, pady=6, wraplength=230)

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
            tk.Button(ctrl, text="시작 로그 보기", font=UI_FONT, relief="groove",
                      padx=10, pady=4, command=self.show_startup_log).pack(
                side="left", padx=4)
            tk.Button(ctrl, text="업데이트 확인", font=UI_FONT, relief="groove",
                      padx=10, pady=4,
                      command=self.manual_update_check).pack(
                side="right", padx=4)
            self.auto_update_var = tk.BooleanVar(value=auto_update_enabled())
            self.auto_update_chk = tk.Checkbutton(
                ctrl, text="자동 업데이트", variable=self.auto_update_var,
                bg=BG, fg=INK, activebackground=BG, font=UI_FONT,
                command=self._toggle_auto_update)
            self.auto_update_chk.pack(side="right", padx=(0, 2))
            self.routine_auto_open_var = tk.BooleanVar(
                value=bool(get_app_setting(ROUTINE_AUTO_OPEN_SETTING, False)))
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
                self.recompute_plan()      # 세트 편입/자동 선택/재배치 반영

        def refresh_records(self):
            records = load_records()
            if getattr(self, "plan_title_lbl", None) is not None:
                self.recompute_plan()      # 새 기록 반영 (밀린 세트 재배치)
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
                f"최고 점수: {best if best is not None else '-'}",
                f"최근 3회: "
                + (" → ".join(str(x) for x in recent) if recent else "-"),
            ]
            self.info_lbl.configure(text="\n".join(lines), fg=INK)
            # 문제지 회차·형·연도가 세트와 다르면(또는 달라서 무시했으면) 경고
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

        # ---------------- 루틴 웹 연동 (v2.4.0) ----------------

        def _start_routine_server(self):
            """로컬 연동 서버 기동: 8765 → 저장된 포트 → 빈 포트. 실패해도 프로그램은
            계속 뜨고(연동 기능만 꺼짐), 실제 포트는 `_설정.루틴포트`에 저장."""
            prefer = ROUTINE_PORT_DEFAULT
            srv = None
            try:
                srv = RoutineServer(
                    sets_fn=lambda: self.sets, ui_call=self._routine_jobs.put,
                    action_fn=self.routine_action,
                    on_change=self.on_routine_changed,
                    exam_running_fn=lambda: self.exam_running, port=prefer)
                srv.start()
            except Exception as e:
                log_error("루틴 서버 시작", e)
                startup_log(f"루틴 서버를 띄우지 못했습니다 ({e}) — 연동 없이 계속")
                srv = None
            self.routine = srv
            if srv is not None:
                try:
                    if get_app_setting(ROUTINE_PORT_SETTING) != srv.port:
                        set_app_setting(ROUTINE_PORT_SETTING, srv.port)
                except Exception:
                    pass
            self.after(100, self._routine_pump)
            return srv

        def _routine_pump(self):
            """서버 스레드가 넘긴 작업을 Tk 스레드에서 실행 (100ms 주기)."""
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
            """[루틴 열기] — 기본 브라우저로 루틴 페이지(토큰 포함 주소) 열기.
            tab='quiz' 면 함수 퀴즈 탭. 반환: 연 주소(실패 None)."""
            if not self.routine_alive():
                messagebox.showwarning(
                    APP_TITLE, "루틴 연동 서버가 실행되지 않아 페이지를 열 수 "
                    "없습니다.\n[시작 로그 보기]에서 '루틴 서버' 줄을 확인하고 "
                    "프로그램을 다시 실행해 보세요.", parent=self)
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
            """시험장 창을 앞으로 (웹에서 동작을 요청했을 때)."""
            try:
                self.deiconify()
                self.lift()
                self.focus_force()
                self.attributes("-topmost", True)
                self.after(400, lambda: self.attributes("-topmost", False))
            except Exception:
                pass

        def _web_pending_plan(self, s, step_kind):
            """웹에서 시작한 시험/재풀이가 오늘 단계 가이드의 어느 스텝인지 찾아
            채점 완료 시 자동 체크되도록 _pending_plan 을 만듭니다."""
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
            """(Tk 스레드) 웹의 /api/action 실행 → (ok, 오류 문구)."""
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
            """(Tk 스레드) 웹 쓰기 반영 뒤 화면 갱신. 오늘의 ⑤ 함수 퀴즈 체크는
            단계 가이드의 '(선택) 함수 퀴즈' 스텝과 동기화합니다."""
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
            """웹 ⑤ 체크 ↔ 오늘 단계 가이드의 함수 퀴즈 스텝 완료 표시."""
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

        # ---------------- 자동 업데이트 ----------------

        def show_toast(self, text, seconds=8):
            """비모달 안내 띠(시작 화면 위쪽). seconds 뒤 자동으로 사라지고
            (0이면 유지) 띠를 클릭하거나 ✕를 누르면 바로 닫힙니다."""
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
            if self.toast.winfo_manager():
                self.toast.pack_forget()

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
            """(워커 스레드) 확인·다운로드·검증. 결과는 _update_state로 전달."""
            try:
                self._update_state = self.updater.check(force=force)
            except Exception as e:
                log_error("업데이트 확인", e)
                self._update_state = "failed"

        def _update_poll(self, token=None, tries=0, manual=False):
            """(메인 스레드) 워커 결과 폴링 — 최대 5분(느린 회선 다운로드)."""
            if token is None:
                token = self._poll_token
            if token != self._poll_token:
                return                      # 새 확인이 시작되어 이 폴링은 종료
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
            """스테이징된 새 버전 적용(시험 중이면 연기) → 재시작."""
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
            """재시작을 미뤄야 하는 상태: 시험 중, 모달 대화상자, 열려 있는
            보조 창(결과·오답노트 등; 단계 가이드는 제외)."""
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
            """열린 창이 없을 때 재시작 (결과 창을 보는 중이면 닫을 때까지
            2초 간격으로 기다림). 파일은 이미 교체되어 있어 언제 다시 시작해도
            새 버전이 뜹니다."""
            if self.exam_running:
                return False          # exam_closed → _apply_update_now 경유
            if self._busy_for_restart():
                if tries == 0:
                    startup_log("재시작 대기: 열린 창·대화상자가 닫히면 재시작")
                self.after(2000, lambda: self._restart_when_idle(tries + 1))
                return False
            return self._try_restart()

        def _try_restart(self):
            """새 프로세스를 먼저 띄운 뒤 현재 창 종료 (성공 여부 반환).

            os.execl은 Windows에서 새 프로세스를 띄우고 현재 프로세스를 즉시
            끝내는 방식이라 자식이 부모 종료·콘솔 문제를 겪을 수 있어
            분리된 Popen으로 대체. 실패하면 직접 실행 안내.
            """
            if self.exam_running:
                return False  # 시험 중에는 재시작하지 않음
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
            """[업데이트 확인] — 자동 업데이트 토글과 무관하게 즉시 확인·적용."""
            if self.updater.checking or self._update_progress is not None:
                messagebox.showinfo(APP_TITLE, "이미 업데이트를 확인하는 중입니다.",
                                    parent=self)
                return None
            if self.updater.pending:          # 시험 중이라 미뤄 둔 업데이트
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
            routine_touch()                    # 루틴 페이지: 시험 진행 중 표시
            self.start_btn.configure(state="disabled", text="진행 중")
            # 문제지 PDF를 먼저 열고 → 잠시 후 Excel (Excel 창이 맨 앞에 오게)
            delay = 0
            if s.get("pdf"):
                ok_pdf, err_pdf = open_file(s["pdf"])
                startup_log(f"PDF 실행: {'성공' if ok_pdf else '실패 ' + err_pdf}"
                            f" · {s['pdf']}")
                delay = EXCEL_OPEN_DELAY_MS
            self.after(delay, lambda: self._launch_workbook(exam))
            TimerWindow(self, exam)

        def _launch_workbook(self, exam, prefer_excel=True, verify=True,
                             source="시험 시작"):
            """풀이 파일을 Excel로 열고 로그 + (verify면) 5초 뒤 실행 확인.
            반환 (성공, 방법, 오류)."""
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
            return True, method, err

        def _verify_excel(self, exam, tries):
            """Excel 실행 후 확인: 잠금 파일/EXCEL.EXE 프로세스. 두 번(5초+3초)
            확인해도 없으면 안내 창 (Windows 외에서는 판단하지 않음)."""
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
            """[풀이 파일 열기]/[다시 열기]: 언제든 풀이 파일 재오픈.
            반환 (성공, 방법, 오류)."""
            return self._launch_workbook(
                exam, prefer_excel=prefer_excel, verify=True,
                source="다시 열기(Excel 직접)" if prefer_excel
                else "기본 프로그램으로 열기")

        def show_startup_log(self):
            """[시작 로그 보기] — 시작 로그 텍스트 창 + 클립보드 복사."""
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
            if self.updater.pending:      # 시험 중 미뤄 둔 업데이트 → 지금 적용
                self.after(500, self._apply_update_now)

        # ---------------- 오늘의 학습 ----------------

        def plan_for(self, no):
            """일정 번호 -> 적응형 plan (계산 실패/범위 밖이면 v4 고정 일정)."""
            try:
                return adaptive_day_plan(self.adaptive, no)
            except Exception as e:
                log_error("적응형 일정 조회", e)
                return plan_for_day(no)

        def recompute_plan(self):
            """[일정 다시 계산] — 기록·세트를 다시 읽어 적응형 일정 재계산."""
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
                todo += "\n재배치: " + plan["사유"]
            if plan.get("경고"):
                todo += "\n경고: " + plan["경고"]
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
            plan = self.plan_for(self.plan_no)
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

        def open_review_mode(self, full_only=False, json_path=None,
                             set_info=None):
            """오답노트 모드. json_path 를 주면 그 채점결과(웹 [오답노트 열기]),
            아니면 선택 세트의 최신 채점결과."""
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
                             copied=copied, goal=goal,
                             linked=self.routine_alive())
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
    app = ExamApp(auto_update=False)     # 네트워크 확인 없이
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
    # 오늘의 학습 카드 + 미리보기 화살표 (9/3 시작, 시험 9/17·9/18 D-day)
    assert app.plan_title_lbl.cget("text"), "오늘의 학습 제목 비어 있음"
    app.plan_no = 3
    app._render_plan_card()
    t3 = app.plan_title_lbl.cget("text")
    assert "Day 3" in t3 and "9/5(토)" in t3, t3
    assert any(k in t3 for k in ("오늘", "예정", "완료", "미완주", "1차 완주")), t3
    assert app.recalc_btn.cget("text") == "일정 다시 계산"
    ad = app.recompute_plan()
    assert ad is None or ad.get("kind") in ("study", "exam", "after"), ad
    if ad and ad.get("kind") == "study":
        today_plan = app.plan_for(routine_day_no())
        assert today_plan.get("적응형") is True
        if today_plan.get("스텝"):
            gA = StepGuideWindow(app, today_plan)   # 적응형 plan으로 단계 창
            app.update_idletasks()
            app.update()
            assert gA.listbox.size() == len(today_plan["스텝"])
            gA.destroy()
    assert "시험1 D" in t3 and "시험2 D" in t3, t3
    app._shift_plan(1)
    assert "Day 4" in app.plan_title_lbl.cget("text")
    app.plan_no = 14
    app._shift_plan(1)                       # d14(9/16) 다음은 시험 1 (시간순)
    assert app.plan_no == PLAN_EXAM1
    assert "시험일" in app.plan_title_lbl.cget("text")
    app._shift_plan(1)                       # 이틀 연속: 시험 1 → 시험 2
    assert app.plan_no == PLAN_EXAM2
    assert "9/18(금)" in app.plan_title_lbl.cget("text")
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
    # v2.2.1: 시작 로그 창 / Excel 확인 안내 창 / 타이머 [풀이 파일 열기]
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
    assert app._show_excel_warning(exam, "갱신") is warn      # 창 재사용
    assert "갱신" in warn.detail_lbl.cget("text")
    _orig_warn = messagebox.showwarning
    warned = []
    messagebox.showwarning = lambda *a, **k: warned.append(a)
    try:
        ok_re, _m, err_re = app.reopen_student(exam)      # student="" → 안내만
    finally:
        messagebox.showwarning = _orig_warn
    assert ok_re is False and "풀이 파일이 없습니다" in err_re and warned
    app._current_exam = exam
    app.exam_closed()                       # 안내 창 정리 + closed 표시
    assert exam["closed"] is True and not warn.winfo_exists()
    app._verify_excel(exam, 0)              # closed면 아무것도 안 함
    # 세트 정보 패널: 문제지 회차 충돌 경고
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
    # v2.3.0: 안내 띠 / 자동 업데이트 토글 / 확인 결과 처리 (네트워크 없이)
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
    assert app._apply_update_now() == "none"          # 대기 중 업데이트 없음
    assert app._on_update_checked("latest") == "latest"  # 자동 경로: 대화상자 없음
    app._on_update_checked("failed")
    assert "자동 업데이트 실패" in app.toast_lbl.cget("text")
    app.hide_toast()
    app.updater.data_synced = 2
    app._on_update_checked("data")
    assert "기대값 파일 2개" in app.toast_lbl.cget("text")
    app.hide_toast()
    assert app._busy_for_restart() is True            # 타이머 창 열림 → 대기
    timer.withdraw()
    app.update_idletasks()
    app.update()
    assert app._busy_for_restart() is False
    _dw = DiagnosisWindow(app, "x")
    app.update_idletasks()
    app.update()
    assert app._busy_for_restart() is True            # 보조 창 열림 → 재시작 대기
    _dw.destroy()
    app.update_idletasks()
    app.update()
    assert app._busy_for_restart() is False
    timer.deiconify()                                 # (실제 재시작은 호출 안 함)
    # 재시작 명령은 만들기만 (실행하지 않음)
    r_args, r_kw = restart_command(argv=["시험장.py"])
    assert r_args[0] == sys.executable and r_args[1].endswith("시험장.py") \
        and r_kw["cwd"] == BASE_DIR, (r_args, r_kw)
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
    # v2.4.0: 루틴 웹 연동 서버 — 기동·상태 조회·[루틴 열기]/[웹에서 퀴즈 풀기]
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
    # 쓰기: 토큰 없으면 401, 토큰 있으면 Tk 스레드(펌프)에서 저장 후 최신 상태
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
        app.update()                       # 펌프(after)가 작업을 실행하도록
        time.sleep(0.05)
    assert _box.get("r", {}).get("ok") is True, _box
    assert _box["r"]["checks"].get("smoke|quiz") is True
    _cfg = load_set_config()
    assert _cfg.get(WEB_CHECKS_KEY, {}).get("smoke|quiz") is True, "_웹체크 저장"
    _cfg.get(WEB_CHECKS_KEY, {}).pop("smoke|quiz", None)
    save_set_config(_cfg)
    # 단계 가이드: 마지막(함수 퀴즈) 스텝에서만 [웹에서 퀴즈 풀기] 표시
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
    g7.start_step()                                  # 퀴즈 스텝 → 웹 탭 열기
    assert _opened == ["quiz"], _opened
    g7.destroy()
    save_step_progress("d03", set())
    del app.open_routine_page
    # 채점 결과 창: 연동 중이면 '자동 반영' 문구
    rw = ResultWindow(app, {"total": 80, "pass_line": 70, "sheets": []}, "",
                      linked=True)
    app.update_idletasks()
    app.update()
    assert "자동 반영" in rw.link_lbl.cget("text")
    rw.destroy()
    # 동작 브리지: 시험 중이면 오류, show 는 ok
    app.exam_running = True
    assert app.routine_action("start_exam", None, None, None, None) == \
        (False, "시험 진행 중")
    app.exam_running = False
    assert app.routine_action("show", None, None, None, None) == (True, None)
    assert app.routine_action("open_pdf", {"name": "x", "pdf": None}, None,
                              "x", None)[0] is False
    timer.finished = True
    timer.destroy()
    app.destroy()
    assert not app.routine.running if app.routine else True
    print("SMOKE OK: 창 생성/위젯 렌더/타이머/오답노트 패널/단계 가이드(세트 "
          "자동 선택·바꾸기·미발견 직접 선택)/오류 대화상자·로그/진단 창/"
          "시작 로그 창/Excel 확인 안내 창/PDF 회차 경고/안내 띠·자동 업데이트 "
          "토글/루틴 연동 서버(상태·쓰기·페이지·퀴즈 버튼·결과 문구)/파괴 정상")


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


def _show_fatal(title, msg):
    """치명적 시작 오류 표시: tkinter messagebox → ctypes MessageBoxW → stderr.
    (pythonw에서는 stderr가 없어 창으로 보여 주지 않으면 아무것도 안 보임)"""
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
