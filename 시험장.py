#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
코코 시험장 — 컴활 2급 실기 모의고사 런처

시작 화면에서 모의고사 세트를 고르고 [시험 시작]을 누르면
문제 파일 사본이 Excel로 열리고 타이머가 시작됩니다.
제출하면 grade.py로 자동 채점해 점수와 리포트를 보여 줍니다.

의존성: Python 표준 라이브러리 + tkinter (채점은 grade.py/openpyxl 필요)
"""

__version__ = "1.0.0"

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
from datetime import datetime

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


def _ensure_records_home():
    """기록.json 보관 폴더 생성 + 구버전 위치의 기록 자동 이전."""
    try:
        os.makedirs(RECORDS_DIR, exist_ok=True)
        if os.path.isfile(_OLD_RECORDS_PATH) \
                and not os.path.isfile(RECORDS_PATH):
            shutil.move(_OLD_RECORDS_PATH, RECORDS_PATH)
    except OSError:
        pass

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


def scan_sets(root):
    """폴더 재귀 스캔: *_문제.(xlsx|xlsm) + 같은 폴더의 *_정답.* 페어 감지.

    반환: [{"name","dir","problem","answer","key","pdf"}]
    """
    sets = []
    if not root or not os.path.isdir(root):
        return sets
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if not d.startswith((".", "__"))]
        fileset = set(filenames)
        for fn in sorted(filenames):
            m = PROBLEM_RE.match(fn)
            if not m:
                continue
            name = m.group("name")
            if name.startswith("풀이_"):
                continue
            answer = None
            for ext in (".xlsm", ".xlsx"):
                cand = f"{name}_정답{ext}"
                if cand in fileset:
                    answer = os.path.join(dirpath, cand)
                    break
            if not answer:
                continue
            key = os.path.join(dirpath, f"{name}_기대값.json")
            pdf = os.path.join(dirpath, f"{name}_문제지.pdf")
            sets.append({
                "name": name,
                "dir": dirpath,
                "problem": os.path.join(dirpath, fn),
                "answer": answer,
                "key": key if os.path.isfile(key) else None,
                "pdf": pdf if os.path.isfile(pdf) else None,
            })
    sets.sort(key=lambda s: s["name"])
    return sets


def build_direct_set(problem, answer):
    """직접 선택한 문제/정답 파일로 세트 정보 구성 (기대값/문제지 자동 감지)."""
    base = os.path.basename(problem)
    m = PROBLEM_RE.match(base)
    name = m.group("name") if m else os.path.splitext(base)[0]
    d = os.path.dirname(problem)
    key = os.path.join(d, f"{name}_기대값.json")
    pdf = os.path.join(d, f"{name}_문제지.pdf")
    return {
        "name": name,
        "dir": d,
        "problem": os.path.abspath(problem),
        "answer": os.path.abspath(answer),
        "key": key if os.path.isfile(key) else None,
        "pdf": pdf if os.path.isfile(pdf) else None,
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
    """세트별 최고점과 최근 3회 점수. {세트명: {"best":.., "recent":[..]}}"""
    by = {}
    for r in records:
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


def make_attempt_copy(problem, set_name, when=None):
    """문제 파일을 풀이_<세트명>_<yyyymmdd_HHMM>.<확장자> 사본으로 복사."""
    when = when or datetime.now()
    ext = os.path.splitext(problem)[1]
    stamp = when.strftime("%Y%m%d_%H%M")
    d = os.path.dirname(os.path.abspath(problem))
    dst = os.path.join(d, f"풀이_{set_name}_{stamp}{ext}")
    n = 2
    while os.path.exists(dst):
        dst = os.path.join(d, f"풀이_{set_name}_{stamp}_{n}{ext}")
        n += 1
    shutil.copy2(problem, dst)
    return dst


def run_grading(grade_py, problem, answer, student,
                key=None, html=None, json_out=None, history=None,
                timeout=600):
    """grade.py를 서브프로세스로 실행.

    반환: (결과 dict 또는 None, stdout, stderr, returncode)
    결과 dict는 --json 출력이 우선, 실패 시 콘솔 총점 파싱 폴백.
    """
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
            if self._detail:
                self._toggle_btn = tk.Button(
                    btns, text="자세히 보기", command=self._toggle,
                    font=UI_FONT, relief="groove")
                self._toggle_btn.pack(side="left")
            tk.Button(btns, text="닫기", command=self.destroy,
                      font=UI_FONT, relief="groove").pack(side="right")
            self._body = frm
            self.grab_set()

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


    class ResultWindow(tk.Toplevel):
        """채점 결과 창: 큰 점수 + 시트별 점수 + 리포트 열기."""

        def __init__(self, master, result, html_path, folder=None):
            super().__init__(master)
            self.title(f"{APP_TITLE} - 채점 결과")
            self.configure(bg=BG)
            self.attributes("-topmost", True)
            total = result.get("total", 0)
            passed = total >= result.get("pass_line", PASS_LINE)
            frm = tk.Frame(self, bg=BG, padx=28, pady=20)
            frm.pack(fill="both", expand=True)
            tk.Label(frm, text="채점 결과", bg=BG, fg=SUB,
                     font=UI_FONT).pack()
            tk.Label(frm, text=f"{total}점", bg=BG,
                     fg=BRAND if passed else RED,
                     font=("Malgun Gothic", 42, "bold")).pack()
            verdict = "합격권" if passed else "미달"
            tk.Label(frm, text=f"합격선 {PASS_LINE}점 기준: {verdict}",
                     bg=BG, fg=INK, font=UI_FONT_BOLD).pack(pady=(0, 10))
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

            tk.Label(self, text=exam["set"]["name"], bg=INK, fg="#A7C8B2",
                     font=("Malgun Gothic", 9)).pack(padx=16, pady=(10, 0))
            self.time_lbl = tk.Label(self, text=self._fmt(), bg=INK,
                                     fg="#7BD59A", font=DIGIT_FONT)
            self.time_lbl.pack(padx=24, pady=(0, 2))
            self.status_lbl = tk.Label(self, text="시험 진행 중", bg=INK,
                                       fg="#A7C8B2", font=("Malgun Gothic", 9))
            self.status_lbl.pack()
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
                                      activestyle="none")
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
            tk.Button(ctrl, text="직접 선택...", font=UI_FONT, relief="groove",
                      padx=10, pady=4, command=self.choose_direct).pack(
                side="left", padx=4)
            tk.Button(ctrl, text="새로 고침", font=UI_FONT, relief="groove",
                      padx=10, pady=4, command=self.refresh_all).pack(
                side="left", padx=4)
            tk.Button(ctrl, text="업데이트 확인", font=UI_FONT, relief="groove",
                      padx=10, pady=4,
                      command=self.manual_update_check).pack(
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
            self.sets = scan_sets(self.scan_root) + direct
            self.listbox.delete(0, "end")
            for s in self.sets:
                mark = "[직접] " if s.get("direct") else ""
                self.listbox.insert("end", f" {mark}{s['name']}")
            if self.sets:
                self.listbox.selection_set(0)
                self._show_info()

        def refresh_records(self):
            records = load_records()
            if not records:
                self.records_lbl.configure(text="아직 응시 기록이 없습니다.")
                return
            lines = []
            for r in records[-5:][::-1]:
                score = r.get("점수")
                score_s = f"{score}점" if score is not None else "채점 실패"
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
            summ = records_summary(load_records()).get(s["name"], {})
            best = summ.get("best")
            recent = summ.get("recent") or []
            lines = [
                f"세트: {s['name']}",
                f"폴더: {s['dir']}",
                f"정답 파일: 있음",
                f"기대값 JSON: {'있음' if s['key'] else '없음'}",
                f"문제지 PDF: {'있음' if s['pdf'] else '없음'}",
                "",
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

        def choose_direct(self):
            problem = filedialog.askopenfilename(
                parent=self, title="문제 파일 선택",
                filetypes=[("Excel 파일", "*.xlsx *.xlsm"), ("모든 파일", "*.*")])
            if not problem:
                return
            answer = filedialog.askopenfilename(
                parent=self, title="정답 파일 선택",
                initialdir=os.path.dirname(problem),
                filetypes=[("Excel 파일", "*.xlsx *.xlsm"), ("모든 파일", "*.*")])
            if not answer:
                return
            s = build_direct_set(problem, answer)
            s["direct"] = True
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

        def start_exam(self):
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
            try:
                student = make_attempt_copy(s["problem"], s["name"])
            except OSError as e:
                CollapsibleErrorDialog(
                    self, APP_TITLE, "풀이 사본을 만들 수 없습니다.", str(e))
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
            exam = {
                "set": s,
                "student": student,
                "minutes": max(MIN_MINUTES,
                               min(MAX_MINUTES, int(self.minutes_var.get()
                                                    or DEFAULT_MINUTES))),
                "started": datetime.now(),
            }
            self.exam_running = True
            self.start_btn.configure(state="disabled", text="시험 진행 중")
            TimerWindow(self, exam)

        def exam_closed(self):
            self.exam_running = False
            self.start_btn.configure(state="normal", text="시험 시작")
            self.refresh_records()
            self._show_info()

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

            def do_run():
                return run_grading(
                    self.grade_py, exam["set"]["problem"],
                    exam["set"]["answer"], exam["student"],
                    key=exam["set"]["key"], html=html_path,
                    json_out=json_path, history=RECORDS_PATH)

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
                ResultWindow(self, result, html_path,
                             folder=os.path.dirname(html_path))
                if os.path.isfile(html_path):
                    open_file(html_path)
            record = {
                "일시": exam["started"].strftime("%Y-%m-%d %H:%M"),
                "세트명": exam["set"]["name"] + (" (연습)" if practice else ""),
                "점수": score,
                "소요시간": format_elapsed(elapsed),
                "리포트": html_path if os.path.isfile(html_path) else None,
            }
            try:
                append_record(record)
            except OSError as e:
                messagebox.showwarning(
                    APP_TITLE, f"기록.json 저장에 실패했습니다: {e}",
                    parent=self)


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
    timer.finished = True
    timer.destroy()
    app.destroy()
    print("SMOKE OK: 창 생성/위젯 렌더/타이머/파괴 정상")


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
