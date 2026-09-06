# -*- coding: utf-8 -*-
"""
코코 채점 더블클릭 실행 도구 (.pyw)

스마트 앱 컨트롤이 .bat 실행을 차단하는 환경용입니다. 더블클릭하면
파일 선택 창 3개(문제 -> 정답 -> 내 풀이)로 파일을 고르고, 같은 폴더의
grade.py로 채점한 뒤 점수를 표시하고 HTML 리포트를 엽니다.
콘솔이 없어도 모든 오류를 메시지 창으로 표시합니다.

v1.7.0: 시험장 v2.3.0과 같은 규칙으로 동작
  - 완전 자동 업데이트: 실행할 때마다 version.json(타임아웃 3초)을 확인해
    새 버전이면 묻지 않고 files(프로그램)+data_files(기대값)를 내려받아
    검증(sha256·__version__·py_compile·json)·백업·적용. 같은 버전이어도
    로컬에 없거나 해시가 다른 기대값은 루트/기대값/에 내려받음.
    세트설정.json(시험장 폴더) `_설정.자동업데이트`=false면 확인 안 함.
  - 기대값 JSON 자동 연결: 문제 파일 폴더의 정확 키 → 루트/기대값/의 정확
    키 → 연도·회차·형 토큰이 2개 이상 겹치고 유일한 파일 순.
  - 로그: 채점결과/코코채점_로그.txt
"""

__version__ = "1.7.0"

import hashlib
import importlib.util
import json
import os
import re
import shutil
import subprocess
import sys
import traceback
import webbrowser
from datetime import datetime

APP = "코코 채점"
PASS_LINE = 70
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPDATE_BASE_URL = ("https://raw.githubusercontent.com/"
                   "xasmine02/cocomate-study-tools/main/")
LOG_PATH = os.path.join(BASE_DIR, "채점결과", "코코채점_로그.txt")
LOG_KEEP = 200
INSTALLED_VERSION_PATH = os.path.join(BASE_DIR, "채점결과", ".installed_version")
EXPECTED_DIR_NAME = "기대값"
AUTO_UPDATE_SETTING = "자동업데이트"
MANUAL_PIP = ("수동 설치: 명령 프롬프트에서  py -m pip install openpyxl\n"
              "(위 명령이 안 되면: python -m pip install openpyxl)")
_VERSION_MARK_RE = re.compile(r'^__version__\s*=\s*["\']([^"\']+)["\']',
                              re.MULTILINE)


def _log(message, path=None, keep=LOG_KEEP):
    """로그 1줄 append (채점결과/코코채점_로그.txt, 최근 keep줄). 실패는 무시."""
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{stamp}] 코코채점 {__version__} " + str(message).replace("\n", " | ")
    p = path or LOG_PATH
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


def _pip_python(exe=None):
    """pip 실행에 쓸 인터프리터. pythonw.exe면 같은 폴더 python.exe로 치환."""
    exe = exe or sys.executable
    base = os.path.basename(exe).lower()
    if base.startswith("pythonw"):
        cand = os.path.join(os.path.dirname(exe),
                            base.replace("pythonw", "python", 1))
        if os.path.isfile(cand):
            return cand
    return exe


def install_openpyxl(exe=None, timeout=600):
    """pip로 openpyxl 설치 (창 숨김). (성공 여부, 출력 요약) 반환."""
    py = _pip_python(exe)
    kwargs = {}
    if sys.platform == "win32":
        kwargs["creationflags"] = 0x08000000  # CREATE_NO_WINDOW
    try:
        proc = subprocess.run(
            [py, "-m", "pip", "install", "openpyxl"],
            capture_output=True, text=True, encoding="utf-8",
            errors="replace", timeout=timeout, **kwargs)
    except Exception as e:
        return False, str(e)
    log = ((proc.stdout or "") + "\n" + (proc.stderr or "")).strip()
    if len(log) > 1200:
        log = log[-1200:]
    return proc.returncode == 0, log


def openpyxl_available():
    """이 인터프리터(=grade.py를 실행할 인터프리터)에 openpyxl이 있는가."""
    try:
        return importlib.util.find_spec("openpyxl") is not None
    except Exception:
        return False


def classify_error(rc, text):
    """채점 실패 원인 분류: 'module' | 'file' | 'other'."""
    t = text or ""
    if rc == 3 or "ModuleNotFoundError" in t or "ImportError" in t \
            or "openpyxl 라이브러리가 설치되어" in t:
        return "module"
    if "FileNotFoundError" in t or "BadZipFile" in t \
            or "찾을 수 없습니다" in t or "열 수 없습니다" in t:
        return "file"
    return "other"


def _show_message(title, msg, error=True):
    """콘솔이 없어도 보이는 알림: tkinter -> ctypes -> stderr 순."""
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        (messagebox.showerror if error else messagebox.showinfo)(title, msg)
        root.destroy()
        return
    except Exception:
        pass
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(None, str(msg), str(title),
                                         0x10 if error else 0x40)
        return
    except Exception:
        pass
    try:
        if sys.__stderr__:
            sys.__stderr__.write(f"{title}: {msg}\n")
    except Exception:
        pass


# ---------------- 세트 식별 · 기대값 자동 연결 (시험장.py와 같은 규칙) ----------------

_ROLE_TOKENS_STRIP = ("문제지", "기대값", "정답지", "답안지", "정답", "답안",
                      "문제")


def norm_set_key(name):
    """파일명 -> 세트 식별 키 (시험장.py norm_set_key와 동일)."""
    s = os.path.splitext(os.path.basename(str(name)))[0].lower()
    s = re.sub(r"[\s_\-()\[\]{}.,·~]+", "", s)
    for tok in _ROLE_TOKENS_STRIP:
        s = s.replace(tok, "")
    return s


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
    for m in re.finditer(r"(?<!\d)(2[0-9])(?!\d)", s):
        toks.add(m.group(1))
    for word in ("상시", "코코", "모의", "복원", "기출", "실기", "필기",
                 "드릴", "계산", "컴활"):
        if word in s:
            toks.add(word)
    return toks


def _token_conflict(a, b):
    for pat in (r"20\d{2}", r"\d{1,2}회", r"[ab가나]형"):
        ca = {t for t in a if re.fullmatch(pat, t)}
        cb = {t for t in b if re.fullmatch(pat, t)}
        if ca and cb and not (ca & cb):
            return True
    return False


def _match_unique(toks, paths, min_shared=2):
    """토큰이 min_shared개 이상 겹치고 충돌 없는 유일한 최고점 후보 (없으면 None)."""
    scored = []
    for p in paths:
        pt = set_tokens(os.path.basename(p))
        shared = 0 if _token_conflict(toks, pt) else len(toks & pt)
        if shared >= min_shared:
            scored.append((shared, p))
    if not scored:
        return None
    best = max(s for s, _p in scored)
    matched = [p for s, p in scored if s == best]
    return matched[0] if len(matched) == 1 else None


def expected_values_dir(base_dir=None):
    """자동 배포되는 기대값 폴더: <루트>/채점/코코채점.pyw 구조면 <루트>/기대값,
    한 폴더에 전부 있는 평면 구조면 <폴더>/기대값."""
    base_dir = base_dir or BASE_DIR
    root = os.path.dirname(base_dir)
    if os.path.basename(base_dir) == "채점" \
            or os.path.isdir(os.path.join(root, "시험장")):
        return os.path.join(root, EXPECTED_DIR_NAME)
    return os.path.join(base_dir, EXPECTED_DIR_NAME)


def derive_set_name(problem_path):
    """문제 파일명에서 세트명 추출 (…_문제.xlsx -> …)."""
    base = os.path.splitext(os.path.basename(str(problem_path)))[0]
    for suf in ("_문제", "-문제", " 문제"):
        if base.endswith(suf):
            return base[: -len(suf)]
    return base


def detect_key(problem_path, key_dirs=None):
    """문제 파일의 기대값 JSON 찾기. 반환: (경로 또는 None, 출처 문구).

    ① 문제 파일 폴더의 정확 키(정규화 키 동일) ② 루트/기대값 폴더(key_dirs)의
    정확 키 ③ 두 곳 후보 중 토큰(연도·회차·형)이 2개 이상 겹치고 유일한 파일.
    같은 이름이 두 곳에 있으면 문제 파일 폴더 쪽을 씁니다.
    """
    d = os.path.dirname(os.path.abspath(problem_path))
    k = norm_set_key(problem_path)
    if key_dirs is None:
        key_dirs = [expected_values_dir()]
    cands = []
    try:
        for fn in sorted(os.listdir(d)):
            if fn.lower().endswith(".json") and "기대값" in fn \
                    and not fn.startswith("풀이_"):
                if norm_set_key(fn) == k:
                    return os.path.join(d, fn), "문제 파일 폴더"
                cands.append(os.path.join(d, fn))
    except OSError:
        pass
    have = {os.path.basename(p) for p in cands}
    for kd in key_dirs:
        if not os.path.isdir(kd) or os.path.abspath(kd) == d:
            continue
        try:
            names = sorted(os.listdir(kd))
        except OSError:
            continue
        for fn in names:
            if not (fn.lower().endswith(".json") and "기대값" in fn):
                continue
            if norm_set_key(fn) == k:
                return os.path.join(kd, fn), "기대값 폴더"
            if fn not in have:
                cands.append(os.path.join(kd, fn))
    toks = set_tokens(derive_set_name(problem_path)) | set_tokens(os.path.basename(d))
    cand = _match_unique(toks, cands)
    return (cand, "토큰 일치") if cand else (None, None)


# ---------------- 자동 업데이트 (공개 저장소, 시험장.py v2.3.0과 같은 규칙) ----------------


def _version_tuple(v):
    nums = re.findall(r"\d+", str(v or ""))
    return tuple(int(x) for x in nums[:3]) if nums else (0,)


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
            for k in ("files", "data_files", "sha256"):
                if not isinstance(info.get(k), dict):
                    info[k] = {}
            return info
    except Exception:
        pass
    return None


def update_available(info, current=None):
    return bool(info) and _version_tuple(info.get("version")) > \
        _version_tuple(current if current is not None else __version__)


def version_marker(text):
    m = _VERSION_MARK_RE.search(text or "")
    return m.group(1) if m else None


def _read_text(path):
    try:
        with open(path, encoding="utf-8", errors="replace") as f:
            return f.read()
    except OSError:
        return ""


def current_version(base_dir=None):
    """설치된 도구 모음 버전(version.json 기준 = 시험장.py __version__).
    시험장.py를 못 찾으면 마지막 적용 기록(.installed_version), 그것도
    없으면 0.0.0(한 번 적용 후 기록됨)."""
    base_dir = base_dir or BASE_DIR
    root = os.path.dirname(base_dir)
    for p in (os.path.join(root, "시험장", "시험장.py"),
              os.path.join(base_dir, "시험장.py")):
        if os.path.isfile(p):
            v = version_marker(_read_text(p))
            if v:
                return v
    v = _read_text(INSTALLED_VERSION_PATH).strip()
    return v or "0.0.0"


def auto_update_enabled(base_dir=None):
    """시험장 폴더의 세트설정.json `_설정.자동업데이트` (없으면 True)."""
    base_dir = base_dir or BASE_DIR
    root = os.path.dirname(base_dir)
    for p in (os.path.join(root, "시험장", "세트설정.json"),
              os.path.join(base_dir, "세트설정.json")):
        if os.path.isfile(p):
            try:
                with open(p, encoding="utf-8-sig") as f:
                    cfg = json.load(f)
                v = (cfg.get("_설정") or {}).get(AUTO_UPDATE_SETTING, True)
                return bool(v)
            except Exception:
                return True
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


def _data_target_path(rel_key, base_dir=None):
    parts = [p for p in str(rel_key).split("/") if p not in ("", ".", "..")]
    if parts and parts[0] == EXPECTED_DIR_NAME:
        parts = parts[1:]
    if not parts:
        raise RuntimeError(f"잘못된 기대값 경로: {rel_key!r}")
    return os.path.join(expected_values_dir(base_dir), *parts)


def _verify_download(repo_rel, data, kind, info, expect_version=None):
    expected = (info.get("sha256") or {}).get(repo_rel)
    if expected:
        got = hashlib.sha256(data).hexdigest()
        if got.lower() != str(expected).lower():
            raise RuntimeError(f"{repo_rel}: sha256 불일치")
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError as e:
        raise RuntimeError(f"{repo_rel}: UTF-8 텍스트가 아닙니다 ({e})")
    if kind == "code":
        ver = version_marker(text)
        if ver is None:
            raise RuntimeError(f"{repo_rel}: __version__ 표식 없음")
        if expect_version and _version_tuple(ver) != _version_tuple(expect_version):
            raise RuntimeError(f"{repo_rel}: 파일 버전 {ver} ≠ version.json {expect_version}")
    else:
        try:
            json.loads(text)
        except ValueError as e:
            raise RuntimeError(f"{repo_rel}: JSON 형식 오류 ({e})")
    return text


def _cleanup_staged(staged):
    for target, tmp, _kind in staged:
        for leftover in (tmp, tmp + "c"):
            try:
                if os.path.isfile(leftover):
                    os.remove(leftover)
            except OSError:
                pass


def _stage_items(items, info, base_url, timeout):
    """[(repo_rel, target, kind)] 다운로드 → 검증 → target.new. 실패 시 정리+예외."""
    import py_compile
    import urllib.parse
    staged = []
    try:
        for repo_rel, target, kind in items:
            data = _http_get(base_url + urllib.parse.quote(str(repo_rel)), timeout)
            expect = info.get("version") if (
                kind == "code" and os.path.basename(target) == "시험장.py") else None
            _verify_download(repo_rel, data, kind, info, expect_version=expect)
            tmp = target + ".new"
            os.makedirs(os.path.dirname(target) or ".", exist_ok=True)
            with open(tmp, "wb") as f:
                f.write(data)
            staged.append((target, tmp, kind))
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


def commit_staged(staged):
    """.bak 백업 후 교체. 하나라도 실패하면 전체 롤백. (성공, 메시지)"""
    replaced = []
    try:
        for target, tmp, _kind in staged:
            existed = os.path.isfile(target)
            if existed:
                shutil.copy2(target, target + ".bak")
            os.replace(tmp, target)
            replaced.append((target, existed))
        return True, f"{len(replaced)}개 파일 교체"
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
    """files(프로그램)+data_files(기대값) 다운로드 -> 검증 -> 원자 교체(.bak).
    실패 시 롤백. (성공 여부, 메시지)"""
    files = (info or {}).get("files") or {}
    data_files = (info or {}).get("data_files") or {}
    if not files and not data_files:
        return False, "업데이트 파일 목록이 비어 있습니다."
    try:
        items = [(v, _update_target_path(k, base_dir), "code") for k, v in files.items()]
        items += [(v, _data_target_path(k, base_dir), "data") for k, v in data_files.items()]
        staged = _stage_items(items, info, base_url, timeout)
    except Exception as e:
        return False, f"업데이트 실패(기존 버전 유지): {e}"
    ok, msg = commit_staged(staged)
    if ok:
        msg = (f"프로그램 {len(files)}개 + 기대값 {len(data_files)}개 파일을 "
               f"v{info.get('version')}(으)로 업데이트했습니다.")
    return ok, msg


def data_files_to_sync(info, base_dir=None):
    """로컬에 없거나 sha256이 다른 기대값 [(repo_rel, target)]."""
    out = []
    sha = (info or {}).get("sha256") or {}
    for rel_key, repo_rel in ((info or {}).get("data_files") or {}).items():
        try:
            target = _data_target_path(rel_key, base_dir)
        except RuntimeError:
            continue
        expected = sha.get(repo_rel)
        if not os.path.isfile(target):
            out.append((repo_rel, target))
        elif expected:
            try:
                with open(target, "rb") as f:
                    local = hashlib.sha256(f.read()).hexdigest()
            except OSError:
                local = ""
            if local.lower() != str(expected).lower():
                out.append((repo_rel, target))
    return out


def sync_data_files(info, base_url=UPDATE_BASE_URL, base_dir=None, timeout=15):
    """버전이 같아도 기대값만 갱신 (원격에서 사라진 파일은 지우지 않음).
    (성공, 개수, 메시지)"""
    todo = data_files_to_sync(info, base_dir)
    if not todo:
        return True, 0, "기대값 최신"
    try:
        staged = _stage_items([(r, t, "data") for r, t in todo], info or {},
                              base_url, timeout)
    except Exception as e:
        return False, 0, f"기대값 갱신 실패(기존 파일 유지): {e}"
    ok, msg = commit_staged(staged)
    return ok, (len(staged) if ok else 0), msg


def _record_installed(version):
    try:
        os.makedirs(os.path.dirname(INSTALLED_VERSION_PATH), exist_ok=True)
        with open(INSTALLED_VERSION_PATH, "w", encoding="utf-8") as f:
            f.write(str(version))
    except OSError:
        pass


def run_auto_update(base_url=UPDATE_BASE_URL, base_dir=None, timeout=3):
    """묻지 않는 자동 업데이트. 반환: (상태, 안내문).
    상태: disabled / offline / latest / applied / data / failed"""
    if not auto_update_enabled(base_dir):
        _log("자동 업데이트 꺼짐(세트설정 _설정.자동업데이트=false) — 건너뜀")
        return "disabled", ""
    info = fetch_update_info(base_url, timeout=timeout)
    if info is None:
        _log("업데이트 확인: 서버 응답 없음(오프라인/차단) — 건너뜀")
        return "offline", ""
    cur = current_version(base_dir)
    remote = str(info.get("version"))
    if update_available(info, cur):
        _log(f"새 버전 v{remote} 발견 (현재 v{cur}) — 내려받는 중")
        ok, msg = apply_update(info, base_url, base_dir)
        _log(("업데이트 적용: " if ok else "") + msg)
        if ok:
            _record_installed(remote)
            return "applied", f"v{remote}(으)로 자동 업데이트됨 — {info.get('notes') or ''}"
        return "failed", msg
    ok, n, msg = sync_data_files(info, base_url, base_dir)
    if not ok:
        _log(msg)
        return "failed", msg
    if n:
        _log(f"기대값 자동 갱신: {n}개 → {expected_values_dir(base_dir)}")
        return "data", f"기대값 파일 {n}개 자동 갱신 ({expected_values_dir(base_dir)})"
    _log(f"업데이트 확인: 최신 (v{cur})")
    return "latest", ""


def find_grade_py():
    """grade.py 탐색: 같은 폴더 -> ../채점/ 순."""
    candidates = [
        os.path.join(BASE_DIR, "grade.py"),
        os.path.join(os.path.dirname(BASE_DIR), "채점", "grade.py"),
    ]
    for c in candidates:
        if os.path.isfile(c):
            return os.path.abspath(c)
    return None


def result_dir_for(student_path):
    """풀이 파일 옆 '채점결과/' 폴더 경로 (생성 포함)."""
    out_dir = os.path.join(
        os.path.dirname(os.path.abspath(student_path)), "채점결과")
    try:
        os.makedirs(out_dir, exist_ok=True)
        return out_dir
    except OSError:
        return os.path.dirname(os.path.abspath(student_path))


def run_grade(grade_py, problem, answer, student, key, html, json_out):
    """grade.py 서브프로세스 실행. (결과 dict 또는 None, stdout, stderr, rc)"""
    cmd = [sys.executable, grade_py,
           "--problem", problem, "--answer", answer, "--student", student,
           "--html", html, "--json", json_out]
    if key:
        cmd += ["--key", key]
    proc = subprocess.run(cmd, capture_output=True, text=True,
                          encoding="utf-8", errors="replace", timeout=600)
    result = None
    if os.path.isfile(json_out):
        try:
            with open(json_out, encoding="utf-8") as f:
                result = json.load(f)
        except Exception:
            result = None
    return result, proc.stdout, proc.stderr, proc.returncode


def open_file(path):
    path = os.path.abspath(path)
    try:
        if sys.platform == "win32":
            os.startfile(path)  # noqa
            return True
        for opener in ("xdg-open", "open"):
            if shutil.which(opener):
                subprocess.Popen([opener, path], stdout=subprocess.DEVNULL,
                                 stderr=subprocess.DEVNULL)
                return True
        webbrowser.open("file:///" + path.replace(os.sep, "/"))
        return True
    except Exception:
        return False


def main():
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox
    except ImportError:
        _show_message(
            APP,
            "이 Python에는 tkinter(화면 라이브러리)가 없습니다.\n\n"
            "https://www.python.org/downloads/ 의 설치판으로 다시 설치하고\n"
            "설치 옵션에서 'tcl/tk and IDLE'을 포함해 주세요(기본값).")
        return 1

    grade_py = find_grade_py()
    if not grade_py:
        _show_message(APP, "grade.py를 찾을 수 없습니다.\n"
                           "이 파일과 같은 폴더에 grade.py가 있어야 합니다.")
        return 1

    root = tk.Tk()
    root.withdraw()

    # 완전 자동 업데이트 (묻지 않음, 실패/오프라인 시 조용히 스킵 + 로그).
    # grade.py는 서브프로세스로 실행하므로 갱신된 채점기가 바로 쓰입니다.
    _log(f"시작 · python={sys.executable} · grade.py={grade_py}")
    update_note = ""
    try:
        state, note = run_auto_update()
        if state in ("applied", "data"):
            update_note = note
            grade_py = find_grade_py() or grade_py
    except Exception as e:
        _log(f"자동 업데이트 예외: {e}")

    def install_with_progress():
        prog = tk.Toplevel(root)
        prog.title(APP)
        prog.attributes("-topmost", True)
        tk.Label(prog, text="openpyxl 설치 중입니다... 잠시 기다려 주세요.",
                 padx=26, pady=18).pack()
        prog.update()
        ok, log = install_openpyxl()
        prog.destroy()
        return ok, log

    # openpyxl 사전 확인 (grade.py는 이 인터프리터로 실행됨)
    if not openpyxl_available():
        if messagebox.askyesno(
                APP, "채점에 필요한 openpyxl이 없습니다.\n"
                "지금 설치할까요? (인터넷 필요)"):
            ok, log = install_with_progress()
            if not ok:
                messagebox.showerror(
                    APP, "openpyxl 설치에 실패했습니다.\n\n"
                    + MANUAL_PIP + "\n\n[pip 출력 요약]\n" + log)
                root.destroy()
                return 1
        else:
            messagebox.showinfo(
                APP, "openpyxl 없이는 채점할 수 없습니다.\n\n" + MANUAL_PIP)
            root.destroy()
            return 1

    ftypes = [("Excel 파일", "*.xlsx *.xlsm"), ("모든 파일", "*.*")]
    problem = filedialog.askopenfilename(title="1/3 문제 파일 선택",
                                         filetypes=ftypes)
    if not problem:
        return 0
    answer = filedialog.askopenfilename(title="2/3 정답 파일 선택",
                                        initialdir=os.path.dirname(problem),
                                        filetypes=ftypes)
    if not answer:
        return 0
    student = filedialog.askopenfilename(title="3/3 내 풀이 파일 선택",
                                         initialdir=os.path.dirname(problem),
                                         filetypes=ftypes)
    if not student:
        return 0

    key, key_origin = detect_key(problem)
    _log(f"채점: 문제={problem} · 기대값={key or '없음'}"
         + (f" ({key_origin})" if key_origin else ""))
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_dir = result_dir_for(student)
    base = os.path.join(out_dir, f"채점결과_{derive_set_name(problem)}_{stamp}")
    html = base + ".html"
    json_out = base + ".json"

    result, out, err, rc = run_grade(grade_py, problem, answer, student,
                                     key, html, json_out)
    if result is None:
        combined = ((err or "") + "\n" + (out or "")).strip()
        kind = classify_error(rc, combined)
        if kind == "module" and messagebox.askyesno(
                APP, "채점에 필요한 openpyxl이 없습니다.\n"
                "지금 설치할까요? (인터넷 필요)"):
            ok, log = install_with_progress()
            if ok:  # 설치 성공 -> 자동 재시도
                result, out, err, rc = run_grade(
                    grade_py, problem, answer, student, key, html, json_out)
                combined = ((err or "") + "\n" + (out or "")).strip()
                kind = classify_error(rc, combined)
            else:
                messagebox.showerror(
                    APP, "openpyxl 설치에 실패했습니다.\n\n"
                    + MANUAL_PIP + "\n\n[pip 출력 요약]\n" + log)
                root.destroy()
                return 1
    if result is None:
        if kind == "module":
            headline = ("채점에 필요한 openpyxl이 설치되어 있지 않아 "
                        "채점하지 못했습니다.\n\n" + MANUAL_PIP)
        elif kind == "file":
            headline = ("채점용 파일을 열 수 없습니다.\n"
                        "문제/정답/풀이 파일의 경로와 형식(.xlsx/.xlsm)이 "
                        "올바른지 확인하세요.")
        else:
            headline = "채점에 실패했습니다."
        detail = combined
        if len(detail) > 1500:
            detail = detail[:1500] + "\n..."
        messagebox.showerror(APP, headline + ("\n\n" + detail if detail else ""))
        root.destroy()
        return 1

    total = result.get("total", 0)
    verdict = "합격권" if total >= result.get("pass_line", PASS_LINE) else "미달"
    # 결과 JSON을 클립보드에 복사 — 루틴 웹 원클릭 연동 (실패는 조용히 무시)
    copied = False
    try:
        root.clipboard_clear()
        root.clipboard_append(json.dumps(result, ensure_ascii=False,
                                         separators=(",", ":")))
        root.update_idletasks()
        copied = True
    except Exception:
        pass
    lines = [f"총점: {total} / 100점", f"합격선 {PASS_LINE}점 기준: {verdict}", ""]
    for s in result.get("sheets", []):
        lines.append(f"  {s.get('name', '?')}: "
                     f"{s.get('earned', 0):g} / {s.get('alloc', 0):g}점")
    if copied:
        lines += ["", "성적이 클립보드에 복사되었습니다 — 루틴 웹페이지에서 "
                      "[성적 붙여넣기]를 누르면 자동 기록됩니다."]
    lines += ["", f"저장 폴더: {out_dir}",
              "", "확인을 누르면 상세 HTML 리포트가 열립니다."]
    if key:
        lines.insert(2, f"(기대값 JSON 자동 적용: {os.path.basename(key)}"
                        + (f" — {key_origin}" if key_origin else "") + ")")
    if update_note:
        lines += ["", update_note]
    messagebox.showinfo(f"{APP} - 결과", "\n".join(lines))
    if os.path.isfile(html):
        open_file(html)
    if messagebox.askyesno(APP, "채점결과 저장 폴더를 열까요?"):
        open_file(out_dir)
    root.destroy()
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except Exception:
        _show_message(f"{APP} - 오류",
                      "실행 중 오류가 발생했습니다.\n\n" + traceback.format_exc())
        sys.exit(1)
