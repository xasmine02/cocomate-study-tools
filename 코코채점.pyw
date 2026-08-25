# -*- coding: utf-8 -*-
"""
코코 채점 더블클릭 실행 도구 (.pyw)

스마트 앱 컨트롤이 .bat 실행을 차단하는 환경용입니다. 더블클릭하면
파일 선택 창 3개(문제 -> 정답 -> 내 풀이)로 파일을 고르고, 같은 폴더의
grade.py로 채점한 뒤 점수를 표시하고 HTML 리포트를 엽니다.
콘솔이 없어도 모든 오류를 메시지 창으로 표시합니다.
"""

__version__ = "1.0.0"

import importlib.util
import json
import os
import re
import shutil
import subprocess
import sys
import time
import traceback
import webbrowser
from datetime import datetime

APP = "코코 채점"
PASS_LINE = 70
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPDATE_BASE_URL = ("https://raw.githubusercontent.com/"
                   "xasmine02/cocomate-study-tools/main/")
UPDATE_STAMP = os.path.join(BASE_DIR, "채점결과", ".update_check")
MANUAL_PIP = ("수동 설치: 명령 프롬프트에서  py -m pip install openpyxl\n"
              "(위 명령이 안 되면: python -m pip install openpyxl)")


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


# ---------------- 자동 업데이트 (공개 저장소) ----------------


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
    """업데이트 확인 하루 1회 제한."""
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
    """다운로드 -> 검증(__version__ + 문법) -> 원자 교체(.bak). 실패 시 롤백."""
    import py_compile
    import urllib.parse
    import urllib.request
    files = (info or {}).get("files") or {}
    if not files:
        return False, "업데이트 파일 목록이 비어 있습니다."
    staged = []
    replaced = []
    try:
        for rel_key, repo_rel in files.items():
            target = _update_target_path(rel_key, base_dir)
            url = base_url + urllib.parse.quote(str(repo_rel))
            req = urllib.request.Request(
                url, headers={"User-Agent": "cocomate-updater"})
            with urllib.request.urlopen(req, timeout=timeout) as r:
                data = r.read()
            if "__version__" not in data.decode("utf-8"):
                raise RuntimeError(f"{repo_rel}: __version__ 표식 없음")
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
        for _t, tmp in staged:
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


def detect_key(problem_path):
    """문제 파일 옆의 <세트명>_기대값.json 자동 감지."""
    base = os.path.basename(problem_path)
    m = re.match(r"^(.+)_문제\.(xlsx|xlsm)$", base, re.IGNORECASE)
    if not m:
        return None
    key = os.path.join(os.path.dirname(problem_path),
                       f"{m.group(1)}_기대값.json")
    return key if os.path.isfile(key) else None


def derive_set_name(problem_path):
    """문제 파일명에서 세트명 추출 (…_문제.xlsx -> …)."""
    base = os.path.splitext(os.path.basename(str(problem_path)))[0]
    for suf in ("_문제", "-문제", " 문제"):
        if base.endswith(suf):
            return base[: -len(suf)]
    return base


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

    # 자동 업데이트 확인 (하루 1회, 실패/오프라인 시 조용히 스킵)
    if should_check_update():
        info = fetch_update_info()
        if update_available(info):
            if messagebox.askyesno(
                    APP, f"새 버전 v{info.get('version')}가 있습니다 — 지금 "
                    f"업데이트할까요?\n변경: {info.get('notes') or '-'}"):
                ok, msg = apply_update(info)
                if ok:
                    messagebox.showinfo(
                        APP, msg + "\n\n프로그램을 다시 실행해 주세요.")
                    root.destroy()
                    return 0
                messagebox.showwarning(APP, msg)

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

    key = detect_key(problem)
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
    lines = [f"총점: {total} / 100점", f"합격선 {PASS_LINE}점 기준: {verdict}", ""]
    for s in result.get("sheets", []):
        lines.append(f"  {s.get('name', '?')}: "
                     f"{s.get('earned', 0):g} / {s.get('alloc', 0):g}점")
    lines += ["", f"저장 폴더: {out_dir}",
              "", "확인을 누르면 상세 HTML 리포트가 열립니다."]
    if key:
        lines.insert(2, f"(기대값 JSON 자동 적용: {os.path.basename(key)})")
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
