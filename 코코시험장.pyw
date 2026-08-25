# -*- coding: utf-8 -*-
"""
코코 시험장 더블클릭 실행 래퍼 (.pyw)

스마트 앱 컨트롤이 .bat 실행을 차단하는 환경을 위한 실행 파일입니다.
.pyw는 Windows에서 pythonw(콘솔 없는 파이썬)로 연결되어 더블클릭만으로
시험장이 열립니다. 콘솔이 없으면 오류가 보이지 않으므로 모든 예외를
메시지 창으로 표시합니다.
"""

import importlib.util
__version__ = "1.0.0"
# 자동 업데이트 확인은 이 파일이 실행하는 시험장.py가 수행합니다.
# (이 래퍼 자체도 시험장.py의 업데이트 대상 파일에 포함되어 함께 갱신됩니다)

import os
import runpy
import subprocess
import sys
import traceback

APP = "코코 시험장"
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


def _install_openpyxl(timeout=600):
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


def _ensure_openpyxl():
    """채점에 필요한 openpyxl이 없으면 설치를 제안 (시험 진행 자체는 막지 않음)."""
    try:
        if importlib.util.find_spec("openpyxl") is not None:
            return
    except Exception:
        return
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        if messagebox.askyesno(
                APP, "채점에 필요한 openpyxl이 없습니다.\n"
                "지금 설치할까요? (인터넷 필요)"):
            prog = tk.Toplevel(root)
            prog.title(APP)
            prog.attributes("-topmost", True)
            tk.Label(prog, text="openpyxl 설치 중입니다... 잠시 기다려 주세요.",
                     padx=26, pady=18).pack()
            prog.update()
            ok, log = _install_openpyxl()
            prog.destroy()
            if ok:
                messagebox.showinfo(APP, "openpyxl 설치가 완료되었습니다.")
            else:
                messagebox.showwarning(
                    APP, "openpyxl 설치에 실패했습니다.\n"
                    "시험은 볼 수 있지만 채점은 실패합니다.\n\n"
                    + MANUAL_PIP + "\n\n[pip 출력 요약]\n" + log)
        else:
            messagebox.showwarning(
                APP, "openpyxl이 없으면 시험은 볼 수 있지만 채점은 "
                "실패합니다.\n\n" + MANUAL_PIP)
        root.destroy()
    except Exception:
        pass


def _show_error(title, msg):
    """콘솔이 없어도 보이는 오류 표시: tkinter -> ctypes -> stderr 순."""
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(title, msg)
        root.destroy()
        return
    except Exception:
        pass
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(None, str(msg), str(title), 0x10)
        return
    except Exception:
        pass
    try:
        if sys.__stderr__:
            sys.__stderr__.write(f"{title}: {msg}\n")
    except Exception:
        pass


def main():
    base = os.path.dirname(os.path.abspath(__file__))
    target = os.path.join(base, "시험장.py")
    if not os.path.isfile(target):
        _show_error(APP, "시험장.py를 찾을 수 없습니다.\n"
                         f"이 파일과 같은 폴더에 있어야 합니다:\n{target}")
        return 1
    try:
        import tkinter  # noqa: F401
    except ImportError:
        _show_error(
            APP,
            "이 Python에는 tkinter(화면 라이브러리)가 없습니다.\n\n"
            "Microsoft 스토어 경량판 등 일부 Python에는 tkinter가 빠져 "
            "있습니다.\nhttps://www.python.org/downloads/ 의 설치판으로 "
            "다시 설치하고,\n설치 옵션에서 'tcl/tk and IDLE'을 포함해 "
            "주세요(기본값).")
        return 1
    if "--smoke" not in sys.argv:  # 스모크 테스트는 대화상자 없이 실행
        _ensure_openpyxl()
    # 작업 디렉터리와 무관하게 파일 위치 기준으로 실행
    if base not in sys.path:
        sys.path.insert(0, base)
    runpy.run_path(target, run_name="__main__")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except Exception:
        _show_error(f"{APP} - 오류",
                    "실행 중 오류가 발생했습니다.\n\n" + traceback.format_exc())
        sys.exit(1)
