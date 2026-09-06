# -*- coding: utf-8 -*-
"""
코코 시험장 더블클릭 실행 래퍼 (.pyw)

스마트 앱 컨트롤이 .bat 실행을 차단하는 환경을 위한 실행 파일입니다.
.pyw는 Windows에서 pythonw(콘솔 없는 파이썬)로 연결되어 더블클릭만으로
시험장이 열립니다. 콘솔이 없으면 오류가 보이지 않으므로:

  1. 모든 예외(ImportError·SyntaxError 포함)를 채점결과/시험장_시작로그.txt에
     기록하고 메시지 창(tkinter → ctypes MessageBoxW 순)으로 traceback 표시
  2. 시험장.py 실행 전 py_compile로 문법 검사 — 실패하면(예: 업데이트가
     중간에 깨짐) 시험장.py.bak가 있을 때 [복구] 버튼으로 되돌리기 제안
"""

import importlib.util
__version__ = "1.8.0"
# 자동 업데이트 확인은 이 파일이 실행하는 시험장.py가 수행합니다.
# (이 래퍼 자체도 시험장.py의 업데이트 대상 파일에 포함되어 함께 갱신됩니다)

import os
import platform
import py_compile
import runpy
import shutil
import subprocess
import sys
import tempfile
import traceback
from datetime import datetime

APP = "코코 시험장"
MANUAL_PIP = ("수동 설치: 명령 프롬프트에서  py -m pip install openpyxl\n"
              "(위 명령이 안 되면: python -m pip install openpyxl)")
BASE = os.path.dirname(os.path.abspath(__file__))
STARTUP_LOG = os.path.join(BASE, "채점결과", "시험장_시작로그.txt")
STARTUP_LOG_KEEP = 200
SMOKE = "--smoke" in sys.argv   # 자동 테스트: 대화상자 없이 실행


def _startup_log(message, path=None, keep=STARTUP_LOG_KEEP):
    """시작 로그 1줄 append (시험장.py와 같은 파일·형식, 최근 keep줄 유지)."""
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = (f"[{stamp}] 래퍼 {__version__} "
            + str(message).replace("\n", " | "))
    p = path or STARTUP_LOG
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
            _startup_log(f"openpyxl 설치 {'성공' if ok else '실패'}")
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


def _show_error(title, msg, detail="", recover=None):
    """콘솔이 없어도 보이는 오류 표시: tkinter(상세 텍스트·복사·복구 버튼)
    → ctypes MessageBoxW → stderr 순. recover=(버튼 문구, 콜백)이면 [복구]
    버튼을 붙이고, 눌렀을 때 콜백 결과(True/False)를 반환합니다."""
    result = {"recovered": False}
    if SMOKE:
        try:
            sys.__stderr__.write(f"{title}: {msg}\n{detail}\n")
        except Exception:
            pass
        return False
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.title(title)
        root.attributes("-topmost", True)
        frm = tk.Frame(root, padx=16, pady=14)
        frm.pack(fill="both", expand=True)
        tk.Label(frm, text=msg, justify="left", wraplength=560,
                 font=("Malgun Gothic", 10)).pack(anchor="w")
        body = (detail or "").strip()
        text = None
        if body:
            text = tk.Text(frm, height=14, width=76, font=("Consolas", 9),
                           wrap="word")
            text.insert("1.0", body)
            text.configure(state="disabled")
            text.pack(fill="both", expand=True, pady=(10, 0))
        btns = tk.Frame(frm)
        btns.pack(fill="x", pady=(10, 0))

        def copy_all():
            try:
                root.clipboard_clear()
                root.clipboard_append(f"[{APP} 시작 오류]\n{msg}\n\n{body}")
                root.update_idletasks()
                copy_btn.configure(text="복사됨")
            except Exception:
                pass

        copy_btn = tk.Button(btns, text="오류 내용 복사", command=copy_all,
                             font=("Malgun Gothic", 10), relief="groove")
        copy_btn.pack(side="left")
        if recover:
            label, callback = recover

            def do_recover():
                try:
                    result["recovered"] = bool(callback())
                except Exception as e:
                    result["recovered"] = False
                    messagebox.showerror(title, f"복구 실패: {e}")
                root.destroy()

            tk.Button(btns, text=label, command=do_recover,
                      font=("Malgun Gothic", 10, "bold"), bg="#107C41",
                      fg="white", relief="flat", padx=10).pack(
                side="left", padx=8)
        tk.Button(btns, text="닫기", command=root.destroy,
                  font=("Malgun Gothic", 10), relief="groove").pack(
            side="right")
        root.mainloop()
        return result["recovered"]
    except Exception:
        pass
    try:
        import ctypes
        full = str(msg) + ("\n\n" + str(detail)[-1500:] if detail else "")
        ctypes.windll.user32.MessageBoxW(None, full, str(title), 0x10)
        return False
    except Exception:
        pass
    try:
        if sys.__stderr__:
            sys.__stderr__.write(f"{title}: {msg}\n{detail}\n")
    except Exception:
        pass
    return False


def _syntax_check(path):
    """py_compile로 문법 검사. (통과 여부, 오류 텍스트) 반환. .pyc는 임시
    폴더에 만들고 바로 지웁니다."""
    tmpdir = tempfile.mkdtemp(prefix="coco_pyc_")
    cfile = os.path.join(tmpdir, "check.pyc")
    try:
        py_compile.compile(path, cfile=cfile, doraise=True)
        return True, ""
    except py_compile.PyCompileError as e:
        return False, str(e)
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def _recover_from_backup(target, backup=None):
    """깨진 target을 target.broken으로 옮기고 backup(.bak)을 복사해 복구.
    (성공 여부, 메시지) 반환. 백업도 문법 검사를 통과해야 합니다."""
    backup = backup or target + ".bak"
    if not os.path.isfile(backup):
        return False, f"백업 파일이 없습니다: {backup}"
    ok, err = _syntax_check(backup)
    if not ok:
        return False, f"백업 파일도 문법 오류가 있어 복구할 수 없습니다: {err}"
    try:
        broken = target + ".broken"
        if os.path.isfile(target):
            os.replace(target, broken)
        shutil.copy2(backup, target)
    except OSError as e:
        return False, f"파일 교체 실패: {e}"
    return True, f"{os.path.basename(backup)} → {os.path.basename(target)} 복구"


def _tk_missing_message():
    return ("이 Python에는 tkinter(화면 라이브러리)가 없습니다.\n\n"
            "Microsoft 스토어 경량판 등 일부 Python에는 tkinter가 빠져 "
            "있습니다.\nhttps://www.python.org/downloads/ 의 설치판으로 "
            "다시 설치하고,\n설치 옵션에서 'tcl/tk and IDLE'을 포함해 "
            "주세요(기본값).")


def main():
    target = os.path.join(BASE, "시험장.py")
    try:
        plat = platform.platform()
    except Exception:
        plat = sys.platform
    _startup_log(f"래퍼 시작 v{__version__} · python={sys.executable} · "
                 f"{plat} · 옵션={sys.argv[1:]}")
    if not os.path.isfile(target):
        _startup_log(f"시험장.py 없음: {target}")
        _show_error(APP, "시험장.py를 찾을 수 없습니다.\n"
                         f"이 파일과 같은 폴더에 있어야 합니다:\n{target}")
        return 1
    try:
        import tkinter  # noqa: F401
    except ImportError:
        _startup_log("tkinter 없음 — 종료")
        _show_error(APP, _tk_missing_message())
        return 1
    # 시험장.py 문법 검사 — 업데이트가 중간에 깨졌을 때 조용히 죽지 않도록
    ok, err = _syntax_check(target)
    if not ok:
        _startup_log(f"시험장.py 문법 오류: {err}")
        bak = target + ".bak"
        recover = None
        hint = ""
        if os.path.isfile(bak) and _syntax_check(bak)[0]:
            hint = ("\n\n이전 버전 백업(시험장.py.bak)이 있습니다. [복구]를 "
                    "누르면 백업으로 되돌린 뒤 바로 실행합니다.")
            recover = ("복구 (백업으로 되돌리기)",
                       lambda: _recover_and_log(target, bak))
        recovered = _show_error(
            f"{APP} - 시작 오류",
            "시험장.py에 문법 오류가 있어 실행할 수 없습니다.\n"
            "(업데이트 파일이 손상되었을 수 있습니다)" + hint, err,
            recover=recover)
        if not recovered:
            return 1
        ok2, err2 = _syntax_check(target)
        if not ok2:
            _startup_log(f"복구 후에도 문법 오류: {err2}")
            _show_error(f"{APP} - 시작 오류",
                        "복구 후에도 문법 오류가 남아 있습니다.", err2)
            return 1
    if not SMOKE:  # 스모크 테스트는 대화상자 없이 실행
        _ensure_openpyxl()
    # 작업 디렉터리와 무관하게 파일 위치 기준으로 실행
    if BASE not in sys.path:
        sys.path.insert(0, BASE)
    _startup_log("시험장.py 실행")
    try:
        runpy.run_path(target, run_name="__main__")
    except SystemExit:
        raise
    except KeyboardInterrupt:
        _startup_log("사용자 중단(KeyboardInterrupt)")
        return 130
    except BaseException:
        tb = traceback.format_exc()
        last = tb.strip().splitlines()[-1] if tb.strip() else ""
        _startup_log(f"시험장.py 예외: {last}")
        _write_error_log(tb)
        _show_error(
            f"{APP} - 오류",
            "시험장 실행 중 오류가 발생했습니다.\n"
            f"오류 내용은 {STARTUP_LOG}\n(및 시험장_오류.log)에 기록되었습니다. "
            "[오류 내용 복사]로 복사해 채팅에 붙여넣어 주세요.", tb)
        return 1
    return 0


def _recover_and_log(target, bak):
    ok, msg = _recover_from_backup(target, bak)
    _startup_log(f"복구 {'성공' if ok else '실패'}: {msg}")
    if not ok:
        raise RuntimeError(msg)
    return True


def _write_error_log(tb):
    """시험장.py의 오류 로그(채점결과/시험장_오류.log)에도 traceback 기록."""
    try:
        p = os.path.join(BASE, "채점결과", "시험장_오류.log")
        os.makedirs(os.path.dirname(p), exist_ok=True)
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(p, "a", encoding="utf-8") as f:
            f.write(f"[{stamp}] 래퍼 {__version__} 시험장.py 실행\n{tb}"
                    + "-" * 60 + "\n")
    except Exception:
        pass


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except BaseException:
        tb = traceback.format_exc()
        _startup_log("래퍼 예외: "
                     + (tb.strip().splitlines()[-1] if tb.strip() else ""))
        _write_error_log(tb)
        if SMOKE:  # 자동 테스트: 대화상자 대신 traceback
            traceback.print_exc()
            sys.exit(1)
        _show_error(f"{APP} - 오류",
                    "실행 중 오류가 발생했습니다.\n\n" + tb.strip().splitlines()[-1],
                    tb)
        sys.exit(1)
