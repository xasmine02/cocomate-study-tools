# -*- coding: utf-8 -*-
"""시험장.py v2.2.1 / 코코시험장.pyw v1.8.0 시작 안정화 헤드리스 테스트.

  1. 기록.json 혼합 픽스처(v1.0~v2.2 + 손상 항목)를 읽어도 절대 죽지 않음
     — 점수 문자열 정규화, dict 형태 파일, BOM, 깨진 JSON 백업
  2. 세트설정.json 옛/깨진 스키마 방어 (리스트·문자열·잘못된 타입)
  3. 시작 로그: 최근 200줄 유지, 오류 로그 한 줄 요약 미러
  4. 재시작 명령: 래퍼(.pyw) 우선, --smoke 제거, Windows 콘솔 플래그
  5. 래퍼(코코시험장.pyw) 헬퍼: py_compile 문법 검사, .bak 복구, 시작 로그
"""
import importlib.machinery
import importlib.util
import json
import os
import re
import sys
import tempfile
from datetime import date

BASE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("sj", os.path.join(BASE, "시험장.py"))
sj = importlib.util.module_from_spec(spec)
spec.loader.exec_module(sj)
N = 0
TMP = tempfile.mkdtemp(prefix="coco_start_")
sj.STARTUP_LOG_PATH = os.path.join(TMP, "시작로그.txt")
sj.ERROR_LOG_PATH = os.path.join(TMP, "오류.log")


def check(desc, cond, extra=""):
    global N
    N += 1
    print(f"  [{N:02d}] {desc}: {'OK' if cond else 'FAIL'}" + (f" ({extra})" if extra else ""))
    if not cond:
        raise AssertionError(f"{desc} {extra}")


def write_json(name, data, bom=False, raw=None):
    p = os.path.join(TMP, name)
    with open(p, "w", encoding="utf-8-sig" if bom else "utf-8") as f:
        if raw is not None:
            f.write(raw)
        else:
            json.dump(data, f, ensure_ascii=False, indent=2)
    return p


def mkset(name):
    return {"name": name, "norm": sj.norm_set_key(name), "dir": TMP,
            "problem": os.path.join(TMP, name + "_문제.xlsx"),
            "answer": os.path.join(TMP, name + "_정답.xlsm"), "key": None, "pdf": None}


check("버전 표식", re.fullmatch(r"\d+\.\d+\.\d+", sj.__version__) is not None
      and tuple(int(x) for x in sj.__version__.split(".")) >= (2, 2, 1), sj.__version__)

print("1. 기록.json 혼합 픽스처")
MIXED = [
    # v1.0 (3b2a6a2): 일시/세트명/점수/소요시간/리포트 — mode 없음
    {"일시": "2026-08-25 10:00", "세트명": "2024 A형", "점수": 72, "소요시간": "38분 10초", "리포트": None},
    {"일시": "2026-08-25 11:00", "세트명": "2024 A형 (연습)", "점수": None, "소요시간": "5분 00초", "리포트": None},
    # v1.4 (faecc32): mode 부분연습 + 영역 + 만점
    {"일시": "2026-08-27 10:00", "세트명": "코코모의고사1회", "점수": 18, "소요시간": "14분 02초",
     "리포트": None, "mode": "부분연습", "영역": "계산작업", "만점": 20},
    # v1.7 (8ca3008): day 추가
    {"일시": "2026-09-03 09:30", "세트명": "2024년 상시1회 2급", "점수": 65, "소요시간": "40분 00초",
     "리포트": "x.html", "mode": "시험", "day": "d01"},
    # v2.1 (a6984d9): 오답재풀이
    {"일시": "2026-09-03 11:00", "세트명": "2024년 상시1회 2급", "점수": 10, "소요시간": "12분 00초",
     "리포트": None, "mode": "오답재풀이", "영역": "오답재풀이(계산작업)", "만점": 15, "day": "d01"},
    # v2.2 (fe6ca6b): 루틴 태그
    {"일시": "2026-09-04 10:00", "세트명": "코코모의고사1회", "점수": 81, "소요시간": "39분 00초",
     "리포트": None, "mode": "시험", "day": "d02", "루틴": "2026-09"},
    # 손상/변형 항목
    {"일시": "2026-09-04 12:00", "세트명": "2024 B형", "점수": "85", "소요시간": 2280},        # 점수 문자열
    {"일시": "2026-09-04 13:00", "세트명": "2024 B형", "점수": "채점 실패"},                   # 숫자 아님
    {"일시": "2026-09-04 14:00", "세트명": "2024 B형", "점수": "77.0점", "만점": "100"},        # 단위 붙은 문자열
    {"일시": 20260905, "세트명": ["리스트"], "점수": True, "mode": 3},                          # 타입 엉망
    {"점수": 50},                                                                              # 세트명·일시 없음
    "문자열 항목", None, 42, ["리스트", "항목"],                                               # dict 아님
]
p = write_json("기록_혼합.json", MIXED)
recs = sj.load_records(p)
check("dict 항목 11건만 살림 (비-dict 4건 제외)", len(recs) == 11, str(len(recs)))
check("점수 문자열 '85' → 85, '77.0점' → 77, '채점 실패' → None, True → None",
      [r["점수"] for r in recs if r["세트명"] == "2024 B형"] == [85, None, 77]
      and recs[9]["점수"] is None)
check("만점 문자열 '100' → 100", recs[8]["만점"] == 100)
check("타입 엉망 항목도 문자열로 정리", recs[9]["세트명"] == "['리스트']" and recs[9]["일시"] == "20260905"
      and recs[9]["mode"] == "3")
check("세트명·일시 없는 항목은 '?'", recs[10]["세트명"] == "?" and recs[10]["일시"] == "?" and recs[10]["점수"] == 50)
check("v1.0 기록(mode 없음)은 시험 모드로 집계", sj.records_summary(recs)["2024 A형"]["best"] == 72)
check("set_records/set_exam_records 동작", len(sj.set_records("2024년 상시1회 2급", recs)) == 2
      and len(sj.set_exam_records("2024년 상시1회 2급", recs)) == 1)
SETS = [mkset(n) for n in ("2024년 상시1회 2급", "코코모의고사1회", "2024 A형", "2024년 상시2회 2급", "2024 B형")]
EXAM_D = date(2026, 9, 10)        # 9/5 = D-5 (시험일은 사용자가 입력하는 값)
plan = sj.build_adaptive_plan(date(2026, 9, 5), recs, SETS, [EXAM_D])
check("적응형 일정 계산이 혼합 기록으로도 예외 없음",
      plan["kind"] == "week" and plan["days"])
check("시험일 미설정이어도 일정 계산이 죽지 않는다 (kind unset · 배정 없음)",
      sj.build_adaptive_plan(date(2026, 9, 5), recs, SETS, [])["kind"] == "unset")
picks = sj.pick_set_for_retry(SETS, recs, 1, date(2026, 9, 5))
check("자동 세트 선택 예외 없음", picks and picks[0][0]["name"], str(picks))
by = sj.exam_records_by_set(SETS, recs)
check("세트별 시험 기록 집계 (날짜 없는 항목 제외)", len(by[sj.norm_set_key("2024 B형")]) == 3
      and all(sj._record_date(r) for rs in by.values() for r in rs))
# dict 형태 파일
p2 = write_json("기록_dict.json", {"기록": MIXED[:3]})
check("{'기록': [...]} 형태", len(sj.load_records(p2)) == 3)
p3 = write_json("기록_dict2.json", {"2024 A형": {"점수": 60, "일시": "2026-09-01 10:00"},
                                    "2024 B형": [{"점수": "70", "일시": "2026-09-02 10:00"}], "_설정": 1})
r3 = sj.load_records(p3)
check("{세트명: 기록} 형태 → 키를 세트명으로", [(r["세트명"], r["점수"]) for r in r3] == [("2024 A형", 60), ("?", 70)])
p4 = write_json("기록_bom.json", MIXED[:2], bom=True)
check("BOM 있는 파일도 읽음", len(sj.load_records(p4)) == 2)
p5 = write_json("기록_str.json", "그냥 문자열")
check("최상위가 문자열이면 빈 목록", sj.load_records(p5) == [])
p6 = write_json("기록_깨짐.json", None, raw="[{\"일시\": \"2026-09-01\", ")
check("깨진 JSON → 빈 목록(예외 없음)", sj.load_records(p6) == [])
sj.append_record({"일시": "2026-09-06 10:00", "세트명": "새 기록", "점수": 90}, p6)
backups = [f for f in os.listdir(TMP) if f.startswith("기록_깨짐.json.corrupt-")]
check("깨진 JSON에 append 시 .corrupt-* 백업 후 새로 작성", len(backups) == 1 and len(sj.load_records(p6)) == 1)
check("normalize_records 잡동사니 입력", sj.normalize_records(None) == [] and sj.normalize_records(5) == []
      and sj.normalize_records({"a": 1}) == [])

print("2. 세트설정.json 옛/깨진 스키마")
cfgp = write_json("세트설정_list.json", [1, 2, 3])
check("최상위 리스트 → {}", sj.load_set_config(cfgp) == {})
weird = {
    "_진행": [1, 2], "_슬롯매핑": "x", "_자동선택": {"d08": "abc", "d09": [{"세트": "코코모의고사1회"}, "2024 A형", 7]},
    "_설정": 5, "_기타": [1],
    "세트a": "문자열", "세트b": {"pdf": ["list"], "key": 3, "problem": 1, "answer": None},
    "정상": {"오답연습": [1], "name": "정상 세트"},
}
cfgp = write_json("세트설정_weird.json", weird)
cfg = sj.load_set_config(cfgp)
check("dict 아닌 하위 항목 제거, 그 외 유지", "_진행" not in cfg and "_슬롯매핑" not in cfg and "_설정" not in cfg
      and "세트a" not in cfg and "세트b" in cfg and cfg["_기타"] == [1])
check("load_step_progress → 빈 집합", sj.load_step_progress("d03", cfgp) == set())
check("load_auto_picks: 문자열 → [], 혼합 리스트 정리",
      sj.load_auto_picks("d08", cfgp) == [] and sj.load_auto_picks("d09", cfgp)
      == [{"세트": "코코모의고사1회", "이유": ""}, {"세트": "2024 A형", "이유": ""}, None])
check("get_app_setting 기본값", sj.get_app_setting("신뢰위치등록", False, cfgp) is False)
check("load_slot_mapping → {}", sj.load_slot_mapping(cfgp) == {})
check("apply_set_config: 잘못된 타입 항목 무시", sj.apply_set_config([], cfg) == [])
check("load_review_state: 리스트 '오답연습' → 빈 집합", sj.load_review_state("정상", "j.json", cfgp) == set())
check("record_review_state: 리스트 '오답연습'을 dict로 교체하고 저장",
      sj.record_review_state("정상", "j.json", {1, 3}, cfgp, count_up=True)
      and sj.load_review_state("정상", "j.json", cfgp) == {1, 3})
check("save_step_progress/load 왕복", sj.save_step_progress("d03", {0, 2}, cfgp)
      and sj.load_step_progress("d03", cfgp) == {0, 2})
check("save_slot_mapping 왕복 (문자열이던 '_슬롯매핑' 교체)",
      sj.save_slot_mapping("2024 A형", mkset("2024 A형"), cfgp)
      and sj.load_slot_mapping(cfgp)["2024 A형"]["name"] == "2024 A형")
check("set_app_setting/get_app_setting 왕복 (숫자였던 '_설정' 교체)",
      sj.set_app_setting("신뢰위치등록", True, cfgp) and sj.get_app_setting("신뢰위치등록", path=cfgp) is True)
check("remember_set: 문자열이던 세트 항목 재생성",
      sj.remember_set(dict(mkset("세트a"), norm="세트a"), cfgp) and sj.load_set_config(cfgp)["세트a"]["name"] == "세트a")
cfgp2 = write_json("세트설정_진행dict.json", {"_진행": {sj._progress_key("d05"): {"1": 1, "x": 2}}})
check("'_진행' 값이 dict면 키를 번호로", sj.load_step_progress("d05", cfgp2) == {1})
cfgp3 = write_json("세트설정_깨짐.json", None, raw="{\"_진행\": ")
check("깨진 세트설정.json → {}", sj.load_set_config(cfgp3) == {})
check("normalize_set_config 잡동사니", sj.normalize_set_config(None) == {} and sj.normalize_set_config([]) == {})

print("3. 시작 로그")
lp = os.path.join(TMP, "로그_trim.txt")
for i in range(250):
    sj.startup_log(f"줄 {i}\n둘째 줄", path=lp, keep=200)
lines = open(lp, encoding="utf-8").read().splitlines()
check("최근 200줄만 유지, 마지막이 최신, 줄바꿈은 ' | '로",
      len(lines) == 200 and lines[-1].endswith("줄 249 | 둘째 줄") and lines[0].endswith("줄 50 | 둘째 줄"))
check("줄 형식 [일시] 버전 메시지", lines[-1].startswith("[20") and f"] {sj.__version__} " in lines[-1])
txt = sj.read_startup_log(lp)
check("read_startup_log: 머리말 + 본문", txt.startswith("[코코 시험장 시작 로그]") and lp in txt and "줄 249" in txt)
check("없는 로그는 안내문", "아직 기록이 없습니다" in sj.read_startup_log(os.path.join(TMP, "없음.txt")))
try:
    raise ValueError("미러 확인")
except ValueError as e:
    sj.log_error("단위 테스트 문맥", e)
slog = open(sj.STARTUP_LOG_PATH, encoding="utf-8").read()
check("log_error(path=None)는 시작 로그에 한 줄 요약 미러",
      "예외: 단위 테스트 문맥 — ValueError: 미러 확인" in slog and "ValueError: 미러 확인" in open(sj.ERROR_LOG_PATH, encoding="utf-8").read())
check("startup_log는 쓰기 실패해도 예외 없음",
      isinstance(sj.startup_log("x", path=os.path.join(TMP, "없는폴더" * 60, "a", "b.txt")), str))

print("4. 재시작 명령")
args, kw = sj.restart_command(argv=["시험장.py", "--smoke", "--scan-root", "/x"], executable="/usr/bin/python3")
check("시험장.py로 시작 → 시험장.py 재실행, --smoke 제거, 옵션 유지",
      args == ["/usr/bin/python3", os.path.join(BASE, "시험장.py"), "--scan-root", "/x"] and kw["cwd"] == BASE
      and kw["close_fds"] is True, str(args))
pyw = os.path.join(BASE, "코코시험장.pyw")
args2, kw2 = sj.restart_command(argv=[pyw], executable="/usr/bin/python3")
check("래퍼(.pyw)로 시작했으면 래퍼 재실행", args2[1] == pyw, str(args2))
if sys.platform != "win32":
    check("Windows 외: start_new_session", kw2.get("start_new_session") is True)
_plat = sys.platform
try:
    sys.platform = "win32"
    # (Linux의 posixpath는 백슬래시를 구분자로 보지 않으므로 슬래시 경로로)
    a3, k3 = sj.restart_command(argv=["시험장.py"], executable="C:/Py/pythonw.exe")
    a4, k4 = sj.restart_command(argv=["시험장.py"], executable="C:/Py/python.exe")
finally:
    sys.platform = _plat
check("Windows pythonw → DETACHED_PROCESS|CREATE_NEW_PROCESS_GROUP",
      k3["creationflags"] == sj.DETACHED_PROCESS | sj.CREATE_NEW_PROCESS_GROUP and "start_new_session" not in k3)
check("Windows python.exe → CREATE_NO_WINDOW|CREATE_NEW_PROCESS_GROUP (콘솔 안 뜸)",
      k4["creationflags"] == sj.CREATE_NO_WINDOW | sj.CREATE_NEW_PROCESS_GROUP)
check("_show_fatal/_platform_text 존재", callable(sj._show_fatal) and isinstance(sj._platform_text(), str))

print("5. 래퍼(코코시험장.pyw) 헬퍼")
loader = importlib.machinery.SourceFileLoader("coco_wrapper", pyw)
wspec = importlib.util.spec_from_loader("coco_wrapper", loader)
W = importlib.util.module_from_spec(wspec)
_argv = sys.argv
sys.argv = ["코코시험장.pyw", "--smoke"]
try:
    loader.exec_module(W)
finally:
    sys.argv = _argv
check("래퍼 버전 1.8.0 + SMOKE 플래그", W.__version__ == "1.8.0" and W.SMOKE is True)
check("래퍼 시작 로그 경로 = 시험장.py와 같은 채점결과/시험장_시작로그.txt",
      W.STARTUP_LOG == os.path.join(BASE, "채점결과", "시험장_시작로그.txt"))
ok, err = W._syntax_check(os.path.join(BASE, "시험장.py"))
check("py_compile 문법 검사: 시험장.py 통과", ok and err == "")
broken = os.path.join(TMP, "시험장.py")
with open(broken, "w", encoding="utf-8") as f:
    f.write("def x(:\n    pass\n")
ok, err = W._syntax_check(broken)
check("문법 오류 감지 (SyntaxError 문구)", not ok and "SyntaxError" in err, err[:60])
ok, msg = W._recover_from_backup(broken)
check("백업 없으면 복구 실패 메시지", not ok and "백업 파일이 없습니다" in msg)
with open(broken + ".bak", "w", encoding="utf-8") as f:
    f.write("__version__ = '0.0.0'\nprint('ok')\n")
ok, msg = W._recover_from_backup(broken)
check(".bak 복구: 깨진 파일 → .broken, 백업 → 시험장.py",
      ok and os.path.isfile(broken + ".broken") and W._syntax_check(broken)[0] and "복구" in msg, msg)
with open(broken, "w", encoding="utf-8") as f:
    f.write("def y(:\n")
with open(broken + ".bak", "w", encoding="utf-8") as f:
    f.write("def z(:\n")
ok, msg = W._recover_from_backup(broken)
check("백업도 깨졌으면 복구 거부", not ok and "백업 파일도 문법 오류" in msg)
wl = os.path.join(TMP, "래퍼로그.txt")
line = W._startup_log("래퍼 테스트\n두 줄", path=wl)
check("래퍼 시작 로그 형식 '[일시] 래퍼 1.8.0 …'", " 래퍼 1.8.0 래퍼 테스트 | 두 줄" in line
      and open(wl, encoding="utf-8").read().strip() == line)
check("래퍼 _show_error는 --smoke에서 대화상자 없이 False", W._show_error("t", "m", "d") is False)
check("래퍼 _tk_missing_message/_pip_python 존재", "tcl/tk" in W._tk_missing_message() and W._pip_python())

print()
print(f"시작 안정화 테스트 {N}건 전부 통과")
