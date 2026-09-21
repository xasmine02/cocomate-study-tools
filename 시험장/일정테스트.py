# -*- coding: utf-8 -*-
"""시험장.py 일정 뼈대 헤드리스 테스트 (tkinter 불필요).

실행: python 일정테스트.py   (시험장.py와 같은 폴더)
  일정은 v3.0.0부터 **사용자가 입력한 시험일 기준 1주(D-7 ~ D-1)**이고, 어떤
  세트를 푸는지는 **보유 세트**로 정해진다(고정 세트 슬롯 없음). 그래서 이
  스위트는 "특정 세트가 특정 날짜에 온다"를 더 이상 검사하지 않는다.

  1. 시험일 설정 · 날짜 -> Day 경계 (D-8→0, D-7→1, D-1→7, 시험일→8, 2차→9, 이후→10)
  2. 자동 세트 선택 규칙 5케이스 (신규 우선/최저점/최근 2일 제외/2세트 상이/기록 없음 폴백)
  3. 오답 재풀이: 오답 시트 추출 + grade.py --sheets 인자
  4. 스텝 시퀀스 무결성 (Day 1~7 전부 · 필수 ①② + 선택 ③④⑤)
  5. 카드 제목/D-day 표기, 진행 저장 키 분리(시험일마다)
  6. v2.5.0 규칙: 목표 사다리 70→80→90→100 · 연속 응시 (v2.7.0: 응시 제한 폐지)
  7. v2.6.0 게임형 요소: 티어 이름·게이지·도전 과제 12종·결과 연출
  8. v2.7.0 점수대 세트 보드: 밴드 경계값·밴드 이동·점수 이력·요약 문구
  9. v3.0.0 첫 실행 안내: 시험일 미설정 · 자동 업데이트 고지 · 세트 넣는 방법
"""
import importlib.util
import json
import os
import re
import sys
import tempfile
from datetime import date, datetime

BASE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("sj", os.path.join(BASE, "시험장.py"))
sj = importlib.util.module_from_spec(spec)
spec.loader.exec_module(sj)

TMP = tempfile.mkdtemp()
N = 0


def check(desc, cond, extra=""):
    global N
    N += 1
    print(f"  [{N:02d}] {desc}: {'OK' if cond else 'FAIL'}"
          + (f" ({extra})" if extra else ""))
    if not cond:
        raise AssertionError(f"{desc} {extra}")


def mkset(name, norm=None):
    return {"name": name, "norm": norm or name, "dir": TMP,
            "problem": os.path.join(TMP, name + "_문제.xlsx"),
            "answer": os.path.join(TMP, name + "_정답.xlsx"),
            "key": None, "pdf": None}


def rec(name, score, day, mode="시험"):
    return {"일시": f"{day.isoformat()} 10:00", "세트명": name, "점수": score,
            "mode": mode}


# 1. 시험일 설정 · 날짜 -> Day 경계 -----------------------------------------
print("1. 시험일 설정 · 날짜 -> Day 경계")
D = date
TD = __import__("datetime").timedelta
EXAM = D(2026, 11, 19)              # 테스트용 시험일 (코드에 박힌 값이 아님)
EX = [EXAM]
EX2 = [EXAM, EXAM + TD(days=1)]     # 2차 시험일까지 넣은 경우


def dd(n):
    """시험일로부터 n일 전(D-n) 날짜."""
    return EXAM - TD(days=n)


check("코드에 박힌 시험일 상수가 없다 (EXAM_DATES·ROUTINE_START·EXAM_DATE 폐지)",
      not any(hasattr(sj, n) for n in
              ("EXAM_DATES", "EXAM_DATE", "ROUTINE_START", "ROUTINE_TAG",
               "PROGRESS_NS", "PRIORITY_SPECS")),
      str([n for n in ("EXAM_DATES", "EXAM_DATE", "ROUTINE_START",
                       "ROUTINE_TAG", "PROGRESS_NS", "PRIORITY_SPECS")
           if hasattr(sj, n)]))
check("ROUTINE_DAYS = 7 · 커리큘럼은 D-7 ~ D-1", sj.ROUTINE_DAYS == 7)
CFGP = os.path.join(TMP, "시험일설정.json")
check("시험일 미설정: exam_dates [] · routine_start None · dday None",
      sj.exam_dates(path=CFGP) == [] and sj.routine_start([]) is None
      and sj.dday_num(D(2026, 11, 1), []) is None
      and sj.exam_date_set([]) is False
      and sj.dday_text(D(2026, 11, 1), []) == "시험일 미설정")
check("시험일 저장 → 설정 파일에서 다시 읽힌다 (2차는 선택)",
      sj.save_exam_dates(EXAM, None, CFGP) == [EXAM]
      and sj.exam_dates(path=CFGP) == [EXAM]
      and sj.save_exam_dates(EXAM, EXAM + TD(days=1), CFGP) == EX2
      and sj.exam_dates(path=CFGP) == EX2)
check("2차가 1차보다 앞서거나 같으면 버린다 · 1차를 지우면 둘 다 지워진다",
      sj.save_exam_dates(EXAM, EXAM, CFGP) == [EXAM]
      and sj.save_exam_dates(EXAM, EXAM - TD(days=1), CFGP) == [EXAM]
      and sj.save_exam_dates(None, None, CFGP) == []
      and sj.exam_dates(path=CFGP) == [])
check("parse_iso_date: 정상/형식 오류/None",
      sj.parse_iso_date("2026-11-19") == EXAM and sj.parse_iso_date("x") is None
      and sj.parse_iso_date(None) is None and sj.parse_iso_date(EXAM) == EXAM)

cases = [(dd(8), 0), (dd(7), 1), (dd(4), 4), (dd(1), 7),
         (EXAM, sj.PLAN_EXAM1), (EXAM + TD(days=1), sj.PLAN_AFTER),
         (EXAM + TD(days=30), sj.PLAN_AFTER), (dd(60), 0)]
for d, want in cases:
    check(f"{d.isoformat()} -> {want}", sj.routine_day_no(d, EX) == want,
          str(sj.routine_day_no(d, EX)))
check("2차 시험일을 넣으면 그 날은 9(PLAN_EXAM2)이고 그 뒤가 이후",
      sj.routine_day_no(EXAM + TD(days=1), EX2) == sj.PLAN_EXAM2
      and sj.routine_day_no(EXAM + TD(days=2), EX2) == sj.PLAN_AFTER
      and sj.routine_day_no(EXAM, EX2) == sj.PLAN_EXAM1)
check("시험일 미설정이면 어떤 날짜도 0(자유 연습)",
      all(sj.routine_day_no(d, []) == 0 for d, _w in cases))
check("Day 번호 -> 날짜 역방향 (1~7 + 시험일) 일치",
      all(sj.routine_day_no(sj.routine_date_for(no, EX2), EX2) == no
          for no in sj.PLAN_ORDER if sj.routine_date_for(no, EX2)))
check("Day 1~7 = D-7 ~ D-1 연속 (시험일은 학습일 뒤)",
      [sj.routine_date_for(n, EX) for n in range(1, 8)]
      == [dd(7 - i) for i in range(7)]
      and sj.routine_date_for(sj.PLAN_EXAM1, EX) == EXAM
      and sj.routine_date_for(sj.PLAN_EXAM2, EX) is None
      and sj.routine_date_for(sj.PLAN_EXAM2, EX2) == EXAM + TD(days=1)
      and not any(sj.routine_date_for(n, EX) in EX for n in range(1, 8)))
check("PLAN_ORDER 시간순: 0, 1~7, 시험, 2차 시험, 이후",
      sj.PLAN_ORDER == [0] + list(range(1, 8))
      + [sj.PLAN_EXAM1, sj.PLAN_EXAM2, sj.PLAN_AFTER])
check("학습 구간 도출: D-7 ~ D-1 단일 구간 (미설정이면 빈 목록)",
      [(g["seg"], g["seg_start"], g["end"], g["exam"])
       for g in sj.study_segments(EX)] == [(1, dd(7), dd(1), EXAM)]
      and sj.study_segments([]) == []
      and sj.next_segment_after(1, EX) is None)
check("2차 시험일이 있어도 학습 구간은 1차 기준 하나뿐",
      [(g["seg_start"], g["end"]) for g in sj.study_segments(EX2)]
      == [(dd(7), dd(1))])

# 2. 자동 세트 선택 5케이스 ------------------------------------------------
print("2. 자동 세트 선택 규칙")
A, B, C, Dd = (mkset("2024 A형"), mkset("2024 B형"), mkset("코코 1회"),
               mkset("코코 2회"))
SETS = [A, B, C, Dd]
today = dd(4)
# (1) 신규 우선: A/B/C 기록 있고 D 없음 -> D
R1 = [rec("2024 A형", 55, dd(12)), rec("2024 B형", 62, dd(11)),
      rec("코코 1회", 48, dd(13))]
p = sj.pick_set_for_retry(SETS, R1, 1, today)
check("① 신규 세트 우선 (기록 없는 코코 2회)",
      p[0][0] is Dd and "신규" in p[0][1], str(p[0][1]))
# (2) 최저점: 전부 기록 있음 -> 최고점 최저(코코 1회 48)
R2 = R1 + [rec("코코 2회", 70, dd(10))]
p = sj.pick_set_for_retry(SETS, R2, 1, today)
check("② 최저점 세트 (코코 1회 48점)", p[0][0] is C and "48" in p[0][1],
      str(p[0][1]))
# 최고점 기준: 코코 1회가 48 -> 80으로 올랐으면 A형(55)이 최저
R2b = R2 + [rec("코코 1회", 80, dd(9))]
p = sj.pick_set_for_retry(SETS, R2b, 1, today)
check("② 최고점 기준 (코코 1회 80 갱신 -> 2024 A형 55)", p[0][0] is A,
      str(p[0][1]))
# (3) 최근 2일 제외: 최저점 코코 1회를 어제 응시 -> 다음 최저(A형 55)
R3 = R2 + [rec("코코 1회", 50, dd(5))]
p = sj.pick_set_for_retry(SETS, R3, 1, today)
check("③ 최근 2일 내 응시 세트 제외 (코코 1회 어제 -> 2024 A형)",
      p[0][0] is A, str(p[0][1]))
R3b = R2 + [rec("코코 1회", 50, dd(6))]   # 2일 전 -> 허용
p = sj.pick_set_for_retry(SETS, R3b, 1, today)
check("③ 2일 전 응시는 허용 (코코 1회 복귀)", p[0][0] is C, str(p[0][1]))
R3c = [rec(s["name"], 60 + i, dd(5)) for i, s in enumerate(SETS)]
p = sj.pick_set_for_retry(SETS, R3c, 1, today)
check("③ 전부 최근이면 최저점 폴백 + 이유 명시",
      p[0][0] is A and "최근 2일" in p[0][1], str(p[0][1]))
# 부분연습/오답재풀이 기록은 '응시'로 치지 않음
R3d = R1 + [rec("코코 2회", 30, dd(5), mode="오답재풀이"),
            rec("코코 2회", 30, dd(5), mode="부분연습")]
p = sj.pick_set_for_retry(SETS, R3d, 1, today)
check("③ 부분연습·오답재풀이 기록만 있는 세트는 여전히 신규",
      p[0][0] is Dd and "신규" in p[0][1], str(p[0][1]))
# (4) 2세트 상이
p = sj.pick_set_for_retry(SETS, R2, 2, today)
check("④ 2세트 날: 서로 다른 세트 (코코 1회 48, 2024 A형 55)",
      len(p) == 2 and p[0][0] is C and p[1][0] is A
      and p[0][0]["norm"] != p[1][0]["norm"], str([x[1] for x in p]))
p = sj.pick_set_for_retry(SETS, R1, 2, today)
check("④ 신규 1개 + 최저점 1개 조합", p[0][0] is Dd and p[1][0] is C)
# (5) 기록 없음 폴백
p = sj.pick_set_for_retry(SETS, [], 2, today)
check("⑤ 기록 없음: 앞에서부터 신규 2개", [x[0] for x in p] == [A, B]
      and all("신규" in x[1] for x in p))
check("⑤ 세트 없음: 빈 결과", sj.pick_set_for_retry([], R1, 1, today) == [])
# resolve_day_sets: 자동 슬롯, 저장된 선택 복원 (하루 1세트)
plan5 = sj.plan_for_day(5, exam=EX)
check("Day 1~7 은 전부 자동 선택 1슬롯 (고정 세트 없음 · 하루 1세트)",
      all(sj.plan_for_day(n, exam=EX)["세트"] == [sj.AUTO]
          for n in range(1, 8)), str(plan5["세트"]))
rs = sj.resolve_day_sets(plan5, SETS, R2, today)
check("Day 5(자동 ×1) -> 최저점 세트 1개", len(rs) == 1 and rs[0][0] is C,
      str([(s["name"], r) for s, r in rs]))
rs2 = sj.resolve_day_sets(plan5, SETS, R2, today,
                          saved=[{"세트": "코코 2회", "이유": "직접 선택"}])
check("저장된 자동 선택 복원 (코코 2회 유지)",
      rs2[0][0] is Dd and "유지" in rs2[0][1])
plan3 = dict(sj.plan_for_day(3, exam=EX), 세트=["2024 A형"])
rs3 = sj.resolve_day_sets(plan3, SETS, [], today)
check("세트를 지정한 합성 plan 은 그 세트로 매칭 (엔진이 고른 세트가 들어오는 자리)",
      len(rs3) == 1 and rs3[0][0] is A, str(plan3["세트"]))
plan2x = dict(sj.plan_for_day(3, exam=EX), 세트=["2024 A형", "2024 상시 2회"])
rs4 = sj.resolve_day_sets(plan2x, SETS, [], today)
check("2슬롯 합성 plan: 없는 세트는 None + 사유",
      rs4[0][0] is A and rs4[1][0] is None and "찾지 못함" in rs4[1][1])

# 3. 오답 재풀이 ----------------------------------------------------------
print("3. 오답 재풀이 시트 추출 + --sheets 인자")
items = [{"sheet": "계산작업", "label": "계산 1"}, {"sheet": "기본작업-2 ", "label": "서식"},
         {"sheet": "계산작업", "label": "계산 3"}, {"sheet": "차트작업", "label": "차트"},
         {"sheet": "", "label": "?"}]
sheets = sj.wrong_sheets_from_items(items)
check("오답 시트 추출 (등장 순·중복 제거·공백 정리)",
      sheets == ["계산작업", "기본작업-2", "차트작업"], str(sheets))
cmd = sj.build_grade_cmd("grade.py", "p.xlsx", "a.xlsx", "s.xlsm",
                         json_out="o.json", sheets=sheets)
i = cmd.index("--sheets")
check("--sheets 인자 = '계산작업,기본작업-2,차트작업'",
      cmd[i + 1] == "계산작업,기본작업-2,차트작업" and "--json" in cmd, str(cmd[i:i + 2]))
check("시트 없으면 --sheets 없음",
      "--sheets" not in sj.build_grade_cmd("g", "p", "a", "s", sheets=[]))
# 최신 전체 채점 JSON에서 재풀이 payload (partial JSON은 건너뜀)
setdir = os.path.join(TMP, "세트X")
os.makedirs(os.path.join(setdir, "채점결과"))
X = mkset("세트X")
X["dir"] = setdir
X["problem"] = os.path.join(setdir, "세트X_문제.xlsx")
full = {"total": 61, "mode": "full", "files": {"problem": X["problem"]},
        "wrong_items": [{"sheet": "계산작업", "label": "a"},
                        {"sheet": "매크로작업", "label": "b"}]}
part = {"total": 10, "mode": "partial", "files": {"problem": X["problem"]},
        "wrong_items": [{"sheet": "계산작업", "label": "a"}]}
fp = os.path.join(setdir, "채점결과", "채점결과_세트X_20260905_1000.json")
pp = os.path.join(setdir, "채점결과", "채점결과_세트X_20260905_1100.json")
json.dump(full, open(fp, "w", encoding="utf-8"), ensure_ascii=False)
json.dump(part, open(pp, "w", encoding="utf-8"), ensure_ascii=False)
os.utime(fp, (1_700_000_000, 1_700_000_000))
os.utime(pp, (1_700_000_100, 1_700_000_100))   # partial이 더 최신
check("full_only=False -> 최신(partial) JSON",
      sj.find_latest_result_json(X) == pp)
check("full_only=True -> 전체 응시 JSON", sj.find_latest_result_json(X, full_only=True) == fp)
kind, payload = sj.retry_payload_for_set(X)
check("재풀이 payload: 오답 시트 2개·15분·mode 오답재풀이",
      kind == "retry" and payload["sheets"] == ["계산작업", "매크로작업"]
      and payload["minutes"] == 15 and payload["mode"] == "오답재풀이"
      and payload["점수"] == 61, str(payload)[:120])
step = sj.plan_for_day(1)["스텝"][3]
kind, payload = sj.resolve_step_action(step, [X], slot_sets=[(X, "일정 지정 세트")])
check("오답재풀이 스텝 -> retry 실행", kind == "retry" and payload["sheets"][0] == "계산작업")
json.dump({"total": 100, "mode": "full", "files": {"problem": X["problem"]},
           "wrong_items": []}, open(fp, "w", encoding="utf-8"))
kind, payload = sj.retry_payload_for_set(X)
check("오답 없으면 info + 자동완료", kind == "info" and payload.get("자동완료") is True)
Y = mkset("세트Y")
kind, payload = sj.retry_payload_for_set(Y)
check("채점 기록 없으면 missing", kind == "missing")
check("오답재풀이 사본 이름 규칙",
      os.path.basename(sj._unique_stem(TMP, "오답재풀이_세트X_20260905_1000")).startswith("오답재풀이_세트X_"))

# 4. 스텝 시퀀스 무결성 (Day 1~7) -------------------------------------------
print("4. 스텝 시퀀스 무결성 (필수 ①② + 선택 ③④⑤)")
TEMPLATE = ["모의", "채점", "오답노트", "오답재풀이"]
for no in range(1, sj.ROUTINE_DAYS + 1):
    plan = sj.plan_for_day(no, exam=EX)
    steps = plan["스텝"]
    kinds = [s["형"] for s in steps]
    n_slots = len(plan["세트"])
    body = kinds[:]
    pre = []
    if body and body[0] == "안내":
        pre.append(body.pop(0))
    tail = []
    while body and body[-1] == "안내":
        tail.insert(0, body.pop())
    req = sj.mandatory_step_indexes(steps)
    ok = (body == TEMPLATE * n_slots
          and all(steps[i].get("슬롯") == (i - len(pre)) // 4
                  for i in range(len(pre), len(pre) + 4 * n_slots))
          and tail and steps[-1].get("선택") is True
          and all(s.get("분") for s in steps)
          and all(s["이름"] and s["설명"] for s in steps)
          # 필수 = 슬롯마다 ①완주 ②채점 뿐, 나머지(오답노트·재풀이·퀴즈·실수
          # 노트)는 전부 선택
          and req == [len(pre) + 4 * k + j
                      for k in range(n_slots) for j in (0, 1)]
          and all(steps[i].get("형") in ("모의", "채점") for i in req))
    check(f"Day {no} (D-{sj.ROUTINE_DAYS + 1 - no}) "
          f"{sj.routine_date_for(no, EX).isoformat()} 슬롯 {n_slots} "
          f"스텝 {len(steps)} = {'+'.join(kinds)}", ok)
RD = sj.ROUTINE_DAYS
check("Day 1~7 전부 하루 1세트(5스텝 = 필수 2 + 선택 3, 마지막 날은 6스텝)",
      [len(sj.plan_for_day(n, exam=EX)["세트"]) for n in range(1, RD + 1)]
      == [1] * RD
      and [len(sj.plan_for_day(n, exam=EX)["스텝"]) for n in range(1, RD + 1)]
      == [5] * (RD - 1) + [6])
check("실수 노트 스텝은 마지막 학습일 Day 7(D-1)뿐이고 선택 단계, 복기 스텝 없음",
      any("실수 노트" in s["이름"] and s.get("선택") is True
          for s in sj.plan_for_day(RD, exam=EX)["스텝"])
      and not any("실수 노트" in s["이름"] for n in range(1, RD)
                  for s in sj.plan_for_day(n, exam=EX)["스텝"])
      and not any("복기" in s["이름"] for n in sj.PLAN_ORDER
                  for s in sj.plan_for_day(n, exam=EX)["스텝"]))
goals = {no: sj.plan_for_day(no, exam=EX)["목표"] for no in range(1, RD + 1)}
check("목표 점수: 기본 일정은 Day 1~7 전부 70(합격선 = 첫 단계)",
      goals == {no: 70 for no in range(1, RD + 1)}, str(goals))
check("목표 사다리 상수·문구: 70·80·90·100 / '목표 70점(합격선)'·'목표 80점(승급)'",
      sj.GOAL_TIERS == [70, 80, 90, 100]
      and sj.goal_label(70) == "목표 70점(합격선)"
      and sj.goal_label(80) == "목표 80점(승급)" and sj.goal_label(None) == "")
check("일정표에 고정 세트가 하나도 없다 (전부 자동 선택 슬롯)",
      all((sj.ROUTINE_PLAN[n].get("세트") or []) in ([], [sj.AUTO])
          for n in sj.ROUTINE_PLAN)
      and sj.all_slot_specs() == list(sj.SLOT_IDENTITY))
check("할일 문구: 완주 → 채점 + 목표 + 선택 안내 (오답노트 필수 표현 없음)",
      "40분 완주 → 채점" in sj.plan_for_day(1, exam=EX)["할일"]
      and "목표 70점(합격선)" in sj.plan_for_day(1, exam=EX)["할일"]
      and "선택" in sj.plan_for_day(1, exam=EX)["할일"],
      sj.plan_for_day(1, exam=EX)["할일"])
check("시험일/이후/준비: 스텝 없음(준비는 안내 4스텝)",
      not sj.plan_for_day(sj.PLAN_EXAM1, exam=EX)["스텝"]
      and not sj.plan_for_day(sj.PLAN_EXAM2, exam=EX2)["스텝"]
      and not sj.plan_for_day(sj.PLAN_AFTER, exam=EX)["스텝"]
      and [s["형"] for s in sj.plan_for_day(0, exam=EX)["스텝"]]
      == ["안내"] * 4)
check("ROUTINE_STEPS 조회 테이블 = plan_for_day 스텝",
      all(sj.ROUTINE_STEPS[n] == sj.plan_for_day(n)["스텝"]
          for n in sj.PLAN_ORDER))
# 스텝 실행 매핑
plan5s = sj.plan_for_day(5, exam=EX)
kind, payload = sj.resolve_step_action(plan5s["스텝"][0], SETS, slot_sets=[(Dd, "신규 세트라서")])
check("자동 슬롯 모의 스텝 -> exam(목표 70, 이유 전달)",
      kind == "exam" and payload["set"] is Dd and payload["목표"] == 70 and "신규" in payload["이유"])
kind, payload = sj.resolve_step_action(plan5s["스텝"][0], SETS, slot_sets=[(None, "없음")])
check("세트 미확정 모의 스텝 -> missing", kind == "missing")
kind, payload = sj.resolve_step_action(plan5s["스텝"][1], SETS)
check("채점 스텝 -> info(자동 체크 안내)", kind == "info")
kind, payload = sj.resolve_step_action(plan5s["스텝"][2], SETS, slot_sets=[(Dd, "")])
check("(선택) 오답노트 스텝 -> review(세트, 도구로 유지)",
      kind == "review" and payload["set"] is Dd
      and plan5s["스텝"][2]["선택"] is True)
step1 = dict(plan5s["스텝"][0], 세트="2024 상시 1회")
kind, payload = sj.resolve_step_action(step1, [A, mkset("2024년 상시1회 2급")])
check("세트 이름이 지정된 스텝: 퍼지 매칭으로 exam",
      kind == "exam" and "상시1회" in payload["set"]["name"])

# 5. 제목/D-day/진행 키 ---------------------------------------------------------
print("5. 카드 제목·D-day·저장 키")
t = sj.plan_title(dict(sj.plan_for_day(3, exam=EX), 제목="첫 응시"),
                  today=dd(5), set_names=["2024 A형"], exam=EX)
check("제목 'Day 3 (D-5) · 11/14(토) · 첫 응시 — 2024 A형' + D-day 병기",
      t.startswith("Day 3 (D-5) · 11/14(토) · 첫 응시 — 2024 A형")
      and "시험 D-5" in t, t)
check("D-day 문구에 2차가 있으면 함께 표기",
      "2차 D-6" in sj.plan_title(sj.plan_for_day(3, exam=EX2), today=dd(5),
                                 exam=EX2))
check("자동 선택 날 제목에 선택 세트명 반영",
      "코코 1회" in sj.plan_title(sj.plan_for_day(5, exam=EX),
                               set_names=["코코 1회"], exam=EX))
check("시험일/이후/준비 제목",
      sj.plan_title(sj.plan_for_day(sj.PLAN_EXAM1, exam=EX), exam=EX)
      .startswith("시험일 · 11/19(목) · 시험일")
      and sj.plan_title(sj.plan_for_day(sj.PLAN_EXAM2, exam=EX2), exam=EX2)
      .startswith("시험일 · 11/20(금) · 2차 시험일")
      and sj.plan_title(sj.plan_for_day(sj.PLAN_AFTER, exam=EX), exam=EX)
      .startswith("루틴 · 루틴 완주")
      and "Day 1" in sj.plan_title(sj.plan_for_day(0, exam=EX), exam=EX))
check("D-day: 시험 당일 'D-day', 이후 D+n, 미설정이면 '시험일 미설정'",
      sj.dday_text(EXAM, EX) == "시험 D-day"
      and sj.dday_text(dd(11), EX) == "시험 D-11"
      and sj.dday_text(EXAM + TD(days=3), EX) == "시험 D+3"
      and sj.dday_text(EXAM, EX2) == "시험 D-day · 2차 D-1"
      and sj.dday_text(EXAM, []) == "시험일 미설정")
check("dday_num · dday_label",
      sj.dday_num(dd(7), EX) == 7 and sj.dday_num(EXAM, EX) == 0
      and sj.dday_label(7) == "D-7" and sj.dday_label(0) == "D-day"
      and sj.dday_label(-2) == "D+2" and sj.dday_label(None) == "")
check("시험일 안내문: 시험 당일 체크리스트, 2차는 실수 노트 (날짜가 박혀 있지 않다)",
      "수험표" in sj.plan_for_day(sj.PLAN_EXAM1, exam=EX)["할일"]
      and "실수 노트" in sj.plan_for_day(sj.PLAN_EXAM2, exam=EX2)["할일"]
      and not any(re.search(r"\d{4}|\d{1,2}\s*/\s*\d{1,2}|\d{1,2}월",
                            sj.plan_for_day(no, exam=EX2)["할일"])
                  for no in sj.PLAN_ORDER))
check("day 라벨 새 체계 d01~d07 (시험일/이후 None)",
      sj.plan_day_tag(1) == "d01" and sj.plan_day_tag(7) == "d07"
      and sj.plan_day_tag(sj.PLAN_EXAM1) is None)
cfgp = os.path.join(TMP, "세트설정.json")
json.dump({"_진행": {"d04": [0, 1, 2]}}, open(cfgp, "w", encoding="utf-8"))  # 구 루틴 진행
check("구 루틴 진행(d04)은 새 루틴 d04에 섞이지 않음",
      sj.load_step_progress("d04", cfgp) == set())
sj.save_step_progress("d04", {0, 1}, cfgp)
cfg = json.load(open(cfgp, encoding="utf-8"))
NS = sj.progress_ns(sj.exam_dates(path=cfgp))
check("새 진행은 접두 키로 저장, 구 키 보존",
      cfg["_진행"]["d04"] == [0, 1, 2] and cfg["_진행"][NS + ":d04"] == [0, 1]
      and sj.load_step_progress("d04", cfgp) == {0, 1})
check("진행 키 접두는 시험일마다 달라 이전 루틴 진행과 섞이지 않는다",
      sj.progress_ns(EX) == "r20261119" and sj.progress_ns([]) == "r0"
      and sj.progress_ns([EXAM + TD(days=7)]) != sj.progress_ns(EX))
sj.save_auto_picks("d08", [{"세트": "코코 2회", "이유": "신규 세트라서"}], cfgp)
check("자동 선택 저장/복원", sj.load_auto_picks("d08", cfgp) == [{"세트": "코코 2회", "이유": "신규 세트라서"}])
summ = sj.records_summary(R3d)
check("records_summary: 오답재풀이/부분연습 제외", "코코 2회" not in summ and summ["2024 A형"]["best"] == 55)

# 6. v2.5.0 규칙: 목표 사다리 · 연속 응시 (v2.7.0: 하루 1회 제한 폐지) ---------
print("6. 목표 사다리 · 응시 제한 없음(v2.7.0) · 연속 응시")
L = sj.level_info
check("기록 없음 -> 현재 목표 70(tier 70, 달성 없음)",
      L([]) == {"tier": 70, "current_goal": 70, "best_total": None,
                "achieved": [], "next_goal": 70}, str(L([])))
check("최고점 58 -> 여전히 목표 70",
      L([rec("코코 1회", 58, dd(5))])["current_goal"] == 70)
check("최고점 72 -> 70 달성, 현재 목표 80",
      L([rec("코코 1회", 72, dd(5))]) ==
      {"tier": 80, "current_goal": 80, "best_total": 72, "achieved": [70],
       "next_goal": 80}, str(L([rec("코코 1회", 72, dd(5))])))
check("최고점 85 -> 70·80 달성, 현재 목표 90",
      L([rec("코코 1회", 85, dd(5))])["current_goal"] == 90
      and L([rec("코코 1회", 85, dd(5))])["achieved"] == [70, 80])
check("최고점 90 -> 목표 100", L([rec("A", 90, dd(5))])["current_goal"] == 100)
check("100 달성 -> tier 'perfect', next_goal None",
      L([rec("A", 100, dd(5))])["tier"] == "perfect"
      and L([rec("A", 100, dd(5))])["next_goal"] is None
      and L([rec("A", 100, dd(5))])["achieved"] == [70, 80, 90, 100])
check("레벨은 세트와 무관한 전체 최고점 기준 (어느 세트든 1회면 달성)",
      L([rec("코코 1회", 40, D(2026, 9, 10)), rec("2024 A형", 82, dd(6))]
        )["current_goal"] == 90)
check("부분연습·오답재풀이·점수 없는 기록은 레벨에 안 셈",
      L([rec("A", 95, D(2026, 9, 11), mode="부분연습"),
         rec("A", 95, D(2026, 9, 11), mode="오답재풀이"),
         {"일시": "2026-09-11 10:00", "세트명": "A", "점수": None,
          "mode": "시험"}])["best_total"] is None)
T = dd(5)                      # 오늘 = D-5
check("v2.7.0: 하루 1회 잠금 관련 함수·상수가 모두 사라짐 (응시 제한 없음)",
      not any(hasattr(sj, n) for n in
              ("exam_locked_today", "daily_lock_message", "DAILY_LOCK_TITLE",
               "lock_remaining_text", "lock_countdown_text")),
      str([n for n in ("exam_locked_today", "daily_lock_message",
                       "DAILY_LOCK_TITLE", "lock_remaining_text",
                       "lock_countdown_text") if hasattr(sj, n)]))
check("응시한 날짜 집합은 그대로 (연속 응시 계산용)",
      sj.attempt_dates([rec("코코 1회", 58, T), rec("A", 70, T),
                        rec("B", 70, dd(6))]) == {T, dd(6)})
check("점수 없는 중단 응시·오답재풀이·부분연습은 응시로 안 셈",
      sj.attempt_dates([{"일시": f"{T.isoformat()} 10:00", "세트명": "코코 1회",
                         "점수": None, "mode": "시험"},
                        rec("코코 1회", 58, T, mode="오답재풀이"),
                        rec("코코 1회", 58, T, mode="부분연습")]) == set())
ST = [rec("A", 60, dd(8)), rec("B", 62, dd(7)),
      rec("C", 64, dd(6)), rec("D", 66, T)]
check("연속 응시: D-8 ~ D-5 4일 -> current 4 · best 4 · today_done",
      sj.streak_info(ST, T, EX) == {"current": 4, "best": 4,
                                    "today_done": True},
      str(sj.streak_info(ST, T, EX)))
check("오늘 미응시면 어제까지의 연속을 유지 (today_done False)",
      sj.streak_info(ST[:3], T, EX) == {"current": 3, "best": 3,
                                        "today_done": False},
      str(sj.streak_info(ST[:3], T, EX)))
check("끊긴 연속: D-20·D-19 + D-6·D-5 -> current 2 · best 2",
      sj.streak_info([rec("A", 60, dd(20)), rec("B", 60, dd(19)),
                      rec("C", 60, dd(6)), rec("D", 60, T)], T, EX)
      == {"current": 2, "best": 2, "today_done": True})
check("시험일 기록은 연속에서 빠지고(today_done False) 전날까지의 연속은 유지",
      sj.streak_info([rec("A", 60, dd(1)), rec("B", 60, EXAM)], EXAM, EX)
      == {"current": 1, "best": 1, "today_done": False},
      str(sj.streak_info([rec("A", 60, dd(1)), rec("B", 60, EXAM)], EXAM, EX)))
check("하루 2건이어도 연속은 1일로 셈",
      sj.streak_info([rec("A", 60, T), rec("B", 70, T)], T, EX)["current"] == 1)
check("연속은 루틴 시작 전(D-7 이전) 기록도 이어서 센다 (시험일만 제외)",
      sj.streak_info([rec("A", 60, dd(12)), rec("B", 60, dd(11)),
                      rec("C", 60, dd(10))], dd(10), EX)["current"] == 3)
check("기록 없음 -> current 0 · best 0",
      sj.streak_info([], T, EX) == {"current": 0, "best": 0,
                                    "today_done": False})
check("클리어 안내문: 완주+채점만으로 클리어, 선택 단계 명시",
      "클리어" in sj.STEP_DONE_MESSAGE and "선택" in sj.STEP_DONE_MESSAGE)


# 7. v2.6.0 게임형 요소 ------------------------------------------------------
print()
print("7. 게임형 요소: 티어 · 게이지 · 도전 과제 · 결과 연출")
check("티어 이름 70 합격 · 80 숙련 · 90 고득점 · 100 만점 (그 외는 빈 문자열)",
      [sj.tier_label(g) for g in sj.GOAL_TIERS] == ["합격", "숙련", "고득점", "만점"]
      and sj.tier_label(None) == "" and sj.tier_label(75) == "")
check("level_from_best = level_info (기록 형식과 무관한 같은 결과)",
      sj.level_from_best(72) == L([rec("A", 72, dd(5))])
      and sj.level_from_best(None) == L([]))
check("게이지 비율 = 최고점 / 현재 목표 (0~1로 자름)",
      sj.gauge_fraction(72, 80) == 0.9 and sj.gauge_fraction(None, 70) == 0.0
      and sj.gauge_fraction(120, 100) == 1.0 and sj.gauge_fraction(50, 0) == 0.0)
check("소요시간 파싱: '38분 10초' · '39:10' · '1:02:03' · 숫자 · 읽을 수 없으면 0",
      (sj.parse_elapsed("38분 10초"), sj.parse_elapsed("39:10"),
       sj.parse_elapsed("1:02:03"), sj.parse_elapsed(2400),
       sj.parse_elapsed("-"), sj.parse_elapsed(None)) == (2290, 2350, 3723, 2400, 0, 0))
check("최고점 갱신 문구: '최고 58 → 76 (+18)' (갱신이 아니면 빈 문자열)",
      sj.best_update_text({"improved": True, "prev_best": 58, "best": 76,
                           "best_delta": 18}) == "최고 58 → 76 (+18)"
      and sj.best_update_text({"improved": False, "prev_best": 58, "best": 58,
                               "best_delta": None}) == ""
      and sj.best_update_text(None) == "")


def SHEETS(calc):
    """계산작업만 실점, 나머지 60점은 무실점인 시트 목록."""
    return [{"name": "기본작업-1", "alloc": 5, "earned": 5},
            {"name": "기본작업-2", "alloc": 10, "earned": 10},
            {"name": "기본작업-3", "alloc": 5, "earned": 5},
            {"name": "계산작업", "alloc": 40, "earned": calc},
            {"name": "분석작업-1", "alloc": 10, "earned": 10},
            {"name": "분석작업-2", "alloc": 10, "earned": 10},
            {"name": "매크로작업", "alloc": 10, "earned": 10},
            {"name": "차트작업", "alloc": 10, "earned": 10}]


def SR(rid, d, norm, total, calc=None, sec=0, mode="시험"):
    """규약 records[] 형식 1건 (achievements 판정 입력)."""
    return {"id": rid, "date": d, "time": "10:00", "mode": mode, "total": total,
            "set": {"name": norm, "norm": norm}, "pass_line": 70,
            "passed": total is not None and total >= 70,
            "sheets": SHEETS(calc) if calc is not None else [],
            "wrong_items": [], "elapsed": None, "elapsed_sec": sec,
            "report_html": None, "report_json": None}


def DONE(recs):
    return [a["id"] for a in sj.achievements(recs) if a["done"]]


A0 = sj.achievements([])
check("도전 과제 12종 · 고정 순서 · 필드(id/name/desc/icon/group/done/when)",
      len(A0) == 12 and [a["id"] for a in A0] == sj.ACHIEVEMENT_IDS
      and all(set(a) == {"id", "name", "desc", "icon", "group", "done", "when"}
              for a in A0) and not any(a["done"] for a in A0))
check("첫 완주 · 첫 합격 · 40분 내 완주 · 기본/분석/기타 무실점",
      DONE([SR("1", "2026-09-03", "a", 72, calc=12, sec=2290)])
      == ["first", "pass70", "base60", "fast40"])
check("40분(2400초)을 넘기면 fast40 미달성, 정확히 40분이면 달성",
      "fast40" not in DONE([SR("1", "2026-09-03", "a", 72, calc=12, sec=2401)])
      and "fast40" in DONE([SR("1", "2026-09-03", "a", 72, calc=12, sec=2400)]))
check("계산작업 40/40 → calc40, 100점이면 tier100(전 영역 무실점)",
      DONE([SR("1", "2026-09-03", "a", 100, calc=40)])
      == ["first", "pass70", "tier80", "tier90", "tier100", "calc40", "base60"])
check("한 영역이라도 실점하면 base60 미달성",
      "base60" not in DONE([SR("1", "2026-09-03", "a", 90,
                               calc=40) | {"sheets": SHEETS(40)[:-1] + [
                                   {"name": "차트작업", "alloc": 10,
                                    "earned": 8}]}]))
STREAK5 = [SR(str(i + 1), f"2026-09-0{i + 3}", k, 72, calc=12)
           for i, k in enumerate("abcde")]
check("3일 연속 · 5일 연속 · 세트 5개 완주 (달성 날짜까지)",
      {a["id"]: a["when"] for a in sj.achievements(STREAK5) if a["done"]}
      | {} == {"first": "2026-09-03", "pass70": "2026-09-03",
               "base60": "2026-09-03", "streak3": "2026-09-05",
               "streak5": "2026-09-07", "sets5": "2026-09-07"},
      str({a["id"]: a["when"] for a in sj.achievements(STREAK5) if a["done"]}))
check("연속이 끊기면 다시 1일부터 (9/3·9/4·9/6 → streak3 미달성)",
      "streak3" not in DONE([SR("1", "2026-09-03", "a", 72, calc=12),
                             SR("2", "2026-09-04", "b", 72, calc=12),
                             SR("3", "2026-09-06", "c", 72, calc=12)]))
check("수동 기록도 완주로 세고, 부분연습·오답재풀이·중단 응시는 빼고 센다",
      DONE([SR("1", "2026-09-03", "a", 80, mode="수동")]) == ["first", "pass70", "tier80"]
      and DONE([SR("1", "2026-09-03", "a", 40, calc=40, mode="부분연습"),
                SR("2", "2026-09-03", "b", 20, mode="오답재풀이"),
                SR("3", "2026-09-04", "c", None)]) == [])
check("같은 세트를 여러 번 응시해도 완주 세트 수는 1 (sets5 미달성)",
      "sets5" not in DONE([SR(str(i), f"2026-09-0{i + 3}", "a", 72, calc=12)
                           for i in range(5)]))
check("이상한 기록이 섞여도 죽지 않는다(리스트 아님·시트 형식 오류·날짜 이상)",
      DONE([None, 5, "x", {"mode": "시험"}, {"date": "2026-09", "total": 90},
            {"id": "1", "date": "2026-09-03", "time": "10:00", "mode": "시험",
             "total": 75, "set": "세트아님", "sheets": [1, {"alloc": "x"}]}])
      == ["first", "pass70"])
SUMM = sj.achievement_summary(STREAK5)
check("achievement_summary: 달성 수 · 전체 · 최근 달성",
      SUMM["done"] == 6 and SUMM["total"] == 12
      and SUMM["latest"]["when"] == "2026-09-07")
check("serialized_best: 규약 records[] 최고점(점수 없는 기록 제외)",
      sj.serialized_best(STREAK5) == 72 and sj.serialized_best([]) is None)
B4 = [SR("1", "2026-09-03", "a", 72, calc=12, sec=2290)]
A4 = B4 + [SR("2", "2026-09-04", "b", 84, calc=24, sec=2350)]
PD = sj.progress_delta(B4, A4)
check("progress_delta: 직전 최고 대비 증감 · 티어 승급 · 새 배지",
      PD["prev_best"] == 72 and PD["best"] == 84 and PD["best_delta"] == 12
      and PD["tier_up"] == [80] and [a["id"] for a in PD["new_badges"]] == ["tier80"]
      and PD["badges_done"] == 5 and PD["badges_total"] == 12, str(PD))
check("progress_delta: 첫 기록이면 prev_best None · best_delta None",
      sj.progress_delta([], B4)["prev_best"] is None
      and sj.progress_delta([], B4)["best_delta"] is None
      and [a["id"] for a in sj.progress_delta([], B4)["new_badges"]]
      == ["first", "pass70", "base60", "fast40"])
check("progress_delta: 점수가 낮아도 승급·새 배지는 없고 증감만 음수",
      sj.progress_delta(A4, A4 + [SR("3", "2026-09-05", "c", 60, calc=0)])["best_delta"] == 0
      and sj.progress_delta(A4, A4 + [SR("3", "2026-09-05", "c", 60, calc=0)])["tier_up"] == [])
check("new_achievements: 없으면 빈 목록", sj.new_achievements(A4, A4) == [])


# 8. v2.7.0 점수대 세트 보드 -------------------------------------------------
print()
print("8. 점수대 밴드 · 세트 보드 · 재응시 점수 갱신")
check("밴드 정의 6개 · 순서 고정 (미응시 → 미달 → 합격 → 숙련 → 고득점 → 만점)",
      [b["id"] for b in sj.SCORE_BANDS] == sj.BAND_IDS
      == ["none", "fail", "pass", "skilled", "high", "perfect"]
      and [b["name"] for b in sj.SCORE_BANDS]
      == ["미응시", "미달", "합격", "숙련", "고득점", "만점"])
check("경계값 69/70 → 미달/합격", sj.score_band(69) == "fail"
      and sj.score_band(70) == "pass")
check("경계값 79/80 → 합격/숙련", sj.score_band(79) == "pass"
      and sj.score_band(80) == "skilled")
check("경계값 89/90 → 숙련/고득점", sj.score_band(89) == "skilled"
      and sj.score_band(90) == "high")
check("경계값 99/100 → 고득점/만점", sj.score_band(99) == "high"
      and sj.score_band(100) == "perfect")
check("0 은 미달 · 기록 없음(None)은 미응시 · 100 초과도 만점",
      sj.score_band(0) == "fail" and sj.score_band(None) == "none"
      and sj.score_band(101) == "perfect")
check("숫자가 아니면 미응시 (True·문자열·잘못된 값)",
      sj.score_band(True) == "none" and sj.score_band("80") == "none"
      and sj.score_band([]) == "none")
check("소수점 점수도 구간 그대로 (69.5 미달 · 99.5 고득점)",
      sj.score_band(69.5) == "fail" and sj.score_band(99.5) == "high")
check("밴드 이름 조회", sj.band_label("fail") == "미달"
      and sj.band_label("perfect") == "만점" and sj.band_label("없음") == "")

BSETS = [{"name": "코코 1회", "norm": "코코1회"},
         {"name": "2024 A형", "norm": "2024a형"},
         {"name": "2024 B형", "norm": "2024b형"},
         {"name": "상시 1회", "norm": "상시1회"}]
BR = [SR("1", "2026-09-03", "코코1회", 58),
      SR("2", "2026-09-05", "코코1회", 76),
      SR("3", "2026-09-04", "2024a형", 92),
      SR("4", "2026-09-06", "2024b형", 100)]
BOARD = sj.score_board(BSETS, BR)
BYN = {r["norm"]: r for r in BOARD["sets"]}
check("세트 진행: 점수 이력·최고점·응시 횟수·마지막 응시일",
      BYN["코코1회"]["scores"] == [58, 76] and BYN["코코1회"]["best"] == 76
      and BYN["코코1회"]["attempts"] == 2
      and BYN["코코1회"]["last_date"] == "2026-09-05", str(BYN["코코1회"]))
check("재응시로 최고점이 오르면 밴드 이동 (미달 → 합격) + 갱신 폭",
      BYN["코코1회"]["prev_best"] == 58 and BYN["코코1회"]["prev_band"] == "fail"
      and BYN["코코1회"]["band"] == "pass" and BYN["코코1회"]["improved"] is True
      and BYN["코코1회"]["best_delta"] == 18
      and sj.best_update_text(BYN["코코1회"]) == "최고 58 → 76 (+18)")
check("한 번만 응시한 세트는 prev_best None · improved True · best_delta None",
      BYN["2024a형"]["prev_best"] is None and BYN["2024a형"]["improved"] is True
      and BYN["2024a형"]["best_delta"] is None
      and sj.best_update_text(BYN["2024a형"]) == "")
check("미응시 세트는 band none · 이력 빈 목록 · 마지막 응시일 None",
      BYN["상시1회"] == {"name": "상시 1회", "norm": "상시1회", "attempts": 0,
                      "best": None, "band": "none", "scores": [],
                      "last_date": None, "prev_best": None,
                      "prev_band": "none", "best_delta": None,
                      "improved": False}, str(BYN["상시1회"]))
check("밴드 묶음: 미응시 1 · 합격 1 · 고득점 1 · 만점 1 (빈 밴드는 0)",
      {b["id"]: b["count"] for b in BOARD["bands"]}
      == {"none": 1, "fail": 0, "pass": 1, "skilled": 0, "high": 1,
          "perfect": 1},
      str({b["id"]: b["count"] for b in BOARD["bands"]}))
check("밴드 요약 한 줄은 빈 밴드를 빼고 순서대로",
      sj.band_summary_text(BOARD["bands"])
      == "미응시 1 · 합격 1 · 고득점 1 · 만점 1",
      sj.band_summary_text(BOARD["bands"]))
check("세트 하나도 없으면 '세트 없음'",
      sj.band_summary_text(sj.score_board([], [])["bands"]) == "세트 없음")
BR2 = BR + [SR("5", "2026-09-07", "코코1회", 71),
            SR("6", "2026-09-08", "2024b형", 62)]
B2 = {r["norm"]: r for r in sj.score_board(BSETS, BR2)["sets"]}
check("점수가 떨어져도 최고점·밴드는 유지, improved False",
      B2["코코1회"]["scores"] == [58, 76, 71] and B2["코코1회"]["best"] == 76
      and B2["코코1회"]["band"] == "pass"
      and B2["코코1회"]["improved"] is False
      and B2["코코1회"]["last_date"] == "2026-09-07")
check("낮은 점수로 다시 풀어도 밴드는 최고점 기준 (만점 유지)",
      B2["2024b형"]["band"] == "perfect" and B2["2024b형"]["best"] == 100)
FAILSETS = [{"name": "가", "norm": "가"}, {"name": "나", "norm": "나"},
            {"name": "다", "norm": "다"}]
FB = sj.score_board(FAILSETS, [SR("1", "2026-09-03", "가", 65),
                               SR("2", "2026-09-03", "나", 40),
                               SR("3", "2026-09-03", "다", 65)])
check("밴드 안 정렬: 최고점 오름차순 → 동점은 이름순",
      [b["sets"] for b in FB["bands"] if b["id"] == "fail"][0]
      == ["나", "가", "다"],
      str([b["sets"] for b in FB["bands"] if b["id"] == "fail"][0]))
check("미응시 밴드는 입력(우선순위) 순서를 지킨다",
      [b["sets"] for b in sj.score_board(BSETS, [])["bands"]
       if b["id"] == "none"][0] == ["코코1회", "2024a형", "2024b형", "상시1회"])
check("보드 판정도 부분연습·오답재풀이·중단 응시를 제외한다",
      sj.score_board([{"name": "가", "norm": "가"}],
                     [SR("1", "2026-09-03", "가", 95, mode="부분연습"),
                      SR("2", "2026-09-03", "가", 90, mode="오답재풀이"),
                      SR("3", "2026-09-04", "가", None)])["sets"][0]["band"]
      == "none")
check("수동 기록(웹 점수 입력)은 보드에 반영",
      sj.score_board([{"name": "가", "norm": "가"}],
                     [SR("1", "2026-09-03", "가", 83, mode="수동")]
                     )["sets"][0]["band"] == "skilled")
check("같은 날 2회 응시도 이력에 둘 다 남고 연속은 1일",
      sj.score_board([{"name": "가", "norm": "가"}],
                     [SR("1", "2026-09-03", "가", 58),
                      dict(SR("2", "2026-09-03", "가", 77), time="15:00")]
                     )["sets"][0]["scores"] == [58, 77]
      and sj.streak_info([rec("가", 58, D(2026, 9, 3)),
                          rec("가", 77, D(2026, 9, 3))],
                         D(2026, 9, 3))["current"] == 1)
check("보드 밴드와 목표 사다리(level)가 어긋나지 않는다 (최고 76 → 합격·목표 80)",
      sj.score_band(76) == "pass"
      and sj.level_from_best(76)["current_goal"] == 80
      and sj.score_band(100) == "perfect"
      and sj.level_from_best(100)["tier"] == "perfect")


# 9. v3.0.0 첫 실행 안내 · 세트 넣는 방법 ------------------------------------
print()
print("9. 첫 실행 안내(자동 업데이트 고지) · 세트 넣는 방법")
NP = os.path.join(TMP, "고지설정.json")
check("첫 실행에는 고지가 '아직'(pending) 상태",
      sj.first_run_notice_pending(NP) is True)
sj.mark_first_run_notice(NP)
check("한 번 띄우면 다음 실행부터는 뜨지 않는다",
      sj.first_run_notice_pending(NP) is False
      and json.load(open(NP, encoding="utf-8"))["_설정"][
          sj.FIRST_RUN_NOTICE_SETTING] is True)
NT = sj.first_run_notice_text()
check("동의 문구(v3.1.0): 동의/거절 두 갈래 · 저장소 주소 · 닫으면 거절 · 설정에서 변경",
      "동의" in NT and "거절" in NT and sj.UPDATE_REPO_URL in NT
      and "github.com" in sj.UPDATE_REPO_URL
      and sj.AUTO_UPDATE_SETTING in NT and "닫으면" in NT
      and "업데이트 확인" in NT and "[설정]" in NT, NT)
check("동의/거절 버튼 문구가 무엇을 하는지 말해 준다",
      "동의" in sj.UPDATE_CONSENT_YES and "거절" in sj.UPDATE_CONSENT_NO
      and "끄기" in sj.UPDATE_CONSENT_NO,
      sj.UPDATE_CONSENT_YES + " / " + sj.UPDATE_CONSENT_NO)
check("저장소 주소는 배포 주소에서 끌어낸다 (따로 박아 두지 않는다)",
      sj.UPDATE_REPO_URL == sj._repo_url(sj.UPDATE_BASE_URL)
      and sj.UPDATE_BASE_URL.startswith("https://raw.githubusercontent.com/")
      and sj.UPDATE_REPO_URL.startswith("https://github.com/"))
check("자동 업데이트는 기본 켜짐이고 설정으로 끌 수 있다",
      sj.auto_update_enabled(NP) is True
      and (sj.set_app_setting(sj.AUTO_UPDATE_SETTING, False, NP),
           sj.auto_update_enabled(NP))[1] is False)
HELP = sj.set_intake_help("/학습폴더")
check("세트 넣는 방법 안내: 문제/정답 파일 이름 규칙 · PDF · 기대값 없어도 됨 · 경로",
      "_문제.xlsx" in HELP and "_정답.xlsx" in HELP and "_문제지.pdf" in HELP
      and "기대값" in HELP and "없어도" in HELP and "/학습폴더" in HELP
      and "세트 인식 진단" in HELP, HELP)
check("시험일 미설정 안내문에 [시험일 설정]과 D-7~D-1 이 들어 있다",
      "시험일 설정" in sj.EXAM_UNSET_TODO and "D-7" in sj.EXAM_UNSET_TODO
      and "D-1" in sj.EXAM_UNSET_TODO)
check("세트 없음 사유 문구",
      sj.NO_SETS_REASON == "세트 없음 — 문제·정답 파일을 넣어 주세요")
check("시험일 미설정이면 Day 0 카드가 '시험일 설정' 안내로 바뀐다",
      sj.plan_for_day(0, exam=[])["제목"] == "시험일 설정"
      and sj.plan_for_day(0, exam=[])["할일"] == sj.EXAM_UNSET_TODO
      and sj.plan_for_day(0, exam=EX)["제목"] == "자유 연습")
check("준비 스텝에 시험일 설정·내 문제 파일 넣기가 들어 있다",
      [x["이름"] for x in sj.PREP_STEPS][:2]
      == ["시험일 설정", "내 문제·정답 파일 넣기"],
      str([x["이름"] for x in sj.PREP_STEPS]))


# 10. v3.1.0 자동 업데이트 동의 / 거절 ---------------------------------------
print()
print("10. v3.1.0 자동 업데이트 동의 / 거절 (거절하면 꺼지고 프로그램은 정상)")
CP = os.path.join(TMP, "동의설정.json")
check("첫 실행에는 아직 안 물어본 상태",
      sj.first_run_notice_pending(CP) is True
      and sj.auto_update_enabled(CP) is True)
check("거절하면 _설정.자동업데이트=false 로 꺼지고 다시 묻지 않는다",
      sj.apply_update_consent(False, CP) is False
      and sj.auto_update_enabled(CP) is False
      and sj.first_run_notice_pending(CP) is False
      and json.load(open(CP, encoding="utf-8"))["_설정"][
          sj.AUTO_UPDATE_SETTING] is False)
_uc = sj.UpdateCoordinator(base_dir=TMP, enabled=lambda: False,
                           log=lambda *a, **k: None)
_auto, _manual = _uc.check(timeout=1), _uc.check(force=True, timeout=1)
check("거절해도 프로그램은 정상 — 자동 확인만 건너뛰고 수동 [업데이트 확인]은 그대로",
      _auto == "disabled" and _manual != "disabled",
      f"자동={_auto} 수동={_manual}")
check("거절 뒤에도 설정으로 다시 켤 수 있다",
      (sj.set_app_setting(sj.AUTO_UPDATE_SETTING, True, CP),
       sj.auto_update_enabled(CP))[1] is True)
CP2 = os.path.join(TMP, "동의설정2.json")
check("동의하면 자동 업데이트가 켜진 채로 표시된다",
      sj.apply_update_consent(True, CP2) is True
      and sj.auto_update_enabled(CP2) is True
      and sj.first_run_notice_pending(CP2) is False)
check("'응답 없음/창 닫음'은 거절과 같은 경로(거짓 값이면 전부 꺼짐)",
      all(sj.apply_update_consent(v, os.path.join(TMP, f"동의{i}.json"))
          is False for i, v in enumerate((None, "", 0, False))))


# 11. v3.1.0 시험일 목록 (추가·제거·0개·마이그레이션) --------------------------
print()
print("11. v3.1.0 시험일 목록 — 여러 개 추가·개별 제거·0개 허용·옛 설정 이전")
EP = os.path.join(TMP, "시험일설정.json")
check("처음에는 0개 (미설정)",
      sj.exam_dates(path=EP) == [] and sj.exam_date_set([]) is False)
D1, D2, D3 = date(2026, 11, 19), date(2026, 11, 22), date(2027, 3, 7)
sj.save_exam_date_list([D2, D1, D3, D1], EP)
check("여러 개 저장 — 오름차순·중복 제거",
      sj.exam_dates(path=EP) == [D1, D2, D3], str(sj.exam_dates(path=EP)))
check("_설정.시험일목록 에 저장되고 구버전용 시험일/시험일2 미러도 남는다",
      json.load(open(EP, encoding="utf-8"))["_설정"][sj.EXAM_DATES_SETTING]
      == [D1.isoformat(), D2.isoformat(), D3.isoformat()]
      and json.load(open(EP, encoding="utf-8"))["_설정"][
          sj.EXAM_DATE_SETTING] == D1.isoformat()
      and json.load(open(EP, encoding="utf-8"))["_설정"][
          sj.EXAM_DATE2_SETTING] == D2.isoformat())
check("개별 제거", sj.remove_exam_date(D2, EP) == ([D1, D3], True)
      and sj.exam_dates(path=EP) == [D1, D3])
check("없는 날짜 제거는 아무 일도 안 일어난다",
      sj.remove_exam_date(D2, EP) == ([D1, D3], False))
check("추가 (이미 있으면 그대로)",
      sj.add_exam_date(D2, EP) == ([D1, D2, D3], True)
      and sj.add_exam_date(D2, EP) == ([D1, D2, D3], False)
      and sj.add_exam_date("아무날", EP)[1] is False)
check("전체 비우기 — 0개도 허용",
      sj.clear_exam_dates(EP) == [] and sj.exam_dates(path=EP) == []
      and sj.exam_date_set([]) is False)
OLD = os.path.join(TMP, "옛설정.json")
json.dump({"_설정": {"시험일": "2026-11-19", "시험일2": "2026-11-22"}},
          open(OLD, "w", encoding="utf-8"), ensure_ascii=False)
check("옛 시험일/시험일2 설정 파일을 읽어도 안 깨지고 목록으로 읽힌다",
      sj.exam_dates(path=OLD) == [D1, D2], str(sj.exam_dates(path=OLD)))
OLD2 = os.path.join(TMP, "옛설정2.json")
json.dump({"_설정": {"시험일": "2026-11-19", "시험일2": None}},
          open(OLD2, "w", encoding="utf-8"), ensure_ascii=False)
check("2차가 null 인 옛 설정도 1개로 읽힌다",
      sj.exam_dates(path=OLD2) == [D1])
BAD = os.path.join(TMP, "깨진설정.json")
json.dump({"_설정": {"시험일목록": ["2026-11-19", "엉망", None, 3]}},
          open(BAD, "w", encoding="utf-8"), ensure_ascii=False)
check("깨진 항목은 버리고 성한 것만 쓴다 (프로그램이 죽지 않는다)",
      sj.exam_dates(path=BAD) == [D1])

# 기준 시험일 = 가장 가까운 미래 (지난 시험일은 목록에 남되 기준에서 빠짐)
T = date(2026, 11, 20)
check("여러 개 — 기준은 가장 가까운 미래 시험일",
      sj.effective_exams([D1, D2, D3], T) == [D2, D3]
      and sj.dday_num(T, [D1, D2, D3]) == 2
      and sj.routine_day_no(D2 - TD(days=1), [D1, D2, D3]) == 7)
check("전부 과거면 목록 그대로 — '시험 이후' 안내로 간다",
      sj.effective_exams([D1, D2], D3) == [D1, D2]
      and sj.routine_day_no(D3, [D1, D2]) == sj.PLAN_AFTER
      and sj.dday_num(D3, [D1, D2]) < 0)
check("지난 시험일이 있어도 다음 시험일 기준으로 일정이 선다 (3.0.0 은 막혔다)",
      sj.segment_for(T, [D1, D3])["kind"] == "free"
      and sj.segment_for(D3 - TD(days=3), [D1, D3])["kind"] == "week"
      and sj.build_adaptive_plan(D3 - TD(days=3), [], [], [D1, D3])["kind"]
      == "week")
check("0개면 일정 대신 안내 + 자유 연습 (죽지 않는다)",
      sj.segment_for(T, [])["kind"] == "unset"
      and sj.build_adaptive_plan(T, [], [], [])["kind"] == "unset"
      and sj.routine_day_no(T, []) == 0
      and sj.dday_num(T, []) is None
      and sj.dday_text(T, []) == "시험일 미설정"
      and sj.plan_for_day(0, exam=[])["제목"] == "시험일 설정"
      and sj.study_segments([], T) == []
      and sj.routine_start([], T) is None)
check("1개면 그 날 기준 D-7 ~ D-1",
      sj.routine_day_no(D1 - TD(days=7), [D1]) == 1
      and sj.routine_day_no(D1 - TD(days=1), [D1]) == 7
      and sj.routine_day_no(D1, [D1]) == sj.PLAN_EXAM1
      and sj.routine_start([D1], D1 - TD(days=7)) == D1 - TD(days=7))
check("진행도 네임스페이스·루틴 태그도 기준 시험일을 따라 옮겨 간다",
      sj.progress_ns([D1, D3], T) == "r" + D3.strftime("%Y%m%d")
      and sj.routine_tag([D1, D3], T) == D3.isoformat()
      and sj.progress_ns([], T) == "r0" and sj.routine_tag([], T) == "")


# 12. v3.1.0 세트 스캔 폴더 설정 ---------------------------------------------
print()
print("12. v3.1.0 세트 스캔 폴더 — GUI 에서 바꾸고 설정에 남는다")
SP = os.path.join(TMP, "스캔설정.json")
ROOT_A = os.path.join(TMP, "공부폴더A")
os.makedirs(ROOT_A, exist_ok=True)
check("기본값은 시험장.py 상위 폴더",
      sj.configured_scan_root(path=SP) == sj.default_scan_root()
      and sj.scan_root_is_default(path=SP) is True)
check("폴더를 고르면 저장되고 다음 실행에도 유지된다",
      sj.save_scan_root(ROOT_A, SP) == os.path.abspath(ROOT_A)
      and sj.configured_scan_root(path=SP) == os.path.abspath(ROOT_A)
      and sj.scan_root_is_default(path=SP) is False
      and json.load(open(SP, encoding="utf-8"))["_설정"][
          sj.SCAN_ROOT_SETTING] == os.path.abspath(ROOT_A))
check("[기본값으로] 되돌리기",
      sj.save_scan_root(None, SP) == sj.default_scan_root()
      and sj.configured_scan_root(path=SP) == sj.default_scan_root()
      and sj.scan_root_is_default(path=SP) is True)
sj.set_app_setting(sj.SCAN_ROOT_SETTING, os.path.join(TMP, "없는폴더"), SP)
check("저장된 폴더가 사라졌으면 조용히 기본값으로 (죽지 않는다)",
      sj.configured_scan_root(path=SP) == sj.default_scan_root())


# 13. v3.1.0 세트 개별 숨기기 -------------------------------------------------
print()
print("13. v3.1.0 세트 숨기기 — 목록·일정·보드에서 빠지고 파일·기록은 그대로")
HP = os.path.join(TMP, "숨김설정.json")
HS = [mkset("가 세트"), mkset("나 세트"), mkset("다 세트")]
check("처음에는 숨긴 세트 없음",
      sj.hidden_set_keys(path=HP) == []
      and sj.split_hidden_sets(list(HS), path=HP)[0] == HS)
check("숨기면 설정에 정규화 키로 남는다",
      sj.set_set_hidden("나 세트", True, HP) == [sj.norm_set_key("나 세트")]
      and sj.hidden_set_keys(path=HP) == [sj.norm_set_key("나 세트")])
shown, gone = sj.split_hidden_sets([dict(s) for s in HS], path=HP)
check("숨긴 세트는 보이는 목록에서 빠진다",
      [s["name"] for s in shown] == ["가 세트", "다 세트"]
      and [s["name"] for s in gone] == ["나 세트"]
      and gone[0]["숨김"] is True)
HREC = [{"일시": "2026-11-01 10:00", "세트명": "나 세트", "점수": 88,
         "mode": "시험"}]
PLAN_H = sj.build_adaptive_plan(date(2026, 11, 15), HREC, shown, [D1])
check("일정에도 숨긴 세트가 안 들어간다",
      all(sl["name"] != "나 세트" for day in PLAN_H["days"].values()
          for sl in day), str(PLAN_H["days"]))
check("점수대 보드에서도 빠진다",
      [r["name"] for r in sj.score_board(shown, [])["sets"]]
      == ["가 세트", "다 세트"])
check("기록은 지우지 않는다 — 숨겨도 기록.json 은 그대로",
      sj.set_exam_records("나 세트", HREC) == HREC)
check("다시 보이게 하면 그대로 돌아온다",
      sj.set_set_hidden("나 세트", False, HP) == []
      and [s["name"] for s in sj.split_hidden_sets(
          [dict(s) for s in HS], path=HP)[0]]
      == ["가 세트", "나 세트", "다 세트"])
sj.set_set_hidden("가 세트", True, HP)
sj.set_set_hidden("다 세트", True, HP)
check("전부 다시 보이기",
      len(sj.hidden_set_keys(path=HP)) == 2 and sj.clear_hidden_sets(HP) == []
      and sj.hidden_set_keys(path=HP) == [])
DIAG = sj.scan_diagnosis_text(
    ROOT_A, config={"_설정": {sj.HIDDEN_SETS_SETTING: [sj.norm_set_key("나 세트")]}},
    sets=[HS[0]], hidden_sets=[HS[1]])
check("세트 인식 진단에 스캔 폴더와 숨긴 세트가 드러난다",
      "스캔 폴더" in DIAG and ROOT_A in DIAG and "숨긴 세트 1개" in DIAG
      and "나 세트" in DIAG, DIAG.split("\n")[0])
check("숨긴 세트가 없으면 '숨긴 세트: 없음'",
      "숨긴 세트: 없음" in sj.scan_diagnosis_text(ROOT_A, config={},
                                              sets=[HS[0]]))

print()
print(f"일정 뼈대 헤드리스 테스트 {N}건 전부 통과")
