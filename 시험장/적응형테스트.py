# -*- coding: utf-8 -*-
"""시험장.py 적응형 일정 엔진 헤드리스 테스트 (tkinter 불필요).

일정(v3.0.0): **사용자가 입력한 시험일** 기준 1주 — 학습일 D-7 ~ D-1.
  · 7일이 안 남았으면 뒤에서부터 잘려 오늘 ~ D-1 만 배정된다.
  · 7일보다 많이 남았으면 D-7 전은 자유 연습(kind free)이고 앞으로 올
    일주일은 미리 보여 준다.
  · 시험일이 지났으면 kind after("시험일을 다시 설정하세요"),
    시험일이 없으면 kind unset.
규칙: 하루 추천 1세트(응시 횟수 제한 없음) · 목표 사다리 70→80→90→100 ·
      세트 선택 ① 미응시(이름순) ② 70점 미만(최고점 낮은 순) ③ 현재 목표
      미달(최고점 낮은 순) ④ 그래도 날이 남으면 보유 세트 순환 재응시
      ⑤ 전 세트가 이미 목표 이상이면 그 날은 자유 복습(빈 슬롯).
      **고정 세트 슬롯은 없다** — 무엇을 푸는지는 보유 세트로 정해진다.

  ⓪ 구간 도출(시험일 기준) · 용량(언제나 1) · 승격·강등·경고 폐지
  ① 보유 세트 수 N = 0 · 1 · 3 · 7 · 10 경계 배정
  ② 오늘 시나리오 (a) 무기록 (b) 72 합격 (c) 58 미달
  ③ 응시 제한 없음: 오늘 응시해도 추천 유지 · 남은 학습일 그대로
  ④ 목표 승급: 70 달성 후 목표 80 — 합격한 세트도 다시 배정(kind goal)
  ⑤ 전 세트 만점(현재 목표 이상) → 빈 슬롯 · all_clear · 자유 복습 카드
  ⑥ 신규 세트 편입 · 재응시 후보 대기줄(retry_queue)
  ⑦ 결정성 + 카드 plan(adaptive_day_plan) 형식 · 시험일/이후/미설정
  ⑧ 시험일 경계: 7일 미만(뒤에서 자르기) · 7일 초과(자유 연습) · 지남 · 미설정
"""
import copy
import importlib.util
import os
from datetime import date, timedelta

BASE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("sj", os.path.join(BASE, "시험장.py"))
sj = importlib.util.module_from_spec(spec)
spec.loader.exec_module(sj)
N = 0
D = date
TD = timedelta

EXAM = D(2026, 11, 19)      # 테스트용 시험일 (코드에 박힌 값이 아님)
EX = [EXAM]
EX2 = [EXAM, EXAM + TD(days=1)]
RD = sj.ROUTINE_DAYS        # 7


def dd(n):
    """시험일로부터 n일 전(D-n) 날짜."""
    return EXAM - TD(days=n)


def check(desc, cond, extra=""):
    global N
    N += 1
    print(f"  [{N:02d}] {desc}: {'OK' if cond else 'FAIL'}" + (f" ({extra})" if extra else ""))
    if not cond:
        raise AssertionError(f"{desc} {extra}")


def mkset(name):
    return {"name": name, "norm": sj.norm_set_key(name), "dir": "/tmp/x",
            "problem": f"/tmp/x/{name}_문제.xlsx", "answer": f"/tmp/x/{name}_정답.xlsm",
            "key": None, "pdf": None}


def pool(n, prefix="세트"):
    """보유 세트 n개 (이름순 = 배정 순서)."""
    return [mkset(f"{prefix} {i:02d}") for i in range(1, n + 1)]


NAMES = [f"세트 {i:02d}" for i in range(1, 10)]
SETS = pool(9)
BY = {s["name"]: s for s in SETS}


def rec(name, score, d, mode="시험"):
    return {"일시": f"{d.isoformat()} 10:00", "세트명": name, "점수": score, "mode": mode}


def names_by_day(plan):
    """{ISO 날짜: [(세트명 또는 AUTO, 목표)...]}"""
    return {d.isoformat(): [(sl["name"] if not sl["auto"] else "AUTO", sl["goal"]) for sl in v]
            for d, v in plan["days"].items()}


def flat_names(plan):
    return [sl["name"] for d in sorted(plan["days"]) for sl in plan["days"][d]]


def kinds_by_day(plan):
    return {d.isoformat(): [sl["kind"] for sl in v] for d, v in plan["days"].items()}


def counts(plan):
    return [len(plan["days"][d]) for d in sorted(plan["days"])]


def queue_of(plan):
    return [(q["name"], q["kind"], q["best"]) for q in plan["retry_queue"]]


WEEK = [dd(RD - i) for i in range(RD)]        # D-7 … D-1


# ⓪ 구간 도출·용량 ---------------------------------------------------------
print("⓪ 구간 도출(시험일 기준) · 하루 1세트 용량")
segs = sj.study_segments(EX)
check("시험일 -> 학습 구간은 D-7 ~ D-1 하나뿐",
      len(segs) == 1 and segs[0]["seg_start"] == dd(RD) and segs[0]["end"] == dd(1)
      and segs[0]["exam"] == EXAM and sj.next_segment_after(1, EX) is None, str(segs))
check("시험일 미설정이면 학습 구간 없음", sj.study_segments([]) == [])
sf = sj.segment_for
check("segment_for: D-20→free · D-7→week(전체) · D-4→week(오늘부터) · D-1→week · 시험일→exam · 이후→after",
      sf(dd(20), EX)["kind"] == "free"
      and sf(dd(RD), EX) == {"kind": "week", "seg": 1, "seg_start": dd(RD),
                             "start": dd(RD), "end": dd(1)}
      and sf(dd(4), EX)["start"] == dd(4) and sf(dd(4), EX)["end"] == dd(1)
      and sf(dd(1), EX)["start"] == dd(1) == sf(dd(1), EX)["end"]
      and sf(EXAM, EX) == {"kind": "exam", "exam_no": 1}
      and sf(EXAM + TD(days=1), EX) == {"kind": "after"})
check("segment_for: 시험일 미설정 → unset · 2차 당일 → exam(순번 2) · 2차 다음 → after",
      sf(dd(3), []) == {"kind": "unset"}
      and sf(EXAM + TD(days=1), EX2) == {"kind": "exam", "exam_no": 2}
      and sf(EXAM + TD(days=2), EX2) == {"kind": "after"})
check("v3.1.0 — 지난 시험일은 기준에서 빠지고 다음 시험일 기준으로 주간이 선다 "
      "(3.0.0 의 between 구간은 없어졌다)",
      sf(EXAM + TD(days=1), [EXAM, EXAM + TD(days=3)])
      == {"kind": "week", "seg": 1, "seg_start": EXAM + TD(days=3) - TD(days=RD),
          "start": EXAM + TD(days=1), "end": EXAM + TD(days=2)}
      and sf(EXAM + TD(days=1), [EXAM, EXAM + TD(days=30)])["kind"] == "free",
      str(sf(EXAM + TD(days=1), [EXAM, EXAM + TD(days=3)])))
END = dd(1)
check("기본 용량: 주말·평일·마감일 구분 없이 언제나 1세트 (DAILY_CAP 1)",
      sj.DAILY_CAP == 1 and all(sj.base_capacity(d, END) == 1 for d in WEEK))
check("D-7~D-1 총 용량 7 = 학습일 7일 × 1 (= 추천 세트 수)",
      sum(sj.base_capacity(d, END) for d in WEEK) == 7)

# ① 보유 세트 수 경계 -------------------------------------------------------
print("① 보유 세트 수 N = 0 · 1 · 3 · 7 · 10 (D-7 무기록)")
p0 = sj.build_adaptive_plan(dd(RD), [], [], EX)
check("N=0: 배정 없음 · free 전부 1 · total_sets 0 · no_sets True",
      counts(p0) == [0] * 7 and [p0["free"][d] for d in WEEK] == [1] * 7
      and p0["total_sets"] == 0 and p0["remaining"] == 0
      and p0["no_sets"] is True and p0["all_clear"] is False, str(counts(p0)))
check("N=0: 사유는 '세트 없음 — 문제·정답 파일을 넣어 주세요'",
      p0["reason"] == sj.NO_SETS_REASON, p0["reason"])
card0 = sj.adaptive_day_plan(p0, 1, EX)
check("N=0 카드: '세트 없음' + 넣는 방법 안내(파일 이름 규칙)",
      card0["제목"] == "세트 없음" and "_문제.xlsx" in card0["할일"]
      and "_정답.xlsx" in card0["할일"] and card0["스텝"] == [], card0["제목"])

p1 = sj.build_adaptive_plan(dd(RD), [], pool(1), EX)
check("N=1: 첫날 first, 나머지 6일은 같은 세트 순환 재응시(kind goal)",
      flat_names(p1) == ["세트 01"] * 7
      and [k for d in sorted(p1["days"]) for k in kinds_by_day(p1)[d.isoformat()]]
      == ["first"] + ["goal"] * 6, str(kinds_by_day(p1)))
check("N=1: 빈 슬롯 없음 · 남은 배정 7 · 사유는 목표·남은 학습일",
      all(v == 0 for v in p1["free"].values()) and p1["remaining"] == 7
      and p1["reason"] == "현재 목표 70점 · 남은 학습일 7일", p1["reason"])

p3 = sj.build_adaptive_plan(dd(RD), [], pool(3), EX)
check("N=3: 앞 3일 첫 응시(이름순) · 남은 4일은 같은 순서로 순환 재응시",
      flat_names(p3) == ["세트 01", "세트 02", "세트 03",
                         "세트 01", "세트 02", "세트 03", "세트 01"],
      str(flat_names(p3)))
check("N=3: 순환 배정의 사유는 '오답 복습 후 재응시'",
      "오답 복습 후 재응시" in p3["days"][dd(4)][0]["why"], p3["days"][dd(4)][0]["why"])

p7 = sj.build_adaptive_plan(dd(RD), [], pool(7), EX)
check("N=7: 7일 = 7세트 첫 응시 정확히 하루 하나 (순환 없음)",
      flat_names(p7) == [f"세트 {i:02d}" for i in range(1, 8)]
      and set(kinds_by_day(p7)[d.isoformat()][0] for d in WEEK) == {"first"},
      str(flat_names(p7)))

p10 = sj.build_adaptive_plan(dd(RD), [], pool(10), EX)
check("N=10: 앞 7세트만 배정(나머지는 미배정 — 경고 아님) · total_sets 10",
      flat_names(p10) == [f"세트 {i:02d}" for i in range(1, 8)]
      and p10["total_sets"] == 10 and p10["remaining"] == 7
      and p10["warning"] is None, str(flat_names(p10)))
check("승격·강등·경고 폐지: boost_days·demoted 빈 목록, warning None, retry_target 0",
      all(p["boost_days"] == [] and p["demoted"] == [] and p["warning"] is None
          and p["retry_target"] == 0 for p in (p0, p1, p3, p7, p10)))
check("레벨·남은 학습일·진행: 목표 70 · attempts_left 7 · 합격 0",
      all(p["level"]["current_goal"] == 70 and p["level"]["tier"] == 70
          and p["attempts_left"] == 7 and p["passed_sets"] == 0
          for p in (p1, p3, p7, p10)))
check("사유는 목표·남은 학습일 한 줄 (미완주 재배치·응시 기회 문구 없음)",
      p7["reason"] == "현재 목표 70점 · 남은 학습일 7일"
      and "재배치" not in p7["reason"] and "기회" not in p7["reason"], p7["reason"])
check("연속 응시 0 · 오늘 미응시 · today_locked 는 하위 호환용으로 항상 False",
      p7["streak"] == {"current": 0, "best": 0, "today_done": False}
      and p7["today_locked"] is False)
check("첫 응시 순서는 이름순 (고정 우선순위 표 없음)",
      [s["name"] for s, _l in sj.prioritized_sets(SETS)] == NAMES
      and [s["name"] for s, _l in sj.prioritized_sets(list(reversed(SETS)))] == NAMES)
check("첫 응시 사유에 '보유 세트 n/전체' 순번이 들어간다",
      p7["days"][dd(RD)][0]["why"] == "첫 응시 · 보유 세트 1/7",
      p7["days"][dd(RD)][0]["why"])

# ② 오늘 시나리오 ----------------------------------------------------------
print("② D-5 (a) 무기록 · (b) 72 합격 · (c) 58 미달")
T = dd(5)
S1 = NAMES[0]
pa = sj.build_adaptive_plan(T, [], SETS, EX)
nba = names_by_day(pa)
check("(a) D-5 무기록: D-5~D-1 각 1세트 = 앞 5세트, 목표 전부 70",
      [nba[d.isoformat()] for d in WEEK[2:]] == [[(n, 70)] for n in NAMES[:5]], str(nba))
check("(a) 남은 학습일 5 · 남은 배정 5 · 지나간 날 D-7·D-6",
      pa["attempts_left"] == 5 and pa["remaining"] == 5 and pa["today_locked"] is False
      and pa["missed_days"] == WEEK[:2] and len(pa["days"]) == 5)
check("(a) 사유: 현재 목표 70점 · 남은 학습일 5일 (경고·강등 없음)",
      pa["reason"] == "현재 목표 70점 · 남은 학습일 5일" and pa["warning"] is None
      and pa["demoted"] == [], pa["reason"])

pb = sj.build_adaptive_plan(T, [rec(S1, 72, T)], SETS, EX)
nbb = names_by_day(pb)
check("(b) 72점 → 70 달성 · 현재 목표 80 · 합격 세트 1",
      pb["level"] == {"tier": 80, "current_goal": 80, "best_total": 72,
                      "achieved": [70], "next_goal": 80}
      and pb["passed_sets"] == 1, str(pb["level"]))
check("(b) 오늘 응시했어도 오늘 추천은 그대로(응시 제한 없음) · 남은 학습일 5 유지",
      nbb[T.isoformat()] == [(NAMES[1], 80)] and pb["free"][T] == 0
      and pb["today_locked"] is False and pb["attempts_left"] == 5
      and pb["streak"] == {"current": 1, "best": 1, "today_done": True}, str(nbb))
check("(b) 이미 푼 세트는 미응시 세트가 남아 있는 동안 다시 배정되지 않고 목표는 전부 80",
      [nbb[d.isoformat()] for d in WEEK[2:]] == [[(n, 80)] for n in NAMES[1:6]]
      and all(S1 != n for v in nbb.values() for n, _g in v), str(nbb))
check("(b) 그 세트는 '목표 80 도전' 후보로 대기줄 1순위 (합격해도 다음 단계 대상)",
      queue_of(pb) == [(S1, "goal", 72)], str(queue_of(pb)))
check("(b) 사유: 현재 목표 80점 · 남은 학습일 5일",
      pb["reason"] == "현재 목표 80점 · 남은 학습일 5일", pb["reason"])

pc = sj.build_adaptive_plan(T, [rec(S1, 58, T)], SETS, EX)
nbc = names_by_day(pc)
check("(c) 58점 → 달성 단계 없음 · 목표 70 유지 · 합격 세트 0",
      pc["level"] == {"tier": 70, "current_goal": 70, "best_total": 58,
                      "achieved": [], "next_goal": 70} and pc["passed_sets"] == 0,
      str(pc["level"]))
check("(c) 합격선 미달 → 재응시 후보 1순위(kind retry, 최고점 58)",
      queue_of(pc) == [(S1, "retry", 58)], str(queue_of(pc)))
check("(c) 오늘도 추천 유지(잠금 없음) · 미응시 세트 먼저 · 목표 전부 70",
      pc["today_locked"] is False
      and [nbc[d.isoformat()] for d in WEEK[2:]] == [[(n, 70)] for n in NAMES[1:6]], str(nbc))
check("(b)(c) 차이: 합격이면 목표 80 사다리 · 미달이면 목표 70 유지 (배정 세트는 같음)",
      [n for v in nbb.values() for n, _g in v] == [n for v in nbc.values() for n, _g in v]
      and pb["level"]["current_goal"] == 80 and pc["level"]["current_goal"] == 70)

# ③ 응시 제한 없음 ---------------------------------------------------------
print("③ 응시 제한 없음: 오늘 응시해도 추천 유지")
DAY_NO = RD + 1 - 5          # D-5 = Day 3
card_b = sj.adaptive_day_plan(pb, DAY_NO, EX)
check("(b) 오늘 카드: 추천 세트 스텝 유지 · '오늘 이미 1회 응시' + 제한 없음 안내 · 잠금 키 없음",
      card_b["제목"] == "오늘" and len(card_b["스텝"]) == 5 and "잠금" not in card_b
      and card_b["오늘응시"] == 1 and "오늘 추천 세트" in card_b["할일"]
      and "횟수 제한은 없습니다" in card_b["할일"], card_b["할일"])
check("(c) 오늘 카드: 합격선 미달 안내(리포트 확인, 체크 강제 아님) + 바로 재응시 권유",
      "리포트" in (sj.adaptive_day_plan(pc, DAY_NO, EX).get("안내") or "")
      and "필수 아님" in sj.adaptive_day_plan(pc, DAY_NO, EX)["안내"]
      and "다시 풀어" in sj.adaptive_day_plan(pc, DAY_NO, EX)["안내"],
      str(sj.adaptive_day_plan(pc, DAY_NO, EX).get("안내")))
check("(b) 합격(70 이상)이면 미달 안내 문구 없음",
      sj.adaptive_day_plan(pb, DAY_NO, EX).get("안내") is None)
check("점수 없는 중단 응시는 연속·오늘 응시로 안 셈",
      sj.adaptive_day_plan(sj.build_adaptive_plan(
          T, [{"일시": f"{T.isoformat()} 10:00", "세트명": S1,
               "점수": None, "mode": "시험"}], SETS, EX), DAY_NO, EX)["오늘응시"] == 0)
check("오답재풀이·부분연습 기록도 오늘 응시로 안 셈",
      sj.adaptive_day_plan(sj.build_adaptive_plan(
          T, [rec(S1, 58, T, mode="오답재풀이"),
              rec(S1, 58, T, mode="부분연습")], SETS, EX), DAY_NO, EX)["오늘응시"] == 0)
check("수동 기록(웹 점수 입력)은 오늘 응시로 인정하되 잠기지 않는다",
      sj.adaptive_day_plan(sj.build_adaptive_plan(
          T, [rec(S1, 58, T, mode="수동")], SETS, EX), DAY_NO, EX)["오늘응시"] == 1
      and sj.build_adaptive_plan(T, [rec(S1, 58, T, mode="수동")], SETS,
                                 EX)["today_locked"] is False)
P_TWICE = sj.build_adaptive_plan(T, [rec(S1, 58, T),
                                     dict(rec(S1, 76, T), 일시=f"{T.isoformat()} 15:00")],
                                 SETS, EX)
check("같은 날 2회 응시: 응시 제한 없이 기록 2건 · 연속은 1일 · 남은 학습일 5 유지",
      sj.adaptive_day_plan(P_TWICE, DAY_NO, EX)["오늘응시"] == 2
      and P_TWICE["streak"] == {"current": 1, "best": 1, "today_done": True}
      and P_TWICE["attempts_left"] == 5 and P_TWICE["today_locked"] is False,
      str(P_TWICE["streak"]))
check("같은 날 재응시로 최고점이 오르면 목표 단계가 바로 올라간다 (58 → 76 → 목표 80)",
      P_TWICE["level"]["best_total"] == 76 and P_TWICE["level"]["current_goal"] == 80
      and queue_of(P_TWICE) == [(S1, "goal", 76)], str(queue_of(P_TWICE)))

# ③-2 진행 중: 미응시 먼저, 그 다음 미달 재응시 ----------------------------
print("③-2 진행 중: 미응시 먼저, 그 다음 미달 재응시(최고점 낮은 순)")
R_mid = [rec(NAMES[0], 60, dd(12)), rec(NAMES[1], 62, dd(11)),
         rec(NAMES[2], 58, dd(10)), rec(NAMES[3], 65, dd(9)),
         rec(NAMES[4], 71, dd(8)), rec(NAMES[5], 66, dd(7)),
         rec(NAMES[6], 68, dd(6)), rec(NAMES[7], 55, dd(5))]
p4 = sj.build_adaptive_plan(dd(4), R_mid, SETS, EX)
nb4 = names_by_day(p4)
check("D-4: 미응시 세트 09 가 먼저(첫 응시), 그 다음은 미달 최고점 낮은 순 08(55) → 03(58) → 01(60)",
      [nb4[d.isoformat()] for d in WEEK[3:]]
      == [[(NAMES[8], 80)], [(NAMES[7], 80)], [(NAMES[2], 80)], [(NAMES[0], 80)]], str(nb4))
check("D-4 목표: 세트 05 가 71로 70 달성 → 현재 목표 80 (모든 슬롯 목표 80)",
      p4["level"]["current_goal"] == 80
      and all(sl["goal"] == 80 for v in p4["days"].values() for sl in v))
check("D-4 kind: 첫 응시 first · 합격선 미달 재응시 retry · 사유에 최고점·미달 표기",
      kinds_by_day(p4)[dd(4).isoformat()] == ["first"]
      and kinds_by_day(p4)[dd(3).isoformat()] == ["retry"]
      and "최고점 55점" in p4["days"][dd(3)][0]["why"]
      and "미달" in p4["days"][dd(3)][0]["why"], p4["days"][dd(3)][0]["why"])
check("D-4 재응시 대기줄: 미달 세트 오름차순 다음에 목표 미달 합격 세트(71)",
      queue_of(p4) == [(NAMES[7], "retry", 55), (NAMES[2], "retry", 58), (NAMES[0], "retry", 60),
                       (NAMES[1], "retry", 62), (NAMES[3], "retry", 65), (NAMES[5], "retry", 66),
                       (NAMES[6], "retry", 68), (NAMES[4], "goal", 71)], str(queue_of(p4)))
check("D-4 진행: 합격 세트 1 / 전체 9 · 남은 배정 4 · 남은 학습일 4",
      p4["passed_sets"] == 1 and p4["total_sets"] == 9 and p4["remaining"] == 4
      and p4["attempts_left"] == 4)
check("응시 없이 지나간 날만 missed_days (경고 아님)",
      p4["missed_days"] == [dd(4 + 1)] * 0 + [d for d in WEEK[:3]
                                              if d not in (dd(7), dd(6), dd(5))]
      and sj.build_adaptive_plan(dd(4), [r for r in R_mid
                                         if r["일시"][:10] != dd(6).isoformat()],
                                 SETS, EX)["missed_days"] == [dd(6)],
      str(p4["missed_days"]))
p4b = sj.build_adaptive_plan(dd(4), R_mid + [rec(NAMES[8], 61, dd(4))], SETS, EX)
check("응시 후에도 오늘 추천은 남는다(제한 없음) · 미달 재응시 순서 유지",
      names_by_day(p4b)[dd(4).isoformat()] == [(NAMES[7], 80)]
      and p4b["today_locked"] is False
      and flat_names(p4b) == [NAMES[7], NAMES[2], NAMES[0], NAMES[8]], str(flat_names(p4b)))
check("재응시 풀은 순환: 남은 날이 많으면 같은 세트가 다시 돌아온다(같은 날 중복은 없음)",
      flat_names(sj.build_adaptive_plan(dd(RD), [rec(NAMES[k], 50 + k, dd(20 - k))
                                                 for k in range(9)], SETS, EX))
      == NAMES[:7])

# ④ 목표 승급: 합격한 세트도 다시 배정 -------------------------------------
print("④ 목표 승급(70→80): 합격 세트도 현재 목표 미달이면 다시 배정")
R_pass = [rec(NAMES[k], 71 + k, dd(20 - k)) for k in range(9)]   # 71~79
p2 = sj.build_adaptive_plan(dd(2), R_pass, SETS, EX)
check("전 세트 71~79점(전부 합격) → 최고점 79 · 현재 목표 80 · 합격 9/9",
      p2["level"] == {"tier": 80, "current_goal": 80, "best_total": 79,
                      "achieved": [70], "next_goal": 80}
      and p2["passed_sets"] == 9, str(p2["level"]))
check("D-2·D-1 = 최고점 낮은 순(71 → 72)으로 '목표 80 도전'(kind goal)",
      names_by_day(p2) == {dd(2).isoformat(): [(NAMES[0], 80)],
                           dd(1).isoformat(): [(NAMES[1], 80)]}
      and kinds_by_day(p2) == {dd(2).isoformat(): ["goal"], dd(1).isoformat(): ["goal"]}
      and "목표 80점" in p2["days"][dd(2)][0]["why"], str(names_by_day(p2)))
check("합격(70 이상) 세트도 재배정 대상 — all_clear 아님",
      p2["all_clear"] is False and len(p2["retry_queue"]) == 9)
R_85 = [rec(NAMES[k], 81 + k, dd(20 - k)) for k in range(9)]     # 81~89
p2b = sj.build_adaptive_plan(dd(2), R_85, SETS, EX)
check("80 단계까지 달성(최고 89) → 현재 목표 90 · 슬롯 목표 90",
      p2b["level"]["current_goal"] == 90 and p2b["level"]["achieved"] == [70, 80]
      and all(sl["goal"] == 90 for v in p2b["days"].values() for sl in v))
p2c = sj.build_adaptive_plan(dd(2), [rec(NAMES[0], 100, dd(20))], SETS, EX)
check("100 달성 → tier perfect · 남은 세트는 목표 100으로 계속 배정",
      p2c["level"]["tier"] == "perfect" and p2c["level"]["next_goal"] is None
      and all(sl["goal"] == 100 for v in p2c["days"].values() for sl in v)
      and p2c["reason"] == "만점 달성 · 남은 학습일 2일", p2c["reason"])

# ⑤ 전 세트 목표 달성 → 빈 슬롯 -------------------------------------------
print("⑤ 전 세트가 현재 목표 이상 → 빈 슬롯(자유 복습)")
R_full = [rec(n, 100, dd(20)) for n in NAMES]
p4c = sj.build_adaptive_plan(dd(4), R_full, SETS, EX)
check("전 세트 100점: 배정 없음(D-4~D-1 빈 슬롯 4) · all_clear · 대기줄 빈 목록",
      counts(p4c) == [0, 0, 0, 0] and all(v == 1 for v in p4c["free"].values())
      and p4c["all_clear"] is True and p4c["retry_queue"] == [] and p4c["remaining"] == 0)
check("전 세트 달성이어도 경고·강등 없음 · 사유는 만점 문구",
      p4c["warning"] is None and p4c["demoted"] == []
      and p4c["reason"] == "만점 달성 · 남은 학습일 4일", p4c["reason"])
card4 = sj.adaptive_day_plan(p4c, RD + 1 - 4, EX)
check("빈 슬롯 카드: '전 세트 목표 달성' · 필수 과제 없음 안내 · 스텝 없음",
      card4["제목"] == "전 세트 목표 달성" and card4["스텝"] == []
      and "필수 과제 없음" in card4["할일"] and "자유 복습" in card4["할일"], card4["할일"])
p4n = sj.build_adaptive_plan(dd(4), [], [], EX)
check("세트가 하나도 없으면 전부 빈 슬롯 · total_sets 0",
      counts(p4n) == [0, 0, 0, 0] and p4n["total_sets"] == 0 and p4n["remaining"] == 0)

# ⑥ 신규 세트·대기줄 -------------------------------------------------------
print("⑥ 신규 세트 편입 · 재응시 후보 대기줄")
NEW = mkset("하나 새 세트")          # 이름순으로 맨 뒤
p4d = sj.build_adaptive_plan(dd(4), R_mid, SETS + [NEW], EX)
check("새로 넣은 세트도 미응시라 첫 응시 대상(이름순 뒤)",
      names_by_day(p4d)[dd(4).isoformat()] == [(NAMES[8], 80)]
      and names_by_day(p4d)[dd(3).isoformat()] == [("하나 새 세트", 80)]
      and p4d["days"][dd(3)][0]["kind"] == "first"
      and "첫 응시" in p4d["days"][dd(3)][0]["why"], str(names_by_day(p4d)))
NEWS = [mkset(f"가나다 세트 {k}") for k in range(1, 4)]
p7n = sj.build_adaptive_plan(dd(RD), [], NEWS + [NEW], EX)
check("이름순 배정(가나다 1 < 가나다 2 < 가나다 3 < 하나 새 세트) 후 순환",
      flat_names(p7n) == ["가나다 세트 1", "가나다 세트 2", "가나다 세트 3", "하나 새 세트",
                          "가나다 세트 1", "가나다 세트 2", "가나다 세트 3"],
      str(flat_names(p7n)))
check("점수 없는 기록만 있는 세트는 '미응시'가 아니라 최하위 재응시 후보",
      queue_of(sj.build_adaptive_plan(dd(3), [{"일시": f"{dd(20).isoformat()} 10:00",
                                               "세트명": NAMES[0], "점수": None,
                                               "mode": "시험"}], SETS, EX))
      == [(NAMES[0], "retry", None)])
check("부분연습·오답재풀이 기록만 있는 세트는 여전히 미응시(첫 응시 대상)",
      names_by_day(sj.build_adaptive_plan(dd(3), [rec(NAMES[0], 90, dd(20), mode="부분연습"),
                                                  rec(NAMES[0], 90, dd(19), mode="오답재풀이")],
                                          SETS, EX))[dd(3).isoformat()]
      == [(NAMES[0], 70)])

# ⑦ 결정성·카드 plan 형식 --------------------------------------------------
print("⑦ 결정성 · 카드 plan 형식")
a = sj.build_adaptive_plan(dd(6), copy.deepcopy(R_mid[:3]), copy.deepcopy(SETS), EX)
b = sj.build_adaptive_plan(dd(6), copy.deepcopy(R_mid[:3]), list(reversed(copy.deepcopy(SETS))), EX)
check("같은 입력(세트 순서 무관) → 같은 출력",
      names_by_day(a) == names_by_day(b) and a["reason"] == b["reason"] and queue_of(a) == queue_of(b))
today_plan = sj.adaptive_day_plan(pa, DAY_NO, EX)     # (a) 무기록 D-5 = Day 3
check("오늘 카드 plan: 적응형 · 세트객체 · 목표들 · 슬롯사유 · 스텝 5(필수 2 + 선택 3)",
      today_plan["적응형"] and today_plan["세트"] == [NAMES[0]]
      and [s["name"] for s in today_plan["세트객체"]] == [NAMES[0]]
      and today_plan["목표들"] == [70] and today_plan["목표"] == 70
      and len(today_plan["스텝"]) == 5 and sj.mandatory_step_indexes(today_plan["스텝"]) == [0, 1]
      and today_plan["스텝"][0]["세트"] == NAMES[0] and today_plan["사유"] == pa["reason"],
      str(today_plan["세트"]))
check("오늘 카드 할일: '오늘 추천 세트' + 완주 → 채점 + 목표 문구 + 선택 안내",
      "오늘 추천 세트" in today_plan["할일"] and "40분 완주 → 채점" in today_plan["할일"]
      and "목표 70점(합격선)" in today_plan["할일"] and "선택" in today_plan["할일"],
      today_plan["할일"])
t = sj.plan_title(today_plan, today=T, exam=EX)
check("제목: 'Day 3 (D-5) · 11/14(토) · 오늘: 세트 01 · 남은 학습일 5일 · 시험 D-5'",
      t == f"Day 3 (D-5) · 11/14(토) · 오늘: {NAMES[0]} · 남은 학습일 5일 · 시험 D-5", t)
fut = sj.adaptive_day_plan(pa, DAY_NO + 2, EX)
check("미래 카드(Day 5 = D-3) '예정:' + 1세트(목표 70)",
      fut["제목"] == "예정" and fut["세트"] == [NAMES[2]] and fut["목표들"] == [70]
      and "예정: " in sj.plan_title(fut, exam=EX))
futL = sj.adaptive_day_plan(pa, RD, EX)
check("마지막 학습일 카드(Day 7 = D-1): 실수 노트 특별(선택 스텝) · 복기 없음",
      futL["특별"] == ["실수노트"] and any("실수 노트" in s["이름"] and s.get("선택") for s in futL["스텝"])
      and not any("복기" in (pl.get("특별") or [])
                  for pl in (sj.adaptive_day_plan(pa, n, EX) for n in range(1, RD + 1))))
past = sj.adaptive_day_plan(sj.build_adaptive_plan(dd(6), [rec(NAMES[0], 60, dd(7))], SETS, EX), 1, EX)
check("지난 날 카드 '완료: 세트(점수)'",
      past["종류"] == "지난" and past["제목"] == "완료" and past["세트표시"] == [f"{NAMES[0]}(60점)"]
      and "완료: " in sj.plan_title(past, exam=EX) and "클리어" in past["할일"])
past2 = sj.adaptive_day_plan(pa, 1, EX)
check("응시 없이 지난 날 카드: '응시 없음' · 스텝 없음 · 재배치·미완주·기회 표현 없음",
      past2["제목"] == "응시 없음" and past2["스텝"] == []
      and "재배치" not in past2["할일"] and "미완주" not in past2["할일"]
      and "기회" not in past2["할일"] and "점수대 보드" in past2["할일"], past2["할일"])
rs = sj.resolve_day_sets(today_plan, SETS, [], T)
check("resolve_day_sets: 적응형 세트객체 우선 사용 + 사유",
      rs[0][0] is BY[NAMES[0]] and rs[0][1].startswith("적응형 배정 · 첫 응시"), str(rs[0][1]))
ex1 = sj.build_adaptive_plan(EXAM, R_mid, SETS, EX2)
ex2p = sj.build_adaptive_plan(EXAM + TD(days=1), R_mid, SETS, EX2)
check("시험일은 kind exam(1·2), 카드는 고정 안내, 그 다음 날은 after",
      ex1["kind"] == "exam" and ex1["exam_no"] == 1 and ex2p["exam_no"] == 2
      and sj.adaptive_day_plan(ex1, sj.PLAN_EXAM1, EX2)["제목"] == "시험일"
      and sj.adaptive_day_plan(ex2p, sj.PLAN_EXAM2, EX2)["제목"] == "2차 시험일"
      and sj.build_adaptive_plan(EXAM + TD(days=2), R_mid, SETS, EX2)["kind"] == "after")
check("시험일·이후에도 레벨·연속 응시는 계산해 담는다(웹 표시용)",
      ex1["level"]["current_goal"] == 80 and ex1["streak"]["best"] >= 1
      and ex1["attempts_left"] == 0 and ex1["days"] == {})
check("일정표 API(plan_for_day/ROUTINE_STEPS) 유지: Day 3 자동 선택 1세트 5스텝 · 목표 70",
      sj.plan_for_day(3, exam=EX)["세트"] == [sj.AUTO] and len(sj.ROUTINE_STEPS[3]) == 5
      and sj.plan_for_day(RD, exam=EX)["목표"] == 70)

# ⑧ 시험일 경계 ------------------------------------------------------------
print("⑧ 시험일 경계: 7일 미만 · 7일 초과 · 지남 · 미설정")
p_cut = sj.build_adaptive_plan(dd(3), [], SETS, EX)
check("D-3에 처음 실행: 뒤에서부터 잘려 D-3·D-2·D-1 3일만 배정",
      sorted(p_cut["days"]) == [dd(3), dd(2), dd(1)]
      and p_cut["attempts_left"] == 3 and p_cut["remaining"] == 3
      and flat_names(p_cut) == NAMES[:3]
      and p_cut["start"] == dd(3) and p_cut["end"] == dd(1), str(sorted(p_cut["days"])))
p_cut1 = sj.build_adaptive_plan(dd(1), [], SETS, EX)
check("D-1에 처음 실행: 하루(D-1)만 배정",
      sorted(p_cut1["days"]) == [dd(1)] and p_cut1["attempts_left"] == 1
      and p_cut1["reason"] == "현재 목표 70점 · 남은 학습일 1일")
p_far = sj.build_adaptive_plan(dd(20), [], SETS, EX)
check("D-20(7일 초과): kind free — 오늘은 자유 연습, D-7부터의 일주일을 미리 보여 준다",
      p_far["kind"] == "free" and sorted(p_far["days"]) == WEEK
      and p_far["start"] == dd(RD) and p_far["end"] == dd(1)
      and flat_names(p_far) == NAMES[:7], p_far["kind"])
check("D-20 사유: 자유 연습 · D-7부터 시작 안내",
      p_far["reason"] == f"자유 연습 · {dd(RD).isoformat()}(D-{RD})부터 일주일 일정 시작 · 현재 목표 70점",
      p_far["reason"])
check("D-8도 아직 자유 연습, D-7부터 week",
      sj.build_adaptive_plan(dd(8), [], SETS, EX)["kind"] == "free"
      and sj.build_adaptive_plan(dd(RD), [], SETS, EX)["kind"] == "week")
card_far = sj.adaptive_day_plan(p_far, 1, EX)
check("자유 연습 구간에서도 Day 1(D-7) 카드는 '예정'으로 세트를 보여 준다",
      card_far["제목"] == "예정" and card_far["세트"] == [NAMES[0]], card_far["제목"])
p_after = sj.build_adaptive_plan(EXAM + TD(days=1), [], SETS, EX)
check("시험일 지남: kind after · 배정 없음 · '시험일을 다시 설정' 안내",
      p_after["kind"] == "after" and p_after["days"] == {}
      and "시험일을 다시 설정" in p_after["reason"], p_after["reason"])
p_between = sj.build_adaptive_plan(EXAM + TD(days=1), [], SETS, [EXAM, EXAM + TD(days=3)])
check("v3.1.0 — 1차가 끝나면 곧바로 2차 기준 주간 일정이 선다 (지난 시험일 제외)",
      p_between["kind"] == "week" and sorted(p_between["days"])
      == [EXAM + TD(days=1), EXAM + TD(days=2)]
      and p_between["exam_dates"] == [EXAM, EXAM + TD(days=3)],
      str(sorted(p_between["days"])))
p_pastex = sj.build_adaptive_plan(dd(3), [], SETS, [EXAM - TD(days=60), EXAM])
check("v3.1.0 — 오래전 시험일이 남아 있어도 다음 시험일 기준으로 계산된다",
      p_pastex["kind"] == "week" and sorted(p_pastex["days"]) == WEEK[-3:],
      str(sorted(p_pastex["days"])))
p_unset = sj.build_adaptive_plan(dd(3), [], SETS, [])
check("시험일 미설정: kind unset · 배정 없음 · 시험 날짜 입력 안내 · exam_dates 빈 목록",
      p_unset["kind"] == "unset" and p_unset["days"] == {}
      and p_unset["attempts_left"] == 0 and p_unset["exam_dates"] == []
      and "시험일 미설정" in p_unset["reason"], p_unset["reason"])
check("시험일 미설정 카드는 고정 안내로 폴백(적응형 아님)",
      sj.adaptive_day_plan(p_unset, 1, [])["제목"] == "Day 1"
      and sj.adaptive_day_plan(p_unset, 0, [])["제목"] == "시험일 설정")
check("2차 시험일이 있어도 커리큘럼은 1차 기준 (D-7~D-1 그대로)",
      sorted(sj.build_adaptive_plan(dd(RD), [], SETS, EX2)["days"]) == WEEK)
check("엔진 결과에 시험일 목록이 담긴다(웹·직렬화가 그대로 쓴다)",
      sj.build_adaptive_plan(dd(3), [], SETS, EX2)["exam_dates"] == EX2)
check("v3.1.0 — 시험일 0개여도 죽지 않고 자유 연습(unset)으로 돈다",
      sj.build_adaptive_plan(dd(3), [], SETS, [])["kind"] == "unset"
      and sj.build_adaptive_plan(dd(3), [], [], [])["no_sets"] is True)
EX3 = [EXAM - TD(days=90), EXAM, EXAM + TD(days=40)]
check("v3.1.0 — 시험일 3개: 기준은 가장 가까운 미래, 지난 것은 목록에만 남는다",
      sj.effective_exams(EX3, dd(3)) == EX3[1:]
      and sj.dday_num(dd(3), EX3) == 3
      and sj.effective_exams(EX3, EXAM + TD(days=1)) == EX3[2:]
      and sj.dday_num(EXAM + TD(days=1), EX3) == 39
      and sj.effective_exams(EX3, EXAM + TD(days=100)) == EX3)

print()
print(f"적응형 일정 테스트 {N}건 전부 통과")
