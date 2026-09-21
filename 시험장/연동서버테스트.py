# -*- coding: utf-8 -*-
"""시험장.py v2.4.0 루틴 웹 연동 서버 헤드리스 테스트 (tkinter 불필요).

실제 RoutineServer 를 임의 포트에 띄워 규약(문서/연동_API.md v2.6.0)을 검증합니다.
  1. GET / — 루틴.html(실제 루틴.html)을 doctype/html/head/body 골격으로 감싸서
     응답, 해시→탭 스크립트 포함, 파일이 없으면 안내 페이지, Cache-Control no-store
  2. GET /api/state — 규약 필드 전부 존재·타입 (최상위·plan·days 16일·슬롯·sets·
     records·review·checks·settings), 웹 픽스처(fixture_state.json)와 키 집합 일치
  3. 쓰기 공통 — 토큰 없음/불일치 401, ?t= 토큰 허용, 본문 오류·필드 누락 400,
     Host 가 루프백이 아니면 403, 응답은 최신 state + ok
  4. POST /api/check → 세트설정 `_웹체크` 저장/삭제, generated 갱신
  5. POST /api/score → 기록.json 에 mode=수동 추가(일시 = date + 현재 시각, 세트명),
     알 수 없는 norm·범위 밖 total 400, 오늘 슬롯 소진 반영
  6. POST /api/review → 오답노트 완료 저장소(이해체크 전체 + 완료 일시), 해제, 없는 기록
  6-2. POST /api/exam_date → 시험일 저장/지우기(웹에서도 프로그램과 똑같이), 형식·순서 오류 400
  7. POST /api/action → 시험 중 start_exam/start_retry 오류, open_pdf 없음, show ok,
     알 수 없는 action 400, action_fn 호출 인자
  8. GET /api/report/<id> → 기록의 채점결과 HTML, 없는 id 404, 경로 탈출 차단
  9. 동시 요청(스레드 30개) 전부 200, 포트 충돌 시 대체 포트, generated 는 상태
     변경마다 반드시 달라짐(시험 시작/종료 토글 포함)
 10. 기록 id 규칙(같은 분 두 건 유일·안정), 채점 기록 wrong_items 필수(빈 배열),
     자동 업데이트 검증이 set_files 의 .html 을 받아들임
"""
import http.client
import importlib.util
import json
import os
import shutil
import sys
import tempfile
import threading
import time
import urllib.parse
from datetime import date, datetime

BASE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("sj", os.path.join(BASE, "시험장.py"))
sj = importlib.util.module_from_spec(spec)
spec.loader.exec_module(sj)
N = 0
TMP = tempfile.mkdtemp(prefix="coco_link_")
sj.STARTUP_LOG_PATH = os.path.join(TMP, "시작로그.txt")
sj.ERROR_LOG_PATH = os.path.join(TMP, "오류.log")
sj.RECORDS_DIR = os.path.join(TMP, "채점결과")
sj.RECORDS_PATH = os.path.join(sj.RECORDS_DIR, "기록.json")
sj.SET_CONFIG_PATH = os.path.join(TMP, "세트설정.json")
os.makedirs(sj.RECORDS_DIR, exist_ok=True)
TODAY = date(2026, 9, 6)
# 시험일은 사용자가 넣는 값이다(코드에 박혀 있지 않다). 픽스처는 2026-09-10 을
# 넣어 D-7 = 9/3, D-1 = 9/9, 오늘(9/6) = D-4 가 되게 한다.
EXAM_DATE = date(2026, 9, 10)
WEEK_START, WEEK_END = date(2026, 9, 3), date(2026, 9, 9)
REAL_HTML = os.path.join(os.path.dirname(BASE), "루틴.html")
# 웹(루틴 v6) 쪽 /api/state 픽스처와 키 집합을 맞춰 보는 선택 검사 — 경로는 환경 변수로
FIXTURE = os.environ.get("COCO_WEB_FIXTURE") or os.path.join(BASE, "fixture_state.json")


def check(desc, cond, extra=""):
    global N
    N += 1
    print(f"  [{N:02d}] {desc}: {'OK' if cond else 'FAIL'}" + (f" ({extra})" if extra else ""))
    if not cond:
        raise AssertionError(f"{desc} {extra}")


def wjson(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)


def rjson(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# 픽스처: 세트 3개(PDF 있음/없음), 채점결과 JSON·HTML, 기록 5건(시험·부분연습·오답재풀이·같은 분 2건)
# ---------------------------------------------------------------------------
SETDIR = os.path.join(TMP, "모의고사")
os.makedirs(os.path.join(SETDIR, "채점결과"), exist_ok=True)


def mkset(name, pdf=True, key=True):
    prob = os.path.join(SETDIR, f"{name}_문제.xlsx")
    ans = os.path.join(SETDIR, f"{name}_정답.xlsx")
    for p in (prob, ans):
        open(p, "wb").write(b"PK\x03\x04")
    pdfp = None
    if pdf:
        pdfp = os.path.join(SETDIR, f"{name}_문제지.pdf")
        open(pdfp, "wb").write(b"%PDF-1.4")
    return {"name": name, "norm": sj.norm_set_key(name), "dir": SETDIR,
            "problem": prob, "answer": ans,
            "key": os.path.join(SETDIR, f"{name}_기대값.json") if key else None,
            "pdf": pdfp}


A = mkset("2024년 상공회의소 샘플 A형")
B = mkset("2024년 상공회의소 샘플 B형")
C = mkset("2026 컴활 2급 실기 상시 복원 기출 1회", pdf=False, key=False)
SETS = [A, B, C]


def mkresult(name, stamp, total, wrong, partial=False):
    base = os.path.join(SETDIR, "채점결과", f"채점결과_{name}_{stamp}")
    data = {"total": total, "pass_line": 70, "passed": total >= 70,
            "mode": "partial" if partial else "full",
            "graded_sheets": ["계산작업"] if partial else ["기본작업-1", "계산작업"],
            "max_total": 40 if partial else 100,
            "files": {"problem": "p", "answer": "a", "student": "s"},
            "sheets": [{"name": "기본작업-1", "alloc": 5, "earned": 5, "missing": [],
                        "details": [], "notes": []},
                       {"name": "계산작업", "alloc": 40, "earned": 40 - sum(w[2] for w in wrong),
                        "missing": [], "details": [], "notes": []}],
            "wrong_items": [{"sheet": w[0], "label": w[1], "lost": w[2], "kind": "value",
                             "category": w[3], "cells": [], "explain": []} for w in wrong]}
    wjson(base + ".json", data)
    with open(base + ".html", "w", encoding="utf-8") as f:
        f.write(f"<!doctype html><html><body><h1>채점결과 {name} {total}점</h1></body></html>")
    return base + ".html"


HTML_A = mkresult(A["name"], "20260905_1015", 58,
                  [("계산작업", "계산작업 1번", 8, "계산작업"),
                   ("계산작업", "계산작업 2번", 8, None)])      # category None → 시트로 보완
HTML_B = mkresult(B["name"], "20260906_1012", 72, [("계산작업", "계산작업 3번", 8, "계산작업")])
HTML_B_RETRY = mkresult(B["name"], "20260906_1140", 8, [], partial=True)
HTML_NOWRONG = mkresult(A["name"], "20260904_2100", 100, [])
RECORDS = [
    {"일시": "2026-09-04 21:00", "세트명": A["name"], "점수": 24, "소요시간": "12:00",
     "리포트": os.path.join(TMP, "없는리포트.html"), "mode": "부분연습", "영역": "계산작업"},
    {"일시": "2026-09-04 21:00", "세트명": A["name"] + " (연습)", "점수": None,
     "소요시간": "01:00", "리포트": None, "mode": "부분연습", "영역": "계산작업"},   # 같은 분 두 번째(채점 실패)
    {"일시": "2026-09-05 10:15", "세트명": A["name"], "점수": 58, "소요시간": "39:10",
     "리포트": HTML_A, "mode": "시험", "day": "d03", "루틴": "2026-09"},
    {"일시": "2026-09-06 10:12", "세트명": B["name"], "점수": 72, "소요시간": "40:00",
     "리포트": HTML_B, "mode": "시험"},
    {"일시": "2026-09-06 11:40", "세트명": B["name"], "점수": 8, "소요시간": "10:00",
     "리포트": HTML_B_RETRY, "mode": "오답재풀이", "영역": "오답재풀이(계산작업)", "만점": 40},
]
wjson(sj.RECORDS_PATH, RECORDS)
wjson(sj.SET_CONFIG_PATH, {"_설정": {"루틴자동열기": True, "자동업데이트": False,
                                   "시험일": EXAM_DATE.isoformat(), "시험일2": None},
                          "_웹체크": {"e01": True, "d00t1": True}})
# 루틴 페이지: 실제 루틴.html 을 그대로 서빙 (없으면 더미)
PAGE = os.path.join(TMP, "루틴.html")
if os.path.isfile(REAL_HTML):
    shutil.copy2(REAL_HTML, PAGE)
else:
    with open(PAGE, "w", encoding="utf-8") as f:
        f.write('<title>더미</title><div class="tabbar"><button role="tab" id="tab-quiz">퀴즈</button></div>')

ACTIONS = []
EXAM = {"running": False}


def action_fn(action, s, rec, norm, rid):
    ACTIONS.append((action, s["norm"] if s else None, rec["json"] if rec else None, norm, rid))
    if action == "start_exam":
        EXAM["running"] = True
        sj.routine_touch()
    return True, None


CHANGES = []
srv = sj.RoutineServer(sets_fn=lambda: SETS, action_fn=action_fn,
                       on_change=lambda k, i: CHANGES.append((k, i)),
                       exam_running_fn=lambda: EXAM["running"], port=0,
                       html_path=PAGE, today=TODAY)
srv.start()
PORT = srv.port


def req(method, path, body=None, token=None, headers=None, host=None, raw=None):
    conn = http.client.HTTPConnection("127.0.0.1", PORT, timeout=10)
    hdr = {"Host": host or f"127.0.0.1:{PORT}"}
    if headers:
        hdr.update(headers)
    data = None
    if raw is not None:
        data = raw
        hdr.setdefault("Content-Type", "application/json")
    elif body is not None:
        data = json.dumps(body, ensure_ascii=False).encode("utf-8")
        hdr.setdefault("Content-Type", "application/json")
    if token:
        hdr["X-Coco-Token"] = token
    conn.request(method, urllib.parse.quote(path, safe="/?=&%#"), body=data, headers=hdr)
    resp = conn.getresponse()
    text = resp.read().decode("utf-8", "replace")
    conn.close()
    try:
        j = json.loads(text)
    except ValueError:
        j = None
    return resp.status, dict(resp.getheaders()), text, j


TOK = srv.token
print(f"서버: 127.0.0.1:{PORT} 페이지={'실제 루틴.html' if os.path.isfile(REAL_HTML) else '더미'}")

# ---------------------------------------------------------------------------
# 1. GET /
# ---------------------------------------------------------------------------
st, hd, page, _ = req("GET", f"/?t={TOK}")
check("GET / 200 + text/html + no-store", st == 200 and hd.get("Content-Type", "").startswith("text/html")
      and hd.get("Cache-Control") == "no-store", f"{st} {hd.get('Content-Type')}")
check("GET / 응답이 <!doctype html><html lang=\"ko\"> 골격으로 시작", page.startswith('<!doctype html><html lang="ko"><head><meta charset="utf-8">'), page[:60])
check("GET / 본문에 루틴 페이지 내용(제목·탭) 포함", "<title>" in page and 'id="tab-quiz"' in page)
check("GET / 끝에 해시→탭 스크립트 + </body></html>", page.rstrip().endswith("</body></html>") and 'getElementById("tab-"+h)' in page)
check("GET / 는 매 요청마다 파일을 읽음(파일 수정 즉시 반영)", True)
with open(PAGE, "a", encoding="utf-8") as f:
    f.write("<!-- 갱신표식 -->")
st, hd, page2, _ = req("GET", "/")
check("루틴.html 갱신 후 재시작 없이 새 내용 서빙", "갱신표식" in page2 and "갱신표식" not in page)
os.rename(PAGE, PAGE + ".bak")
st, hd, miss, _ = req("GET", "/")
check("루틴.html 없으면 안내 페이지(200, doctype, 파일명·경로 안내)", st == 200 and miss.startswith("<!doctype html>")
      and "루틴.html" in miss and "업데이트 확인" in miss and TMP in miss)
os.rename(PAGE + ".bak", PAGE)
with open(PAGE, "w", encoding="utf-8") as f:
    f.write("<!DOCTYPE html><html><body>이미 골격 있음</body></html>")
st, hd, page3, _ = req("GET", "/")
check("이미 <!doctype 으로 시작하는 파일은 감싸지 않음", page3 == "<!DOCTYPE html><html><body>이미 골격 있음</body></html>")
shutil.copy2(REAL_HTML, PAGE) if os.path.isfile(REAL_HTML) else None
st, _h, _t, j = req("GET", "/favicon.ico")
check("favicon 204", st == 204)
st, _h, _t, j = req("GET", "/없는경로")
check("없는 경로 404 JSON", st == 404 and j and j.get("ok") is False)

# ---------------------------------------------------------------------------
# 2. GET /api/state 스키마
# ---------------------------------------------------------------------------
st, hd, text, S = req("GET", "/api/state")
check("GET /api/state 200 + application/json; charset=utf-8 + no-store",
      st == 200 and hd.get("Content-Type") == "application/json; charset=utf-8" and hd.get("Cache-Control") == "no-store")
check("JSON 은 ensure_ascii=False (한글 그대로)", "상공회의소" in text)
TOP = {"version": str, "today": str, "generated": str, "exam_running": bool, "plan": dict,
       "sets": list, "records": list, "review": dict, "checks": dict, "settings": dict,
       "corrections": list, "achievements": list,      # 2.6.0 이의제기 반영 · 도전 과제
       "bands": list,                                   # 2.7.0 점수대 보드 요약
       "exam": dict}                                    # 3.0.0 사용자 시험일
check("최상위 필드·타입", all(isinstance(S.get(k), t) for k, t in TOP.items()), str({k: type(S.get(k)).__name__ for k in TOP}))
check("version=3.0.0 · today=프로그램 기준(2026-09-06) · exam_running=false", S["version"] == "3.0.0" and S["today"] == "2026-09-06" and S["exam_running"] is False)
check("exam: 사용자가 넣은 시험일 · D-day · 루틴 길이 7 · 시작일(D-7) · 기준/지난 시험일(3.1.0)",
      S["exam"] == {"dates": ["2026-09-10"], "active": ["2026-09-10"],
                    "past": [], "next": "2026-09-10", "set": True, "dday": 4,
                    "routine_days": 7, "start": "2026-09-03"}, json.dumps(S["exam"], ensure_ascii=False))
check("최상위 게임 요소: level{tier,current_goal,best_total,achieved,next_goal} · streak{current,best,today_done} · attempts_left(= 남은 학습일) · today_locked(2.7.0: 항상 false)",
      S["level"] == {"tier": 80, "current_goal": 80, "best_total": 72, "achieved": [70], "next_goal": 80}
      and S["streak"] == {"current": 2, "best": 2, "today_done": True}
      and S["attempts_left"] == 4 and S["today_locked"] is False, json.dumps(S["level"], ensure_ascii=False))
check("corrections: 정정 없는 기록만 있으면 빈 목록(웹은 아무것도 그리지 않음)",
      S["corrections"] == [], S["corrections"])
ACH_KEYS = {"id", "name", "desc", "icon", "group", "done", "when"}
check("achievements: 12개 고정 순서 · 필드 · 달성 판정(첫 완주·기본무실점은 달성, 만점·5일 연속은 미달)",
      len(S["achievements"]) == 12 and all(set(a) == ACH_KEYS for a in S["achievements"])
      and [a["id"] for a in S["achievements"]][:3] == ["first", "pass70", "tier80"]
      and {a["id"]: a["done"] for a in S["achievements"]}["first"] is True
      and {a["id"]: a["done"] for a in S["achievements"]}["tier100"] is False
      and {a["id"]: a["done"] for a in S["achievements"]}["streak5"] is False,
      str([(a["id"], a["done"], a["when"]) for a in S["achievements"]]))
check("generated 는 마이크로초 ISO 일시", len(S["generated"]) == 26 and datetime.fromisoformat(S["generated"]) is not None, S["generated"])
P = S["plan"]
PLAN_KEYS = {"kind", "seg", "today", "start", "end", "seg_start", "reason", "warning", "remaining", "retry_target",
             "missed_days", "boost_days", "demoted", "days", "preview",
             # 2.5.0 추가 (스키마 호환 — 기존 필드는 그대로)
             "level", "streak", "attempts_left", "today_locked", "passed_sets",
             "total_sets", "all_clear", "retry_queue",
             # 3.0.0 추가 (사용자 시험일 기준 1주 · 보유 세트)
             "exam_dates", "exam_set", "dday", "no_sets", "routine_days"}
check("plan 필드 전부", PLAN_KEYS <= set(P), str(PLAN_KEYS - set(P)))
check("plan: week·seg 1·start 9/6(오늘)·end 9/9(D-1)·seg_start 9/3(D-7)", P["kind"] == "week" and P["seg"] == 1 and P["start"] == "2026-09-06"
      and P["end"] == "2026-09-09" and P["seg_start"] == "2026-09-03")
check("plan: 시험일·D-day·루틴 길이·보유 세트 있음",
      P["exam_dates"] == ["2026-09-10"] and P["exam_set"] is True
      and P["dday"] == 4 and P["routine_days"] == 7 and P["no_sets"] is False)
check("plan: missed_days 9/3·9/4(ISO) · reason 은 목표·남은 학습일 한 줄 · remaining 정수",
      P["missed_days"] == ["2026-09-03", "2026-09-04"] and isinstance(P["reason"], str)
      and P["reason"] == "현재 목표 80점 · 남은 학습일 4일" and isinstance(P["remaining"], int), P["reason"])
check("plan: 승격·강등·경고·재응시 최소 폐지 → boost_days·demoted 빈 목록 · warning null · retry_target 0",
      P["boost_days"] == [] and P["demoted"] == [] and P["warning"] is None and P["retry_target"] == 0)
check("plan 게임 요소는 최상위와 같은 값 · 진행률용 합격/전체 세트 · all_clear",
      P["level"] == S["level"] and P["streak"] == S["streak"] and P["attempts_left"] == 4
      and P["today_locked"] is False and P["passed_sets"] == 1 and P["total_sets"] == 3
      and P["all_clear"] is False, str((P["passed_sets"], P["total_sets"])))
check("plan.retry_queue: 재응시 후보 대기줄 = A형 58(retry) → B형 72(goal, 목표 80 도전)",
      [(q["set"], q["kind"], q["best"]) for q in P["retry_queue"]]
      == [(A["norm"], "retry", 58), (B["norm"], "goal", 72)], json.dumps(P["retry_queue"], ensure_ascii=False))
D = P["days"]
check("days: D-7(9/3) ~ 시험일(9/10) 8일 오름차순", [d["date"] for d in D] == [f"2026-09-{n:02d}" for n in range(3, 11)])
DAY_KEYS = {"date", "no", "kind", "title", "capacity", "promoted", "deadline", "review_day", "free", "slots"}
check("days 항목 필드", all(set(d) == DAY_KEYS for d in D), str(set(D[0])))
kinds = {d["date"]: d["kind"] for d in D}
check("kind: 9/3~9/5 past · 9/10 exam · next/review 없음(단일 구간)", all(kinds[f"2026-09-{n:02d}"] == "past" for n in (3, 4, 5))
      and kinds["2026-09-10"] == "exam" and not any(k in ("next", "review") for k in kinds.values()), str(kinds))
check("kind: 9/6~9/9 는 mock 또는 rest", all(kinds[f"2026-09-{n:02d}"] in ("mock", "rest") for n in range(6, 10)), str(kinds))
check("past·exam·next 의 slots 는 []", all(d["slots"] == [] for d in D if d["kind"] in ("past", "exam", "next")))
check("no: Day 번호(9/3=1 … 9/9=7, 9/10=시험일 8)", [d["no"] for d in D] == list(range(1, 8)) + [sj.PLAN_EXAM1])
check("deadline 9/9(D-1) · review_day 전부 false · exam capacity 0 · 학습일 capacity 는 주말·평일 없이 전부 1",
      [d["date"] for d in D if d["deadline"]] == ["2026-09-09"] and not any(d["review_day"] for d in D)
      and all(d["capacity"] == 0 for d in D if d["kind"] == "exam") and kinds["2026-09-06"] in ("mock", "rest")
      and all(d["capacity"] == 1 for d in D if d["kind"] != "exam")
      and not any(d["promoted"] for d in D))
check("free: 배정된 날은 0 (오늘 이미 응시했어도 추천이 남으므로 0) · 배정도 후보도 없으면 1(자유 복습)",
      all(d["free"] == 0 for d in D if d["slots"]) and [d["free"] for d in D if d["date"] == "2026-09-06"] == [0]
      and all(d["free"] == 0 for d in D if d["kind"] in ("past", "exam")))
check("past title: 9/3 '응시 없음'(경고·재배치 아님) · 9/5 완료(A형 기록)",
      [d["title"] for d in D if d["date"] in ("2026-09-03", "2026-09-05")] == ["응시 없음", "완료"])
check("오늘(9/6) title: 이미 응시했어도 추천은 남는다 → '추천 1세트'",
      [d["title"] for d in D if d["date"] == "2026-09-06"] == ["추천 1세트"])
today_day = [d for d in D if d["date"] == "2026-09-06"][0]
SLOT_KEYS = {"kind", "set", "goal", "why", "best", "counts_as_retry"}
all_slots = [sl for d in D for sl in d["slots"]]
check("슬롯 필드(kind/set/goal/why/best/counts_as_retry, done·record_id 선택)", all_slots and all(SLOT_KEYS <= set(sl) for sl in all_slots), str(all_slots[:1]))
check("슬롯 kind ∈ first|retry|goal · set 은 {name,norm,pdf} (2.5.0 은 자동 선택 슬롯을 쓰지 않음)",
      all(sl["kind"] in ("first", "retry", "goal") for sl in all_slots)
      and all(sl["set"] is not None and set(sl["set"]) == {"name", "norm", "pdf"} for sl in all_slots))
check("오늘(9/6)은 이미 응시(B형 72)했지만 제한이 없어 추천 슬롯(C형 첫 응시)이 그대로 · kind mock",
      len(today_day["slots"]) == 1 and today_day["kind"] == "mock"
      and today_day["slots"][0]["set"]["norm"] == C["norm"],
      json.dumps(today_day, ensure_ascii=False)[:200])
check("counts_as_retry 는 2.5.0 에서 쓰지 않음(항상 false, 필드는 유지)",
      all(sl["counts_as_retry"] is False for sl in all_slots))
retry_slots = [sl for sl in all_slots if sl["kind"] == "retry" and sl["set"]]
check("재응시 슬롯: goal = 현재 목표 80 · best = 그 세트 최고점(A형 58) · 사유에 '합격선 70 미달'",
      retry_slots and all(sl["goal"] == 80 for sl in retry_slots)
      and any(sl["set"]["norm"] == A["norm"] and sl["best"] == 58 and "합격선 70 미달" in sl["why"] for sl in retry_slots),
      json.dumps(retry_slots, ensure_ascii=False)[:300])
goal_slots = [sl for sl in all_slots if sl["kind"] == "goal"]
check("목표 도전 슬롯: 합격(72)한 B형도 현재 목표 80 미달이라 다시 배정 · 사유 '목표 80점'",
      goal_slots and all(sl["set"]["norm"] == B["norm"] and sl["best"] == 72
                         and "목표 80점" in sl["why"] for sl in goal_slots),
      json.dumps(goal_slots, ensure_ascii=False)[:300])
check("preview: 다음 학습 구간이 없으므로 null (필드는 유지)", "preview" in P and P["preview"] is None)
by_date_slots = {d["date"]: d["slots"] for d in D}
check("9/6 C 첫 응시 → 9/7부터 A형(58)·B형(72) 번갈아 재응시 · 하루 1슬롯 · 목표 전부 80",
      [[sl["set"]["norm"] for sl in by_date_slots[f"2026-09-{n:02d}"]] for n in range(6, 10)]
      == [[C["norm"]], [A["norm"]], [B["norm"]], [A["norm"]]]
      and all(sl["goal"] == 80 for v in by_date_slots.values() for sl in v)
      and len(by_date_slots["2026-09-09"]) == 1,
      str({k: [(sl["set"]["norm"], sl["goal"]) for sl in v] for k, v in by_date_slots.items()}))
# sets
SET_KEYS = {"name", "norm", "pdf", "key", "attempts", "best", "review_done", "retry_done",
            # 2.7.0 점수대 보드
            "band", "scores", "last_date", "prev_best", "prev_band", "best_delta", "improved"}
SS = {s["norm"]: s for s in S["sets"]}
check("sets: 3개 · 필드 전부", len(S["sets"]) == 3 and all(set(s) == SET_KEYS for s in S["sets"]))
check("sets: A형 pdf true·key true·attempts 1(부분연습 제외)·best 58 / C pdf false·key false·attempts 0·best null",
      SS[A["norm"]]["pdf"] and SS[A["norm"]]["key"] and SS[A["norm"]]["attempts"] == 1 and SS[A["norm"]]["best"] == 58
      and SS[C["norm"]]["pdf"] is False and SS[C["norm"]]["key"] is False and SS[C["norm"]]["attempts"] == 0 and SS[C["norm"]]["best"] is None)
check("sets: B형 attempts 1(부분연습·오답재풀이 제외)·best 72·retry_done true(이후 오답재풀이)", SS[B["norm"]]["attempts"] == 1
      and SS[B["norm"]]["best"] == 72 and SS[B["norm"]]["retry_done"] is True and SS[B["norm"]]["review_done"] is False, str(SS[B["norm"]]))
# 2.7.0 점수대 보드 (sets[] 밴드 필드 + 최상위 bands 요약)
check("sets 밴드: A형 58 → fail · B형 72 → pass · C 미응시 → none",
      (SS[A["norm"]]["band"], SS[B["norm"]]["band"], SS[C["norm"]]["band"])
      == ("fail", "pass", "none"),
      str({k: v["band"] for k, v in SS.items()}))
check("sets 점수 이력·마지막 응시일 (부분연습·오답재풀이 제외)",
      SS[A["norm"]]["scores"] == [58] and SS[A["norm"]]["last_date"] == "2026-09-05"
      and SS[B["norm"]]["scores"] == [72] and SS[B["norm"]]["last_date"] == "2026-09-06"
      and SS[C["norm"]]["scores"] == [] and SS[C["norm"]]["last_date"] is None,
      str([SS[A["norm"]]["scores"], SS[B["norm"]]["scores"]]))
check("sets 갱신 필드: 첫 응시면 prev_best null · improved true · best_delta null",
      SS[A["norm"]]["prev_best"] is None and SS[A["norm"]]["prev_band"] == "none"
      and SS[A["norm"]]["improved"] is True and SS[A["norm"]]["best_delta"] is None
      and SS[C["norm"]]["improved"] is False)
BANDS = {b["id"]: b for b in S["bands"]}
check("bands: 6개 고정 순서(none→fail→pass→skilled→high→perfect) · {id,name,desc,count,sets}",
      [b["id"] for b in S["bands"]] == ["none", "fail", "pass", "skilled", "high", "perfect"]
      and all(set(b) == {"id", "name", "desc", "count", "sets"} for b in S["bands"]),
      str([b["id"] for b in S["bands"]]))
check("bands 개수·세트 묶음: 미응시 C · 미달 A형 · 합격 B형 · 나머지 0",
      [b["count"] for b in S["bands"]] == [1, 1, 1, 0, 0, 0]
      and BANDS["none"]["sets"] == [C["norm"]] and BANDS["fail"]["sets"] == [A["norm"]]
      and BANDS["pass"]["sets"] == [B["norm"]],
      str({b["id"]: b["sets"] for b in S["bands"]}))
check("bands 이름 = 규약 SCORE_BANDS 와 같음 · 요약 문구",
      [b["name"] for b in S["bands"]] == ["미응시", "미달", "합격", "숙련", "고득점", "만점"]
      and sj.band_summary_text(S["bands"]) == "미응시 1 · 미달 1 · 합격 1",
      sj.band_summary_text(S["bands"]))
check("bands·sets 는 score_board(순수 함수) 결과와 같다 (웹 폴백과 같은 정의)",
      sj.score_board(SETS, S["records"])["bands"] == S["bands"]
      and all(SS[r["norm"]]["band"] == r["band"] and SS[r["norm"]]["scores"] == r["scores"]
              for r in sj.score_board(SETS, S["records"])["sets"]))
# records
REC_KEYS = {"id", "date", "time", "set", "mode", "total", "pass_line", "passed", "report_html", "report_json", "sheets", "wrong_items",
            "elapsed", "elapsed_sec"}   # 2.6.0: 소요시간(게임형 '40분 내 완주' 판정)
R = S["records"]
check("records: 5건 전부(부분연습·오답재풀이 포함) · 필드 전부", len(R) == 5 and all(set(r) == REC_KEYS for r in R))
ids = [r["id"] for r in R]
check("id: 일시 14자리 · 같은 분 두 건은 초 자리로 구분 · 유일", ids == ["20260904210000", "20260904210001", "20260905101500", "20260906101200", "20260906114000"], str(ids))
st2, _h, _t, S2 = req("GET", "/api/state")
check("id: 다시 조회해도 같은 id(안정)", [r["id"] for r in S2["records"]] == ids)
rA = R[2]
check("A형 기록: elapsed '39:10' → elapsed_sec 2350 (수동 기록은 0)",
      R[2]["elapsed"] == "39:10" and R[2]["elapsed_sec"] == 2350, str((R[2]["elapsed"], R[2]["elapsed_sec"])))
check("A형 기록: date/time/set{name,norm}/mode 시험/total 58/pass_line 70/passed false", rA["date"] == "2026-09-05" and rA["time"] == "10:15"
      and rA["set"] == {"name": A["name"], "norm": A["norm"]} and rA["mode"] == "시험" and rA["total"] == 58 and rA["pass_line"] == 70 and rA["passed"] is False)
check("A형 기록: report_html/json 파일명 · sheets {name,alloc,earned} · wrong_items {label,category,lost}(category None → 시트 보완)",
      rA["report_html"] == os.path.basename(HTML_A) and rA["report_json"] == os.path.basename(HTML_A)[:-5] + ".json"
      and rA["sheets"][0] == {"name": "기본작업-1", "alloc": 5, "earned": 5}
      and [w["category"] for w in rA["wrong_items"]] == ["계산작업", "계산작업"] and rA["wrong_items"][0]["lost"] == 8 and rA["wrong_items"][0]["label"] == "계산작업 1번", str(rA))
check("리포트 파일이 없는 기록: report_html/json null · sheets/wrong_items []", R[0]["report_html"] is None and R[0]["report_json"] is None and R[0]["sheets"] == [] and R[0]["wrong_items"] == [])
check("' (연습)' 접미 기록: 세트명 정리 · total null · passed false", R[1]["set"]["name"] == A["name"] and R[1]["total"] is None and R[1]["passed"] is False)
check("오답재풀이 기록: mode 오답재풀이 · 채점 기록이라 wrong_items [] 포함(오답 없음)", R[4]["mode"] == "오답재풀이" and R[4]["report_json"] and R[4]["wrong_items"] == [])
check("B형 72점 기록 passed true", R[3]["passed"] is True)
# review / checks / settings
check("review: 채점 JSON 이 있는 기록마다 {done, when}", set(S["review"]) == {"20260905101500", "20260906101200", "20260906114000"}
      and all(set(v) == {"done", "when"} for v in S["review"].values()) and S["review"]["20260905101500"]["done"] is False)
check("checks: _웹체크 그대로 · settings(자동 열기·자동 업데이트·시험일·스캔 폴더·숨긴 세트) 그대로",
      S["checks"] == {"e01": True, "d00t1": True}
      and S["settings"] == {"auto_open_routine": True, "auto_update": False,
                            "scan_root": "", "hidden_sets": [],
                            "exam_dates": ["2026-09-10"],
                            "exam_date": "2026-09-10", "exam_date2": None},
      json.dumps(S["settings"], ensure_ascii=False))
if os.path.isfile(FIXTURE):
    FX = rjson(FIXTURE)
    fday = FX["plan"]["days"][0]
    fslot = next(sl for d in FX["plan"]["days"] for sl in d["slots"])
    check("웹 픽스처와 키 집합 일치(최상위·plan·day·slot·set·record)", set(FX) <= set(S) and set(FX["plan"]) == set(P)
          and set(fday) == DAY_KEYS and set(fslot) <= set(all_slots[0]) and set(FX["sets"][0]) == SET_KEYS and set(FX["records"][0]) == REC_KEYS,
          str((set(FX) - set(S), set(FX["plan"]) ^ set(P))))

# ---------------------------------------------------------------------------
# 3. 쓰기 공통: 토큰·본문·Host
# ---------------------------------------------------------------------------
st, _h, _t, j = req("POST", "/api/check", {"key": "x", "value": True})
check("토큰 없음 → 401 {ok:false,error}", st == 401 and j == {"ok": False, "error": "토큰 불일치"}, f"{st} {j}")
st, _h, _t, j = req("POST", "/api/check", {"key": "x", "value": True}, token="wrong-token")
check("토큰 불일치 → 401", st == 401 and j["ok"] is False)
st, _h, _t, j = req("POST", f"/api/check?t={TOK}", {"key": "e02", "value": True})
check("?t= 쿼리 토큰 허용 → 200 ok:true + 최신 state", st == 200 and j["ok"] is True and j["checks"].get("e02") is True and "plan" in j and j["today"] == "2026-09-06")
st, _h, _t, j = req("POST", "/api/check", raw="{깨진".encode("utf-8"), token=TOK)
check("본문 JSON 파싱 실패 → 400", st == 400 and j["ok"] is False and "JSON" in j["error"])
st, _h, _t, j = req("POST", "/api/check", {"value": True}, token=TOK)
check("필수 필드(key) 누락 → 400", st == 400 and j["ok"] is False)
st, _h, _t, j = req("POST", "/api/check", raw=b"[1,2]", token=TOK)
check("JSON 객체가 아니면 400", st == 400)
st, _h, _t, j = req("POST", "/api/없음", {"a": 1}, token=TOK)
check("없는 API 경로 POST → 404", st == 404 and j["ok"] is False)
st, _h, _t, j = req("GET", "/api/state", host="evil.example.com")
check("Host 가 루프백이 아니면 403 (DNS 리바인딩 방지)", st == 403 and j["ok"] is False)
st, _h, _t, j = req("GET", "/api/state", host=f"localhost:{PORT}")
check("Host localhost 허용", st == 200 and j["today"] == "2026-09-06")

# ---------------------------------------------------------------------------
# 4. /api/check
# ---------------------------------------------------------------------------
g0 = req("GET", "/api/state")[3]["generated"]
st, _h, _t, j = req("POST", "/api/check", {"key": "2026-09-06|quiz", "value": True}, token=TOK)
cfg = rjson(sj.SET_CONFIG_PATH)
check("check true → 세트설정 _웹체크 저장 + 응답 checks 반영 + generated 변경", st == 200 and j["ok"] and cfg["_웹체크"].get("2026-09-06|quiz") is True
      and j["checks"]["2026-09-06|quiz"] is True and j["generated"] != g0 and j["generated"] > g0, str(cfg.get("_웹체크")))
check("on_change('check') 호출", CHANGES and CHANGES[-1] == ("check", {"key": "2026-09-06|quiz", "value": True}), str(CHANGES[-1:]))
g1 = j["generated"]
st, _h, _t, j = req("POST", "/api/check", {"key": "2026-09-06|quiz", "value": False}, token=TOK)
cfg = rjson(sj.SET_CONFIG_PATH)
check("check false → 키 삭제 · generated 또 변경", j["ok"] and "2026-09-06|quiz" not in cfg["_웹체크"] and "2026-09-06|quiz" not in j["checks"] and j["generated"] > g1)
check("기존 체크(e01·d00t1·e02) 유지", cfg["_웹체크"] == {"e01": True, "d00t1": True, "e02": True}, str(cfg["_웹체크"]))
st, _h, _t, j = req("POST", "/api/check", {"key": "", "value": True}, token=TOK)
check("빈 key → 400", st == 400)

# ---------------------------------------------------------------------------
# 5. /api/score
# ---------------------------------------------------------------------------
before = len(rjson(sj.RECORDS_PATH))
st, _h, _t, j = req("POST", "/api/score", {"date": "2026-09-06", "set": C["norm"], "total": 66}, token=TOK)
recs = rjson(sj.RECORDS_PATH)
new = recs[-1]
check("score → 기록.json 에 mode=수동 1건 추가(세트명=norm 의 세트 이름, 점수 66)", st == 200 and j["ok"] and len(recs) == before + 1
      and new["mode"] == "수동" and new["세트명"] == C["name"] and new["점수"] == 66 and new["리포트"] is None, str(new))
check("수동 기록 일시 = date + 현재 시각(HH:MM) · day 태그 d04 · 루틴 표식", new["일시"].startswith("2026-09-06 ") and len(new["일시"]) == 16
      and new["day"] == "d04" and new["루틴"] == EXAM_DATE.isoformat(), new["일시"])
rn = [r for r in j["records"] if r["mode"] == "수동"]
check("응답 records 에 수동 기록(id 14자리·sheets/wrong_items []·report null)", len(rn) == 1 and len(rn[0]["id"]) == 14
      and rn[0]["sheets"] == [] and rn[0]["wrong_items"] == [] and rn[0]["report_json"] is None and rn[0]["set"]["norm"] == C["norm"] and rn[0]["passed"] is False)
tday = [d for d in j["plan"]["days"] if d["date"] == "2026-09-06"][0]
check("수동 기록 후에도 잠기지 않는다 → 오늘 추천 슬롯 유지 · kind mock · 남은 학습일 4 유지",
      len(tday["slots"]) == 1 and tday["kind"] == "mock"
      and j["plan"]["today_locked"] is False
      and j["plan"]["attempts_left"] == 4 and j["plan"]["remaining"] == 4,
      str((tday["kind"], j["plan"]["remaining"], j["plan"]["attempts_left"])))
check("수동 66점 세트 C 는 합격선 미달 → 재응시 대기줄 1순위(최고점 낮은 순), 첫응시 풀은 비었음",
      [(q["set"], q["kind"], q["best"]) for q in j["plan"]["retry_queue"]]
      == [(A["norm"], "retry", 58), (C["norm"], "retry", 66), (B["norm"], "goal", 72)]
      and not any(sl["kind"] == "first" for d in j["plan"]["days"] for sl in d["slots"]),
      json.dumps(j["plan"]["retry_queue"], ensure_ascii=False))
check("sets: C attempts 1 · best 66 (수동 기록 포함)", [s for s in j["sets"] if s["norm"] == C["norm"]][0]["attempts"] == 1
      and [s for s in j["sets"] if s["norm"] == C["norm"]][0]["best"] == 66)
check("on_change('score') 호출", CHANGES[-1][0] == "score" and CHANGES[-1][1]["set"] is C)
for body, why in (({"date": "2026-09-06", "set": "없는norm", "total": 50}, "알 수 없는 norm"),
                  ({"date": "2026-09-06", "set": C["norm"], "total": 101}, "total 101"),
                  ({"date": "2026-09-06", "set": C["norm"], "total": -1}, "total -1"),
                  ({"date": "2026-09-06", "set": C["norm"], "total": 70.5}, "total 70.5"),
                  ({"date": "2026-09-06", "set": C["norm"], "total": "70"}, "total 문자열"),
                  ({"date": "2026/09/06", "set": C["norm"], "total": 70}, "date 형식"),
                  ({"date": "2026-09-06", "total": 70}, "set 누락")):
    st, _h, _t, j = req("POST", "/api/score", body, token=TOK)
    check(f"score 잘못된 요청({why}) → 400", st == 400 and j["ok"] is False, f"{st} {j}")
check("잘못된 요청은 기록에 안 남음", len(rjson(sj.RECORDS_PATH)) == before + 1)

# ---------------------------------------------------------------------------
# 6. /api/review
# ---------------------------------------------------------------------------
st, _h, _t, j = req("POST", "/api/review", {"record_id": "20260905101500", "done": True}, token=TOK)
cfg = rjson(sj.SET_CONFIG_PATH)
rv = cfg[A["norm"]]["오답연습"]
jn = os.path.basename(HTML_A)[:-5] + ".json"
check("review done → 이해체크 전체(0,1) + 완료 일시 저장(프로그램 오답노트 창과 같은 저장소)", st == 200 and j["ok"]
      and rv["이해체크"][jn] == [0, 1] and rv["완료"][jn], str(rv))
check("응답 review[id] = {done:true, when}", j["review"]["20260905101500"]["done"] is True and j["review"]["20260905101500"]["when"] == rv["완료"][jn])
check("sets: A형 review_done true(최근 기록 기준)", [s for s in j["sets"] if s["norm"] == A["norm"]][0]["review_done"] is True)
check("프로그램 쪽 load_review_state 로도 같은 체크가 보임", sj.load_review_state(A["norm"], jn, sj.SET_CONFIG_PATH) == {0, 1})
st, _h, _t, j = req("POST", "/api/review", {"record_id": "20260905101500", "done": False}, token=TOK)
rv = rjson(sj.SET_CONFIG_PATH)[A["norm"]]["오답연습"]
check("review done:false → 체크·완료 해제", j["ok"] and rv["이해체크"][jn] == [] and jn not in rv["완료"] and j["review"]["20260905101500"]["done"] is False)
# 프로그램 오답노트 창 경로: 전체 체크 → 완료, 하나 풀면 해제
sj.record_review_state(A["norm"], jn, {0, 1}, sj.SET_CONFIG_PATH, total=2)
d1 = req("GET", "/api/state")[3]["review"]["20260905101500"]
sj.record_review_state(A["norm"], jn, {0}, sj.SET_CONFIG_PATH, total=2)
d2 = req("GET", "/api/state")[3]["review"]["20260905101500"]
check("프로그램 오답노트 창의 전체 '이해했음' 체크 = 웹 ③ 완료, 하나 풀면 미완료", d1["done"] is True and d1["when"] and d2["done"] is False, str((d1, d2)))
st, _h, _t, j = req("POST", "/api/review", {"record_id": "99999999999999", "done": True}, token=TOK)
check("없는 기록 → 200 ok:false", st == 200 and j == {"ok": False, "error": "기록을 찾을 수 없습니다"})
st, _h, _t, j = req("POST", "/api/review", {"record_id": "20260904210000", "done": True}, token=TOK)
check("채점 JSON 없는 기록 → 200 ok:false", st == 200 and j["ok"] is False and "JSON" in j["error"])
st, _h, _t, j = req("POST", "/api/review", {"done": True}, token=TOK)
check("record_id 누락 → 400", st == 400)

# ---------------------------------------------------------------------------
# 6-2. /api/exam_date (v3.0.0) — 웹에서도 시험일을 바꿀 수 있어야 한다
# ---------------------------------------------------------------------------
g_ex = S3["generated"] if "S3" in dir() else None
st, _h, _t, j = req("POST", "/api/exam_date",
                    {"date": "2026-09-12", "date2": "2026-09-13"}, token=TOK)
check("exam_date 저장 → ok:true · 최신 state 에 반영 · 설정 파일에도 저장",
      st == 200 and j["ok"] is True
      and j["exam"]["dates"] == ["2026-09-12", "2026-09-13"]
      and j["settings"]["exam_date"] == "2026-09-12"
      and j["settings"]["exam_date2"] == "2026-09-13"
      and rjson(sj.SET_CONFIG_PATH)["_설정"]["시험일"] == "2026-09-12",
      json.dumps(j["exam"], ensure_ascii=False))
check("exam_date 저장 후 일정이 새 시험일 기준으로 다시 계산된다 (D-7 = 9/5)",
      j["plan"]["seg_start"] == "2026-09-05" and j["plan"]["end"] == "2026-09-11"
      and j["plan"]["kind"] == "week", json.dumps(j["plan"]["seg_start"]))
check("on_change('exam_date') 호출", CHANGES[-1][0] == "exam_date", str(CHANGES[-1:]))
st, _h, _t, j = req("POST", "/api/exam_date", {"date": None, "date2": None}, token=TOK)
check("date null → 시험일 지움(미설정) · plan.kind unset",
      st == 200 and j["ok"] is True and j["exam"] == {"dates": [], "active": [],
                                                      "past": [], "next": None,
                                                      "set": False,
                                                      "dday": None, "routine_days": 7,
                                                      "start": None}
      and j["plan"]["kind"] == "unset", json.dumps(j["exam"], ensure_ascii=False))
for body, why in (({"date": "2026-13-45"}, "달·일 범위 밖"),
                  ({"date": "내일"}, "형식 아님"),
                  ({"date": "2026-09-12", "date2": "2026-09-11"}, "2차가 앞섬"),
                  ({"date": "2026-09-12", "date2": "2026-09-12"}, "2차가 같음"),
                  ({"date": "2026-09-12", "date2": "x"}, "2차 형식 오류")):
    st, _h, _t, j = req("POST", "/api/exam_date", body, token=TOK)
    check(f"exam_date 400: {why}", st == 400 and j["ok"] is False, str(j))
# v3.1.0: 목록 형식 (여러 개 추가·개별 제거·0개)
st, _h, _t, j = req("POST", "/api/exam_date",
                    {"dates": ["2026-09-20", "2026-09-12", "2026-09-12"]},
                    token=TOK)
check("v3.1.0 exam_date {dates:[...]} — 여러 개 저장 · 정렬·중복 제거 · 목록 설정에 기록",
      st == 200 and j["ok"] is True
      and j["exam"]["dates"] == ["2026-09-12", "2026-09-20"]
      and j["settings"]["exam_dates"] == ["2026-09-12", "2026-09-20"]
      and rjson(sj.SET_CONFIG_PATH)["_설정"]["시험일목록"]
      == ["2026-09-12", "2026-09-20"],
      json.dumps(j["exam"], ensure_ascii=False))
check("v3.1.0 지난 시험일은 목록에 남고 기준(active/next)에서만 빠진다",
      j["exam"]["active"] == ["2026-09-12", "2026-09-20"]
      and j["exam"]["next"] == "2026-09-12" and j["exam"]["past"] == [],
      json.dumps(j["exam"], ensure_ascii=False))
st, _h, _t, j = req("POST", "/api/exam_date",
                    {"dates": ["2026-09-01", "2026-09-20"]}, token=TOK)
check("v3.1.0 지난 시험일(9/1)이 섞여도 일정은 다음 시험일(9/20) 기준",
      st == 200 and j["exam"]["past"] == ["2026-09-01"]
      and j["exam"]["next"] == "2026-09-20" and j["exam"]["dday"] == 14
      and j["plan"]["end"] == "2026-09-19", json.dumps(j["exam"], ensure_ascii=False))
st, _h, _t, j = req("POST", "/api/exam_date", {"dates": []}, token=TOK)
check("v3.1.0 빈 목록 → 시험일 0개(자유 연습) · 프로그램은 정상",
      st == 200 and j["ok"] is True and j["exam"]["dates"] == []
      and j["plan"]["kind"] == "unset")
st, _h, _t, j = req("POST", "/api/exam_date", {"dates": ["내일"]}, token=TOK)
check("v3.1.0 dates 형식 오류는 400", st == 400 and j["ok"] is False, str(j))
st, _h, _t, j = req("POST", "/api/exam_date", {"dates": "2026-09-12"}, token=TOK)
check("v3.1.0 dates 가 배열이 아니면 400", st == 400 and j["ok"] is False, str(j))
st, _h, _t, j = req("POST", "/api/exam_date", {"date": EXAM_DATE.isoformat()}, token=TOK)
check("date2 생략 가능(2차 없음) · 원래 시험일로 되돌림",
      st == 200 and j["ok"] and j["exam"]["dates"] == [EXAM_DATE.isoformat()]
      and j["settings"]["exam_date2"] is None)
st, _h, _t, j = req("POST", "/api/exam_date", {"date": "2026-09-12"})
check("토큰 없으면 401", st == 401 and j["ok"] is False)

# ---------------------------------------------------------------------------
# 7. /api/action
# ---------------------------------------------------------------------------
ACTIONS.clear()
st, _h, _t, j = req("POST", "/api/action", {"action": "show", "set": None, "record_id": None}, token=TOK)
check("show → ok:true + state · action_fn(show) 호출", st == 200 and j["ok"] and ACTIONS == [("show", None, None, None, None)], str(ACTIONS))
st, _h, _t, j = req("POST", "/api/action", {"action": "open_pdf", "set": C["norm"], "record_id": None}, token=TOK)
check("open_pdf: PDF 없는 세트 → 200 ok:false '문제지 PDF 없음'", st == 200 and j == {"ok": False, "error": "문제지 PDF 없음"})
st, _h, _t, j = req("POST", "/api/action", {"action": "open_pdf", "set": A["norm"], "record_id": None}, token=TOK)
check("open_pdf: PDF 있는 세트 → ok:true, action_fn 에 세트 전달", j["ok"] and ACTIONS[-1][0] == "open_pdf" and ACTIONS[-1][1] == A["norm"])
st, _h, _t, j = req("POST", "/api/action", {"action": "open_review", "set": None, "record_id": "20260905101500"}, token=TOK)
check("open_review(record_id) → 기록의 세트·채점 JSON 이 action_fn 에 전달", j["ok"] and ACTIONS[-1] == ("open_review", A["norm"], HTML_A[:-5] + ".json", None, "20260905101500"), str(ACTIONS[-1]))
st, _h, _t, j = req("POST", "/api/action", {"action": "start_retry", "set": B["norm"], "record_id": None}, token=TOK)
check("start_retry(set) → ok:true", j["ok"] and ACTIONS[-1][0] == "start_retry" and ACTIONS[-1][1] == B["norm"])
st, _h, _t, j = req("POST", "/api/action", {"action": "open_report", "set": None, "record_id": "20260904210000"}, token=TOK)
check("open_report: 리포트 없는 기록 → ok:false", st == 200 and j["ok"] is False)
st, _h, _t, j = req("POST", "/api/action", {"action": "start_exam", "set": "없는norm", "record_id": None}, token=TOK)
check("알 수 없는 세트 → ok:false", st == 200 and j == {"ok": False, "error": "알 수 없는 세트"})
# v2.7.0: 오늘 이미 응시한 세트(B형 72)도 잠기지 않는다 — 점수 갱신용 재응시
st, _h, _t, j = req("POST", "/api/action", {"action": "start_exam", "set": B["norm"], "record_id": None}, token=TOK)
check("오늘 이미 응시한 세트도 start_exam ok:true (응시 제한 없음 — 점수 갱신)",
      st == 200 and j["ok"] is True and ACTIONS[-1][:2] == ("start_exam", B["norm"]), str(ACTIONS[-1]))
EXAM["running"] = False
sj.routine_touch()
st, _h, _t, j = req("POST", "/api/action", {"action": "start_exam", "set": A["norm"], "record_id": None}, token=TOK)
check("같은 날 다른 세트도 이어서 start_exam ok:true (하루 1회 잠금 없음)",
      st == 200 and j["ok"] is True and ACTIONS[-1][:2] == ("start_exam", A["norm"]), str(ACTIONS[-1]))
EXAM["running"] = False
sj.routine_touch()
st, _h, _t, j = req("POST", "/api/action", {"action": "폭파", "set": None, "record_id": None}, token=TOK)
check("알 수 없는 action → 400", st == 400 and j["ok"] is False)
gx = req("GET", "/api/state")[3]["generated"]
st, _h, _t, j = req("POST", "/api/action", {"action": "start_exam", "set": None, "record_id": None}, token=TOK)
check("start_exam(set null=자동 선택) → ok:true · exam_running true 반영 · generated 변경", j["ok"] and j["exam_running"] is True
      and ACTIONS[-1] == ("start_exam", None, None, None, None) and j["generated"] > gx)
st, _h, _t, j = req("POST", "/api/action", {"action": "start_exam", "set": A["norm"], "record_id": None}, token=TOK)
check("시험 진행 중 start_exam → 200 {ok:false, error:'시험 진행 중'}", st == 200 and j == {"ok": False, "error": "시험 진행 중"})
st, _h, _t, j = req("POST", "/api/action", {"action": "start_retry", "set": B["norm"], "record_id": None}, token=TOK)
check("시험 진행 중 start_retry → 오류", j == {"ok": False, "error": "시험 진행 중"})
st, _h, _t, j = req("POST", "/api/action", {"action": "show", "set": None, "record_id": None}, token=TOK)
check("시험 진행 중에도 show 는 ok", j["ok"] is True)
EXAM["running"] = False
sj.routine_touch()
st, _h, _t, S3 = req("GET", "/api/state")
check("시험 종료(exam_running 토글)도 generated 갱신", S3["exam_running"] is False and S3["generated"] > j["generated"])
# action_fn 이 (False, 사유) → ok:false
srv.action_fn = lambda *a: (False, "세트를 찾을 수 없습니다")
st, _h, _t, j = req("POST", "/api/action", {"action": "open_review", "set": A["norm"], "record_id": None}, token=TOK)
check("프로그램 동작이 실패를 돌려주면 ok:false + 사유", j == {"ok": False, "error": "세트를 찾을 수 없습니다"})
srv.action_fn = action_fn

# ---------------------------------------------------------------------------
# 8. /api/report/<id>
# ---------------------------------------------------------------------------
st, hd, text, _ = req("GET", "/api/report/20260905101500")
check("report: 기록의 채점결과 HTML 그대로(text/html)", st == 200 and hd.get("Content-Type", "").startswith("text/html") and "채점결과" in text and "58점" in text)
st, _h, _t, j = req("GET", "/api/report/20260904210000")
check("report: 리포트 없는 기록 404", st == 404)
st, _h, _t, j = req("GET", "/api/report/99999999999999")
check("report: 없는 id 404", st == 404)
secret = os.path.join(TMP, "비밀.html")
open(secret, "w").write("비밀")
for bad in ("../비밀.html", "..%2F비밀.html", "%2e%2e/비밀.html", "/etc/passwd", os.path.basename(secret)):
    st, _h, text, _ = req("GET", "/api/report/" + bad)
    check(f"report 경로 탈출 차단({bad}) → 404", st == 404 and "비밀" not in text, f"{st}")

# ---------------------------------------------------------------------------
# 9. 동시 요청 · 포트 충돌 · generated
# ---------------------------------------------------------------------------
results = []
lock = threading.Lock()


def worker(i):
    try:
        if i % 3 == 0:
            r = req("POST", "/api/check", {"key": f"c{i}", "value": True}, token=TOK)
        elif i % 3 == 1:
            r = req("GET", "/api/state")
        else:
            r = req("GET", "/")
        with lock:
            results.append(r[0])
    except Exception as e:
        with lock:
            results.append(str(e))


ths = [threading.Thread(target=worker, args=(i,)) for i in range(30)]
[t.start() for t in ths]
[t.join(20) for t in ths]
cfg = rjson(sj.SET_CONFIG_PATH)
check("동시 요청 30개 전부 200", len(results) == 30 and all(r == 200 for r in results), str(results))
check("동시 check 10건 모두 저장(직렬화)", all(cfg["_웹체크"].get(f"c{i}") is True for i in range(0, 30, 3)), str(sorted(cfg["_웹체크"])))
srv2 = sj.RoutineServer(sets_fn=lambda: SETS, port=PORT, html_path=PAGE, today=TODAY)
srv2.start()
check("포트 충돌: 우선 포트가 사용 중이면 다른 빈 포트", srv2.port != PORT and srv2.port > 0, f"{PORT} → {srv2.port}")
check("두 번째 서버도 응답 · 토큰은 서로 다름", http.client.HTTPConnection("127.0.0.1", srv2.port, timeout=5).request("GET", "/api/state") is None and srv2.token != TOK)
srv2.stop()
check("stop 뒤 running false", srv2.running is False)
gens = set()
for _ in range(20):
    gens.add(sj.routine_touch())
check("routine_touch 20회 연속 호출 → 전부 다른 값(단조 증가)", len(gens) == 20 and sorted(gens) == sorted(gens))
# 파일이 GUI 밖에서 바뀌어도(grade.py·손 편집) generated 갱신
g_before = req("GET", "/api/state")[3]["generated"]
time.sleep(0.02)
recs = rjson(sj.RECORDS_PATH)
recs.append({"일시": "2026-09-06 12:00", "세트명": A["name"], "점수": 61, "mode": "시험"})
wjson(sj.RECORDS_PATH, recs)
os.utime(sj.RECORDS_PATH, None)
S4 = req("GET", "/api/state")[3]
check("기록.json 을 밖에서 고쳐도 다음 조회에서 generated 갱신 + 기록 반영", S4["generated"] > g_before and any(r["id"] == "20260906120000" for r in S4["records"]))
log = open(sj.STARTUP_LOG_PATH, encoding="utf-8").read()
check("시작 로그: 서버 시작·요청 요약(경로·상태)만 기록, /api/state 폴링은 기록 안 함",
      "루틴 서버 시작: http://127.0.0.1:" in log and "POST /api/check → 200" in log and "POST /api/check → 401" in log
      and "GET /api/state → 200" not in log and "GET /api/state → 403" in log and srv.stats["state"] > 10, str(srv.stats))

# ---------------------------------------------------------------------------
# 10. 순수 함수·업데이트 검증
# ---------------------------------------------------------------------------
ids2 = [i for i, _r in sj.records_with_ids([{"일시": "2026-09-06 10:12"}, {"일시": "2026-09-06 10:12"}, {"일시": "2026-09-06 10:12:36"}, {"일시": "?"}, {"일시": "2026-09-06T10:12"}])]
check("records_with_ids: 초 없음 '00', 중복 순번, 초 있음 그대로, 깨진 일시는 순번 id, T 구분자 허용",
      ids2 == ["20260906101200", "20260906101201", "20260906101236", "00000000000003", "20260906101202"], str(ids2))
EXD = [EXAM_DATE]
S_exam = sj.build_state(SETS, records=RECORDS, cfg={}, today=EXAM_DATE, exam=EXD)
kx = {d["date"]: d["kind"] for d in S_exam["plan"]["days"]}
check("시험 당일(9/10) 상태: plan.kind exam · seg null · start/end null · preview null · 9/9 past · 9/10 exam",
      S_exam["plan"]["kind"] == "exam"
      and S_exam["plan"]["seg"] is None and S_exam["plan"]["start"] is None and S_exam["plan"]["preview"] is None
      and kx["2026-09-09"] == "past" and kx["2026-09-10"] == "exam"
      and all(d["slots"] == [] for d in S_exam["plan"]["days"]), str(kx))
S_2 = sj.build_state(SETS, records=RECORDS, cfg={}, today=WEEK_END, exam=EXD)
k2 = {d["date"]: d["kind"] for d in S_2["plan"]["days"]}
d9 = [d for d in S_2["plan"]["days"] if d["date"] == "2026-09-09"][0]
check("마지막 학습일(9/9 = D-1) 상태: seg 1 · start=end=9/9 · mock(C 첫응시 1슬롯·deadline·실수 노트 제목·승격 없음) · 9/8 past · 9/10 exam · preview null",
      S_2["plan"]["seg"] == 1 and S_2["plan"]["start"] == "2026-09-09" and S_2["plan"]["end"] == "2026-09-09" and k2["2026-09-09"] == "mock"
      and d9["deadline"] is True and d9["promoted"] is False and d9["title"] == "추천 1세트 + 실수 노트" and len(d9["slots"]) == 1
      and d9["slots"][0]["set"]["norm"] == C["norm"] and d9["slots"][0]["kind"] == "first"
      and S_2["plan"]["attempts_left"] == 1 and k2["2026-09-08"] == "past" and k2["2026-09-10"] == "exam"
      and S_2["plan"]["preview"] is None, str((k2, d9["title"])))
S_lock = sj.build_state(SETS, records=RECORDS, cfg={}, today=date(2026, 9, 7), exam=EXD)
check("v2.7.0: 응시 제한 폐지 — today_locked 는 오늘 응시 여부와 무관하게 항상 false",
      S_lock["today_locked"] is False and S["today_locked"] is False
      and S_lock["attempts_left"] == 3
      and S_lock["streak"] == {"current": 2, "best": 2, "today_done": False},
      str((S_lock["today_locked"], S_lock["attempts_left"])))
S_after = sj.build_state(SETS, records=RECORDS, cfg={}, today=date(2026, 9, 20), exam=EXD)
check("루틴 종료 후: kind after · 시험일 exam · 나머지 past · '시험일을 다시 설정' 안내",
      S_after["plan"]["kind"] == "after"
      and all(d["kind"] in ("past", "exam") for d in S_after["plan"]["days"])
      and "시험일을 다시 설정" in S_after["plan"]["reason"], S_after["plan"]["reason"])
S_unset = sj.build_state(SETS, records=RECORDS, cfg={}, today=TODAY, exam=[])
check("시험일 미설정: kind unset · days 빈 목록 · exam.set false · 시험 날짜 입력 안내",
      S_unset["plan"]["kind"] == "unset" and S_unset["plan"]["days"] == []
      and S_unset["exam"] == {"dates": [], "active": [], "past": [],
                              "next": None, "set": False, "dday": None,
                              "routine_days": 7, "start": None}
      and "시험일 미설정" in S_unset["plan"]["reason"],
      json.dumps(S_unset["exam"], ensure_ascii=False))
S_far = sj.build_state(SETS, records=RECORDS, cfg={}, today=date(2026, 8, 20), exam=EXD)
check("시험일까지 7일보다 많이 남음: kind free · D-7(9/3)부터의 일주일을 미리 보여 준다",
      S_far["plan"]["kind"] == "free" and S_far["plan"]["seg_start"] == "2026-09-03"
      and S_far["plan"]["end"] == "2026-09-09"
      and "자유 연습" in S_far["plan"]["reason"], S_far["plan"]["reason"])
S_cut = sj.build_state(SETS, records=[], cfg={}, today=date(2026, 9, 8), exam=EXD)
check("시험일까지 7일 미만(D-2에 처음 실행): 뒤에서부터 잘려 9/8·9/9 2일만 배정",
      S_cut["plan"]["start"] == "2026-09-08" and S_cut["plan"]["attempts_left"] == 2
      and [d["date"] for d in S_cut["plan"]["days"] if d["kind"] == "mock"]
      == ["2026-09-08", "2026-09-09"], str([(d["date"], d["kind"]) for d in S_cut["plan"]["days"]]))
S_nosets = sj.build_state([], records=[], cfg={}, today=TODAY, exam=EXD)
check("보유 세트 0개: 배정 없음 · no_sets true · '문제·정답 파일을 넣어 주세요' 안내",
      S_nosets["plan"]["no_sets"] is True and S_nosets["plan"]["total_sets"] == 0
      and all(d["slots"] == [] for d in S_nosets["plan"]["days"])
      and S_nosets["plan"]["reason"] == sj.NO_SETS_REASON, S_nosets["plan"]["reason"])
check("빈 세트·빈 기록도 죽지 않음",
      sj.build_state([], records=[], cfg={}, today=TODAY, exam=EXD)["plan"]["days"][3]["kind"] in ("mock", "rest"))
check("render_routine_page: doctype 없는 파일 감쌈 / 있는 파일 그대로", sj.render_routine_page("<title>x</title>").startswith(sj.ROUTINE_PAGE_HEAD)
      and sj.render_routine_page("  <!DOCTYPE html><html></html>").strip().startswith("<!DOCTYPE"))
html_bytes = "<title>루틴</title>".encode("utf-8")
info = {"version": "2.4.0", "sha256": {"루틴.html": sj.sha256_hex(html_bytes)}}
check("_verify_download: set_files 의 루틴.html 은 sha256 + UTF-8 로 통과(텍스트 반환)", sj._verify_download("루틴.html", html_bytes, "data", info) == "<title>루틴</title>")
try:
    sj._verify_download("루틴.html", html_bytes, "data", {"version": "2.4.0"})
    bad = False
except RuntimeError:
    bad = True
check("_verify_download: sha256 없는 루틴.html 은 거부", bad)
try:
    sj._verify_download("루틴.html", b"\xff\xfe\x00", "data", {"sha256": {"루틴.html": sj.sha256_hex(b"\xff\xfe\x00")}})
    bad = False
except RuntimeError:
    bad = True
check("_verify_download: UTF-8 아닌 루틴.html 은 거부", bad)
tgt = sj._data_target_path(sj.ROUTINE_DATA_KEY, os.path.join(TMP, "설치", "시험장"))
check("설치 위치: '시험장/루틴.html' → <루트>/시험장/루틴.html (= 시험장 폴더)", tgt == os.path.join(TMP, "설치", "시험장", "루틴.html"), tgt)
check("routine_html_path 는 시험장 폴더의 루틴.html 우선", sj.routine_html_path(os.path.join(TMP, "설치", "시험장")) == tgt)
check("normalize_set_config: _웹체크 는 dict 만 살림", sj.normalize_set_config({"_웹체크": "x"}).get("_웹체크") is None
      and sj.normalize_set_config({"_웹체크": {"a": True}})["_웹체크"] == {"a": True})
pub = os.path.join(os.path.dirname(os.path.dirname(BASE)), "..", "cocomate-study-tools", "version.json")
pub = os.path.normpath(os.path.join(os.path.dirname(os.path.dirname(BASE)), "cocomate-study-tools", "version.json"))
if os.path.isfile(pub):
    vj = rjson(pub)
    if sj._version_tuple(vj.get("version", "0")) >= (2, 4, 0):
        check("공개 version.json(2.4.0+): set_files 에 '시험장/루틴.html' → '루틴.html' + sha256", vj["set_files"].get("시험장/루틴.html") == "루틴.html"
              and "루틴.html" in vj["sha256"] and "루틴.html" not in json.dumps(vj["data_files"], ensure_ascii=False))

srv.stop()
shutil.rmtree(TMP, ignore_errors=True)
print()
print(f"연동 서버 테스트 {N}건 전부 통과")
