# -*- coding: utf-8 -*-
"""시험장.py v2.6.0 자동 재채점(이의제기 반영) 헤드리스 테스트.

  a. 채점기 버전만 오래된 기록 → 자동으로 다시 채점되고 점수가 정정된다
  b. 정답 파일 해시가 바뀐 기록 → 다시 채점된다
  c. 채점 근거가 그대로인 기록 → 대상이 아니다(건너뜀, 파일도 안 건드림)
  d. 풀이 파일이 없어 다시 채점할 수 없는 기록 → 정정 데이터로 대체된다
  e. 원점수를 정정이력에 보관하고, 같은 정정을 두 번 반영하지 않는다
  f. 정정된 점수로 레벨·연속 응시·합격 세트·일정 배정·웹 상태가 다시 계산된다
  g. 손상된 기록.json·정정 파일·이상한 값에도 죽지 않는다

합성 문제/정답/풀이 xlsx + 테스트용 가짜 채점기(stub grade.py)를 써서
grade.py 의 채점 로직과 무관하게 '재채점 기계'만 검증합니다. 진짜 grade.py 로는
버전 표식 조회(grader_version)만 확인합니다.
"""
import importlib.util
import json
import os
import shutil
import sys
import tempfile
import zipfile
from datetime import date, datetime

BASE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("sj", os.path.join(BASE, "시험장.py"))
sj = importlib.util.module_from_spec(spec)
spec.loader.exec_module(sj)

N = 0
TMP = tempfile.mkdtemp(prefix="coco_regrade_")
sj.STARTUP_LOG_PATH = os.path.join(TMP, "시작로그.txt")
sj.ERROR_LOG_PATH = os.path.join(TMP, "오류.log")
sj.SET_CONFIG_PATH = os.path.join(TMP, "세트설정.json")


def check(desc, cond, extra=""):
    global N
    N += 1
    print(f"  [{N:02d}] {desc}: {'OK' if cond else 'FAIL'}" + (f" ({extra})" if extra else ""))
    if not cond:
        raise AssertionError(f"{desc} {extra}")


# ---------------------------------------------------------------------------
# 합성 파일 (문제·정답·풀이 xlsx + 가짜 채점기)
# ---------------------------------------------------------------------------

_MIN_CT = ('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.'
           'openxmlformats.org/package/2006/content-types"><Default Extension'
           '="rels" ContentType="application/vnd.openxmlformats-package.'
           'relationships+xml"/><Override PartName="/xl/workbook.xml" '
           'ContentType="application/vnd.openxmlformats-officedocument.'
           'spreadsheetml.sheet.main+xml"/></Types>')


def make_xlsx(path, marker):
    """합성 xlsx 1개 (openpyxl 이 있으면 진짜 워크북, 없으면 최소 ZIP).

    marker 를 바꾸면 파일 내용(=sha256)이 달라집니다."""
    try:
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "계산작업"
        ws["A1"] = "합성 세트"
        ws["A2"] = str(marker)
        wb.save(path)
        return path
    except Exception:
        pass
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", _MIN_CT)
        z.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships xmlns='
                   '"http://schemas.openxmlformats.org/package/2006/relationships"/>')
        z.writestr("xl/workbook.xml", f'<?xml version="1.0"?><workbook><!--{marker}--></workbook>')
    return path


STUB_GRADER = '''#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""테스트용 가짜 채점기 — 정답 파일 sha256 으로 점수를 정합니다."""
__version__ = "{ver}"
import argparse, hashlib, json, os, sys


def main():
    ap = argparse.ArgumentParser()
    for name in ("--problem", "--answer", "--student", "--key", "--html",
                 "--history", "--sheets"):
        ap.add_argument(name)
    ap.add_argument("--json", dest="json_out")
    a = ap.parse_args()
    for p in (a.problem, a.answer, a.student):
        if not p or not os.path.isfile(p):
            print("파일 없음: " + str(p), file=sys.stderr)
            return 2
    with open(a.answer, "rb") as f:
        sha = hashlib.sha256(f.read()).hexdigest()
    table = {{}}
    tpath = os.path.join(os.path.dirname(os.path.abspath(__file__)), "점수표.json")
    try:
        with open(tpath, encoding="utf-8") as f:
            table = json.load(f)
    except Exception:
        pass
    if table.get("_실패"):
        print("의도적 실패", file=sys.stderr)
        return 3
    total = table.get(sha, table.get("_기본", 70))
    data = {{"total": total, "pass_line": 70, "passed": total >= 70,
            "mode": "full", "graded_sheets": ["계산작업"], "max_total": 100,
            "generated": "2026-09-13T10:00:00",
            "files": {{"problem": os.path.abspath(a.problem),
                      "answer": os.path.abspath(a.answer),
                      "student": os.path.abspath(a.student)}},
            "sheets": [{{"name": "계산작업", "alloc": 40,
                        "earned": 40 if total >= 90 else 24,
                        "missing": [], "details": [], "notes": []}}],
            "notes": [],
            "wrong_items": ([] if total >= 90 else
                            [{{"sheet": "계산작업", "label": "3번 함수",
                              "lost": 8, "kind": "formula",
                              "category": "계산작업", "cells": ["E5"]}}])}}
    if a.json_out:
        with open(a.json_out, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    if a.html:
        with open(a.html, "w", encoding="utf-8") as f:
            f.write("<h1>합성 채점 리포트 " + str(total) + "점</h1>")
    print("총점 100 " + str(total))
    return 0


if __name__ == "__main__":
    sys.exit(main())
'''


def write_grader(folder, ver):
    os.makedirs(folder, exist_ok=True)
    p = os.path.join(folder, "grade.py")
    with open(p, "w", encoding="utf-8") as f:
        f.write(STUB_GRADER.format(ver=ver))
    return p


def score_table(folder, table):
    with open(os.path.join(folder, "점수표.json"), "w", encoding="utf-8") as f:
        json.dump(table, f, ensure_ascii=False)


class Case:
    """세트 폴더 하나 + 기록.json + 가짜 채점기로 이루어진 시나리오 작업장."""

    def __init__(self, name, grader_ver="2.0.3"):
        self.root = os.path.join(TMP, name)
        self.set_dir = os.path.join(self.root, "세트")
        self.grade_dir = os.path.join(self.root, "채점기")
        os.makedirs(self.set_dir, exist_ok=True)
        self.problem = make_xlsx(os.path.join(self.set_dir, "합성세트_문제.xlsx"), "P")
        self.answer = make_xlsx(os.path.join(self.set_dir, "합성세트_정답.xlsx"), "A")
        self.grade_py = write_grader(self.grade_dir, grader_ver)
        score_table(self.grade_dir, {"_기본": 95})
        self.records_path = os.path.join(self.root, "기록.json")
        self.corrections = os.path.join(self.root, "점수정정.json")
        self.set = {"name": "합성 세트", "norm": sj.norm_set_key("합성 세트"),
                    "dir": self.set_dir, "problem": self.problem,
                    "answer": self.answer, "key": None, "pdf": None}

    def attempt(self, stamp="20260905_2205"):
        p = os.path.join(self.set_dir, f"풀이_합성 세트_{stamp}.xlsm")
        make_xlsx(p, "S" + stamp)
        return p

    def report(self, stamp="20260905_2205", total=86):
        d = os.path.join(self.set_dir, "채점결과")
        os.makedirs(d, exist_ok=True)
        html = os.path.join(d, f"채점결과_합성 세트_{stamp}.html")
        with open(html, "w", encoding="utf-8") as f:
            f.write(f"<h1>{total}점</h1>")
        with open(os.path.splitext(html)[0] + ".json", "w", encoding="utf-8") as f:
            json.dump({"total": total, "pass_line": 70, "passed": total >= 70,
                       "sheets": [{"name": "계산작업", "alloc": 40, "earned": 24}],
                       "wrong_items": []}, f, ensure_ascii=False)
        return html

    def save(self, records, raw=None):
        with open(self.records_path, "w", encoding="utf-8") as f:
            if raw is not None:
                f.write(raw)
            else:
                json.dump(records, f, ensure_ascii=False, indent=2)

    def load(self):
        return sj.load_records(self.records_path)

    def write_corrections(self, obj, raw=None):
        with open(self.corrections, "w", encoding="utf-8") as f:
            if raw is not None:
                f.write(raw)
            else:
                json.dump(obj, f, ensure_ascii=False)

    def coord(self, **kw):
        kw.setdefault("sets", [self.set])
        kw.setdefault("grade_py", self.grade_py)
        kw.setdefault("records_path", self.records_path)
        kw.setdefault("corrections_file", self.corrections)
        kw.setdefault("when", datetime(2026, 9, 13, 10, 0, 0))
        return sj.RegradeCoordinator(**kw)


def base_record(**kw):
    r = {"일시": "2026-09-05 22:05", "세트명": "합성 세트", "점수": 86,
         "소요시간": "38분 10초", "리포트": None, "mode": "시험"}
    r.update(kw)
    return r


print("1. 채점 근거 기록 (채점기 버전 · 파일 해시 · 풀이 경로)")
c = Case("evidence")
stu = c.attempt()
ev = sj.grading_evidence(problem=c.problem, answer=c.answer, student=stu,
                         key=None, grade_py=c.grade_py)
check("grading_evidence: 채점기버전·문제해시·정답해시·풀이경로",
      ev["채점기버전"] == "2.0.3" and len(ev["정답해시"]) == 64
      and ev["문제해시"] == sj.file_sha256(c.problem)
      and ev["풀이경로"] == os.path.abspath(stu) and "기대값해시" not in ev)
key_json = os.path.join(c.set_dir, "합성세트_기대값.json")
with open(key_json, "w", encoding="utf-8") as f:
    json.dump({"a": 1}, f)
ev2 = sj.grading_evidence(problem=c.problem, answer=c.answer, student=stu,
                          key=key_json, grade_py=c.grade_py)
check("기대값 파일이 있으면 기대값해시도 저장", ev2["기대값해시"] == sj.file_sha256(key_json))
check("grader_version: 캐시 후에도 같은 값 · 없는 경로는 None",
      sj.grader_version(c.grade_py) == "2.0.3"
      and sj.grader_version(os.path.join(TMP, "없음.py")) is None)
real = sj.find_grade_py()
check("진짜 grade.py 의 __version__ 도 읽힘 (--print-version 없이 상수 조회)",
      real is None or (sj.grader_version(real) or "").count(".") == 2,
      str(sj.grader_version(real)))
check("REGRADE 상수", sj.CORRECTIONS_DATA_KEY == "정정/점수정정.json"
      and sj.CORRECTION_HISTORY_FIELD == "정정이력"
      and sj.corrections_path().endswith(os.path.join("정정", "점수정정.json")))

print()
print("2. (a) 채점기 버전만 오래된 기록 → 자동 재채점")
c = Case("case_a")
stu = c.attempt()
html = c.report(total=86)
rec = base_record(리포트=html, 채점기버전="1.9.0", 풀이경로=stu,
                  문제해시=sj.file_sha256(c.problem),
                  정답해시=sj.file_sha256(c.answer))
c.save([rec])
plan = c.coord().plan()
check("plan: 사유에 새 채점기 버전이 들어감", len(plan) == 1 and "채점기 2.0.3" in plan[0][3], plan and plan[0][3])
summary = c.coord().run()
recs = c.load()
check("(a) 86 → 95 점으로 정정", recs[0]["점수"] == 95 and summary["regraded"] == 1
      and len(summary["changes"]) == 1, recs[0]["점수"])
check("(a) 원점수는 정정이력에 보관", recs[0]["정정이력"][0]["이전총점"] == 86
      and recs[0]["정정이력"][0]["새총점"] == 95
      and recs[0]["정정이력"][0]["출처"] == "재채점"
      and "채점기 2.0.3" in recs[0]["정정이력"][0]["사유"])
check("(a) 새 리포트를 가리키고 예전 리포트는 남아 있음",
      recs[0]["리포트"].endswith("_재채점.html") and os.path.isfile(recs[0]["리포트"])
      and os.path.isfile(html) and os.path.isfile(os.path.splitext(recs[0]["리포트"])[0] + ".json"))
check("(a) 채점기 버전이 현재 값으로 갱신됨", recs[0]["채점기버전"] == "2.0.3")
check("(a) 새 채점결과 JSON 의 시트·오답이 규약 형식으로 읽힘",
      (sj.load_result_summary(os.path.splitext(recs[0]["리포트"])[0] + ".json") or {})
      .get("sheets") == [{"name": "계산작업", "alloc": 40, "earned": 40}])
summary2 = c.coord().run()
check("(a) 두 번째 실행은 대상 없음 (채점 근거가 갱신돼 다시 채점하지 않음)",
      summary2["targets"] == 0 and summary2["changes"] == [], summary2)
check("(a) 정정이력이 늘어나지 않음", len(c.load()[0]["정정이력"]) == 1)

print()
print("3. (b) 정답 파일 해시 변경 → 자동 재채점")
c = Case("case_b")
stu = c.attempt()
rec = base_record(점수=86, 리포트=c.report(total=86), 채점기버전="2.0.3",
                  풀이경로=stu, 문제해시=sj.file_sha256(c.problem),
                  정답해시=sj.file_sha256(c.answer))
c.save([rec])
check("(b-사전) 바뀐 것이 없으면 대상 아님", c.coord().plan() == [])
make_xlsx(c.answer, "A2")                     # 이의제기 인정 → 정답 파일 정정
score_table(c.grade_dir, {sj.file_sha256(c.answer): 94, "_기본": 86})
plan = c.coord().plan()
check("(b) 정답 해시가 달라 대상이 되고 사유가 '정답 파일 정정'",
      len(plan) == 1 and plan[0][3] == "정답 파일 정정", plan and plan[0][3])
summary = c.coord().run()
recs = c.load()
check("(b) 86 → 94 정정 + 정답해시 갱신", recs[0]["점수"] == 94
      and recs[0]["정답해시"] == sj.file_sha256(c.answer)
      and recs[0]["정정이력"][0]["사유"] == "정답 파일 정정")
check("(b) 안내 문구", sj.correction_summary_text(summary["changes"])
      == "이의제기 반영: 합성 세트 86 → 94점",
      sj.correction_summary_text(summary["changes"]))
make_xlsx(c.problem, "P2")
score_table(c.grade_dir, {"_기본": 94})
plan = c.coord().plan()
check("(b) 문제 파일이 바뀌어도 대상", len(plan) == 1 and plan[0][3] == "문제 파일 정정")
summary = c.coord().run()
check("(b) 점수가 그대로면 조용히 넘어감(정정이력·안내 없음)",
      summary["regraded"] == 1 and summary["changes"] == []
      and len(c.load()[0]["정정이력"]) == 1 and c.load()[0]["점수"] == 94)
check("(b) 그래도 문제해시는 갱신돼 다음 실행에서 다시 채점하지 않음",
      c.coord().plan() == [] and c.load()[0]["문제해시"] == sj.file_sha256(c.problem))

print()
print("4. (c) 변화 없음 → 건너뜀 / 대상 선별 규칙")
c = Case("case_c")
stu = c.attempt()
good = base_record(리포트=c.report(), 채점기버전="2.0.3", 풀이경로=stu,
                   문제해시=sj.file_sha256(c.problem),
                   정답해시=sj.file_sha256(c.answer))
manual = base_record(일시="2026-09-06 09:00", 점수=72, mode="수동", 리포트=None)
partial = base_record(일시="2026-09-06 10:00", 점수=20, mode="부분연습",
                      채점기버전="1.0.0", 풀이경로=stu)
retry = base_record(일시="2026-09-06 11:00", 점수=12, mode="오답재풀이",
                    채점기버전="1.0.0", 풀이경로=stu)
aborted = base_record(일시="2026-09-06 12:00", 점수=None, 채점기버전="1.0.0",
                      풀이경로=stu)
c.save([good, manual, partial, retry, aborted])
before = json.load(open(c.records_path, encoding="utf-8"))
summary = c.coord().run()
check("(c) 근거가 그대로면 대상 아님", summary["targets"] == 0 and summary["regraded"] == 0)
check("(c) 수동·부분연습·오답재풀이·중단 응시는 대상이 아님",
      all(sj.regrade_reason(r, c.set, "9.9.9") == ""
          for r in (manual, partial, retry, aborted)))
check("(c) 파일을 건드리지 않음", json.load(open(c.records_path, encoding="utf-8")) == before)
check("(c) 세트를 못 찾는 기록은 건너뜀",
      sj.RegradeCoordinator(sets=[], grade_py=c.grade_py,
                            records_path=c.records_path).plan() == [])
old = base_record(일시="2026-09-07 09:00", 채점기버전=None, 풀이경로=stu)
check("(c) 버전 기록이 아예 없는 옛 기록은 대상 (사유에 '버전 기록 없음')",
      "버전 기록 없음" in sj.regrade_reason(old, c.set, "2.0.3"))
check("(c) 채점기를 못 찾으면(현재 버전 None) 버전만으로는 대상이 아님",
      sj.regrade_reason(old, c.set, None) == "")

print()
print("5. (d) 풀이 파일이 없는 기록 → 정정 데이터로 대체")
c = Case("case_d")
rec = base_record(리포트=c.report(total=86), 채점기버전="1.0.0",
                  풀이경로=os.path.join(c.set_dir, "없는풀이.xlsm"))
c.save([rec])
rid = sj.records_with_ids(c.load())[0][0]
check("(d) 기록 id 는 일시 14자리", rid == "20260905220500", rid)
plan = c.coord().plan()
check("(d) 대상이지만 풀이 파일이 없음", len(plan) == 1 and plan[0][4] is None)
c.write_corrections({"corrections": [
    {"report_id": rid, "set": "합성 세트", "new_total": 91,
     "reason": "이의제기 인정 — 자료 입력 오탐"}]})
summary = c.coord().run()
recs = c.load()
check("(d) 정정 데이터로 86 → 91 정정", recs[0]["점수"] == 91
      and summary["data_applied"] == 1 and summary["skipped"] == 1
      and summary["regraded"] == 0)
check("(d) 정정이력 출처·사유", recs[0]["정정이력"][0]["출처"] == "정정데이터"
      and recs[0]["정정이력"][0]["사유"] == "이의제기 인정 — 자료 입력 오탐"
      and recs[0]["정정이력"][0]["이전총점"] == 86)
check("(d) 예전 리포트를 지우지 않고 그대로 가리킴", os.path.isfile(recs[0]["리포트"]))
check("(d) 안내 문구", sj.correction_summary_text(summary["changes"])
      == "이의제기 반영: 합성 세트 86 → 91점")
check("(d) 세트명이 다른 정정 항목은 무시",
      sj.RegradeCoordinator(sets=[c.set], grade_py=c.grade_py,
                            records_path=c.records_path,
                            corrections_file=c.corrections)._set_for(recs[0]) is not None)

print()
print("6. (e) 정정 이력 보관 · 중복 정정 방지")
summary2 = c.coord().run()
recs = c.load()
check("(e) 같은 정정을 다시 반영하지 않음", summary2["data_applied"] == 0
      and summary2["changes"] == [] and len(recs[0]["정정이력"]) == 1
      and recs[0]["점수"] == 91)
c.write_corrections({"corrections": [
    {"report_id": rid, "set": "합성 세트", "new_total": 91, "reason": "재공지"},
    {"report_id": rid, "new_total": 96, "reason": "이의제기 2차 인정"}]})
summary3 = c.coord().run()
recs = c.load()
check("(e) 점수가 또 바뀌면 이력이 쌓임(91 → 96, 같은 값 91 은 중복이라 무시)",
      recs[0]["점수"] == 96 and len(recs[0]["정정이력"]) == 2
      and recs[0]["정정이력"][1]["이전총점"] == 91
      and recs[0]["정정이력"][1]["새총점"] == 96, summary3)
check("(e) 정정이력 항목에 식별 키가 남아 중복을 막음",
      recs[0]["정정이력"][0][sj.CORRECTION_KEY_FIELD] == f"data:{rid}:91"
      and recs[0]["정정이력"][1][sj.CORRECTION_KEY_FIELD] == f"data:{rid}:96")
check("(e) 재채점 정정도 apply_correction 단위로 중복을 막는다",
      sj.apply_correction(dict(recs[0]), 96, "재공지", key=f"data:{rid}:96") is None)
check("(e) 점수가 같으면 정정이력을 남기지 않음",
      sj.apply_correction({"점수": 90}, 90, "같은 점수") is None)

print()
print("6-2. (e2) delta 가산점 — 다시 채점해도 인정한 점수가 유지된다")
# 문제지 쪽 결함처럼 채점기가 알 수 없는 사유는 total 을 못박지 않고
# 채점 결과 '위에' 얹는다. 못박아 두면 다음 재채점이 덮어써 인정이 날아간다.
c = Case("case_e2")
stu = c.attempt()
rec = base_record(점수=82, 리포트=c.report(total=82), 채점기버전="2.0.5",
                  풀이경로=stu)
c.save([rec])
rid = sj.records_with_ids(c.load())[0][0]
score_table(c.grade_dir, {"_기본": 87})      # 새 채점기가 매기는 점수
c.write_corrections({"corrections": [
    {"report_id": rid, "set": "합성 세트", "delta": 5,
     "reason": "문제지 결함 — 기본작업-1 공백 표기"}]})
write_grader(c.grade_dir, "2.0.6")
summary = c.coord().run()
recs = c.load()
check("(e2) 재채점 87 + 가산점 5 = 92", recs[0]["점수"] == 92
      and summary["regraded"] == 1 and summary["data_applied"] == 0, summary)
check("(e2) 사유에 채점기 사유와 가산점이 함께 남는다",
      "+5점(문제지 결함 — 기본작업-1 공백 표기)"
      in recs[0]["정정이력"][-1]["사유"], recs[0]["정정이력"][-1]["사유"])
summary_b = c.coord().run()
recs = c.load()
check("(e2) 같은 버전으로 다시 돌려도 92 그대로 (중복 가산 없음)",
      recs[0]["점수"] == 92 and summary_b["changes"] == [], summary_b)
score_table(c.grade_dir, {"_기본": 90})      # 또 다음 채점기 버전
write_grader(c.grade_dir, "2.0.7")
summary_c = c.coord().run()
recs = c.load()
check("(e2) 채점기가 또 바뀌어도 가산점은 새 결과 위에 다시 얹힌다 (90+5=95)",
      recs[0]["점수"] == 95, summary_c)

# 풀이 파일이 없어 재채점을 못 하는 기록은 정정 데이터가 현재 점수에 더한다
c2 = Case("case_e3")
rec2 = base_record(점수=80, 리포트=c2.report(total=80), 채점기버전="1.0.0",
                   풀이경로=os.path.join(c2.set_dir, "없는풀이.xlsm"))
c2.save([rec2])
rid2 = sj.records_with_ids(c2.load())[0][0]
c2.write_corrections({"corrections": [
    {"report_id": rid2, "delta": 5, "reason": "문제지 결함"}]})
s2a = c2.coord().run()
recs2 = c2.load()
check("(e3) 재채점 못 하는 기록은 현재 점수 80 + 5 = 85",
      recs2[0]["점수"] == 85 and s2a["data_applied"] == 1, s2a)
s2b = c2.coord().run()
recs2 = c2.load()
check("(e3) 다시 돌려도 85 그대로 (가산점 중복 없음)",
      recs2[0]["점수"] == 85 and s2b["data_applied"] == 0, s2b)
check("(e3) delta 항목의 식별 키는 부호를 포함한다",
      recs2[0]["정정이력"][0][sj.CORRECTION_KEY_FIELD] == f"data:{rid2}:+5")
c2.write_corrections({"corrections": [{"report_id": "z", "reason": "빈 항목"}]})
check("(e3) new_total 도 delta 도 없는 항목은 무시",
      sj.load_corrections(c2.corrections) == [],
      str(sj.load_corrections(c2.corrections)))

print()
print("6-3. (e4) 리포트 시각의 초가 어긋나도 같은 분 기록을 찾는다")
c3 = Case("case_e4")
rec3 = base_record(일시="2026-09-15 13:02:16", 점수=87,
                   리포트=c3.report(total=87), 채점기버전="2.0.6",
                   풀이경로=os.path.join(c3.set_dir, "없는풀이.xlsm"))
c3.save([rec3])
rid3 = sj.records_with_ids(c3.load())[0][0]
check("(e4) 기록 id 는 초까지", rid3 == "20260915130216", rid3)
c3.write_corrections({"corrections": [
    {"report_id": "20260915130200", "delta": 5, "reason": "문제지 결함"}]})
s3 = c3.coord().run()
check("(e4) 초가 다른 id 도 같은 분 기록 하나면 찾아 반영 (87+5=92)",
      c3.load()[0]["점수"] == 92 and s3["data_applied"] == 1, s3)
check("(e4) 같은 분 기록이 둘이면 찾지 않는다",
      sj.resolve_correction_target({"20260915130216": {}, "20260915130201": {}},
                                   "20260915130200")[1] is None)
check("(e4) 아예 없는 id 는 그대로 못 찾음",
      sj.resolve_correction_target({"20260915130216": {}},
                                   "20260914130200")[1] is None)

print()
print("7. (f) 파급 재계산 — 레벨 · 연속 응시 · 합격 세트 · 일정 · 웹 상태")
c = Case("case_f")
sets = []
for i, nm in enumerate(("합성 세트", "베타 세트", "감마 세트")):
    d = os.path.join(c.root, f"세트{i}")
    os.makedirs(d, exist_ok=True)
    s = {"name": nm, "norm": sj.norm_set_key(nm), "dir": d,
         "problem": make_xlsx(os.path.join(d, "문제.xlsx"), "P" + nm),
         "answer": make_xlsx(os.path.join(d, "정답.xlsx"), "A" + nm),
         "key": None, "pdf": None}
    sets.append(s)
stu0 = os.path.join(sets[0]["dir"], "풀이_합성 세트_20260911_1000.xlsm")
make_xlsx(stu0, "S0")
r1 = {"일시": "2026-09-11 10:00", "세트명": "합성 세트", "점수": 66,
      "소요시간": "39분 00초", "리포트": None, "mode": "시험",
      "채점기버전": "1.0.0", "풀이경로": stu0}
r2 = {"일시": "2026-09-12 10:00", "세트명": "베타 세트", "점수": 74,
      "소요시간": "39분 00초", "리포트": None, "mode": "시험"}
c.save([r1, r2])
today = date(2026, 9, 12)
EXAM_D = [date(2026, 9, 15)]       # 오늘 = D-3 (시험일은 사용자가 입력하는 값)
lv0 = sj.level_info(c.load())
stk0 = sj.streak_info(c.load(), today, EXAM_D)
plan0 = sj.build_adaptive_plan(today, c.load(), sets, EXAM_D)
check("(f-사전) 최고 74 → 현재 목표 80 · 합격 1세트 · 66점 세트는 재응시 대기줄 1순위",
      lv0["best_total"] == 74 and lv0["current_goal"] == 80
      and plan0["passed_sets"] == 1 and stk0["current"] == 2
      and (plan0["retry_queue"][0]["name"], plan0["retry_queue"][0]["kind"])
      == ("합성 세트", "retry"),
      [(q["name"], q["kind"], q["best"]) for q in plan0["retry_queue"]])
score_table(c.grade_dir, {"_기본": 92})
summary = sj.RegradeCoordinator(sets=sets, grade_py=c.grade_py,
                                records_path=c.records_path,
                                corrections_file=c.corrections,
                                when=datetime(2026, 9, 13, 10, 0)).run()
recs = c.load()
check("(f) 66 → 92 정정", recs[0]["점수"] == 92 and len(summary["changes"]) == 1)
lv = sj.level_info(recs)
plan1 = sj.build_adaptive_plan(today, recs, sets, EXAM_D)
check("(f) 레벨 재계산: 최고 92 · 70·80·90 달성 · 현재 목표 100",
      lv["best_total"] == 92 and lv["achieved"] == [70, 80, 90]
      and lv["current_goal"] == 100, lv)
check("(f) 합격 세트 수 재계산 1 → 2", plan1["passed_sets"] == 2)
check("(f) 대기줄에서 '합격선 미달 재응시'가 사라지고 '목표 도전'만 남음",
      [(q["name"], q["kind"]) for q in plan1["retry_queue"]]
      == [("베타 세트", "goal"), ("합성 세트", "goal")],
      [(q["name"], q["kind"], q["best"]) for q in plan1["retry_queue"]])
check("(f) 연속 응시는 그대로(날짜는 안 바뀜)",
      sj.streak_info(recs, today, EXAM_D)["current"] == 2)
st = sj.build_state(sets, records=recs, cfg={}, today=today, exam=EXAM_D)
check("(f) build_state 에 corrections 추가 (웹 배지용)",
      isinstance(st["corrections"], list) and len(st["corrections"]) == 1
      and st["corrections"][0]["before"] == 66
      and st["corrections"][0]["after"] == 92
      and st["corrections"][0]["delta"] == 26
      and st["corrections"][0]["source"] == "재채점"
      and st["corrections"][0]["set"]["norm"] == sets[0]["norm"]
      and st["corrections"][0]["record_id"] == "20260911100000", st["corrections"])
check("(f) 웹 진행률(level·passed_sets)도 같은 값으로 나감",
      st["level"]["best_total"] == 92 and st["plan"]["passed_sets"] == 2)

print()
print("8. (g) 손상된 기록 · 정정 파일에도 죽지 않음")
c = Case("case_g")
c.save(None, raw="{이건 JSON 이 아니다")
check("(g) 깨진 기록.json → 대상 없음, 예외 없음",
      c.coord().run()["checked"] == 0 and c.load() == [])
c.save([None, 3, "문자열", {"세트명": "합성 세트"},
        {"일시": "2026-09-05 22:05", "세트명": "합성 세트", "점수": "86점",
         "mode": 7, "정정이력": "리스트가 아님"}])
summary = c.coord().run()
check("(g) 이상한 항목이 섞여도 살아남음", isinstance(summary["changes"], list))
check("(g) serialize_corrections 는 이상한 정정이력을 건너뜀",
      sj.serialize_corrections(c.load(), [c.set]) == [])
weird = [{"일시": "2026-09-05 22:05", "세트명": "합성 세트", "점수": 80,
          "정정이력": [None, 5, {"새총점": None}, {"새총점": 90, "이전총점": "80점"}]}]
out = sj.serialize_corrections(weird, [c.set])
check("(g) 살릴 수 있는 정정이력만 살림", len(out) == 1 and out[0]["after"] == 90
      and out[0]["before"] == 80 and out[0]["delta"] == 10, out)
for raw in ("", "[]", "{}", "깨짐", '{"corrections": 3}',
            '{"corrections": [1, {"report_id": ""}, {"new_total": 5}]}'):
    c.write_corrections(None, raw=raw)
    check(f"(g) 정정 파일 방어: {raw[:24]!r} → 빈 목록",
          sj.load_corrections(c.corrections) == [])
check("(g) 정정 파일이 아예 없어도 []",
      sj.load_corrections(os.path.join(TMP, "없는파일.json")) == [])
c.write_corrections({"corrections": [{"report_id": "x", "new_total": "91점",
                                      "reason": None}]})
got = sj.load_corrections(c.corrections)
check("(g) 숫자 문자열 new_total 은 숫자로, 사유가 비면 기본 문구",
      got == [{"report_id": "x", "set": "", "new_total": 91, "delta": 0,
               "reason": "이의제기 인정"}], got)
c2 = Case("case_g2")
stu = c2.attempt()
c2.save([base_record(채점기버전="1.0.0", 풀이경로=stu)])
score_table(c2.grade_dir, {"_실패": True})
summary = c2.coord().run()
check("(g) 채점기가 실패해도 기록은 그대로, 오류만 모음",
      summary["regraded"] == 0 and len(summary["errors"]) == 1
      and c2.load()[0]["점수"] == 86, summary["errors"])
check("(g) grade.py 를 못 찾으면 조용히 실패 문구",
      sj.regrade_record({}, c2.set, grade_py=os.path.join(TMP, "없음.py"))[3]
      == "grade.py 를 찾을 수 없습니다")

print()
print("9. 풀이 파일 찾기 · 진행 표시 · 안내 문구 · 배포")
c = Case("legacy")
stu = c.attempt("20260905_2205")
rec = base_record(리포트=c.report(), 채점기버전="1.0.0")   # 풀이경로 없음(옛 기록)
c.save([rec])
check("풀이경로가 없는 옛 기록도 일시로 풀이_ 사본을 찾아 재채점",
      sj.attempt_file_for(rec, c.set) == stu)
summary = c.coord().run()
check("옛 기록 재채점 후 풀이경로가 기록에 저장됨",
      c.load()[0]["풀이경로"] == os.path.abspath(stu) and c.load()[0]["점수"] == 95)
check("일시를 못 읽으면 None", sj.attempt_file_for({"일시": "언젠가"}, c.set) is None)
seen = []
c2 = Case("progress")
recs = []
for i in range(3):
    s = c2.attempt(f"2026090{i + 1}_1000")
    recs.append(base_record(일시=f"2026-09-0{i + 1} 10:00", 점수=60 + i,
                            채점기버전="1.0.0", 풀이경로=s))
c2.save(recs)
c2.coord(progress=lambda i, n, nm: seen.append((i, n, nm))).run()
check("진행 콜백이 '2/3' 형태로 불림",
      seen == [(1, 3, "합성 세트"), (2, 3, "합성 세트"), (3, 3, "합성 세트")], seen)
check("여러 건 안내 문구는 목록 + '클릭하면 자세히'",
      sj.correction_summary_text([
          {"name": "코코 1회", "before": 86, "after": 95},
          {"name": "코코 2회", "before": 70, "after": 74},
          {"name": "A형", "before": 60, "after": 66},
          {"name": "B형", "before": 50, "after": 55}])
      == "이의제기 반영: 4건 — 코코 1회 86 → 95점 / 코코 2회 70 → 74점 / "
         "A형 60 → 66점 외 1건 (클릭하면 자세히)",
      sj.correction_summary_text([
          {"name": "코코 1회", "before": 86, "after": 95},
          {"name": "코코 2회", "before": 70, "after": 74},
          {"name": "A형", "before": 60, "after": 66},
          {"name": "B형", "before": 50, "after": 55}]))
check("정정 건이 없으면 빈 문구", sj.correction_summary_text([]) == "")
# 개발 저장소는 <루트>/정정/, 배포본은 스위트와 같은 폴더의 정정/ 이다.
_skel_cands = [os.path.join(os.path.dirname(BASE), "정정", "점수정정.json"),
               os.path.join(BASE, "정정", "점수정정.json")]
skel = next((c for c in _skel_cands if os.path.isfile(c)), _skel_cands[0])
skel_data = json.load(open(skel, encoding="utf-8")) \
    if os.path.isfile(skel) else None
check("저장소에 정정/점수정정.json 이 있고 형식이 맞음",
      isinstance(skel_data, dict)
      and isinstance(skel_data.get("corrections"), list), str(skel_data)[:120])
check("정정 항목마다 report_id 와 new_total/delta 중 하나가 있다",
      all(isinstance(it.get("report_id"), str) and it["report_id"].strip()
          and (it.get("new_total") is not None or it.get("delta"))
          for it in skel_data["corrections"]), str(skel_data)[:200])
# 배포.py 는 제작용 스크립트라 배포본에는 들어가지 않는다. 개발 저장소에서만 검사한다.
_dep_path = os.path.join(os.path.dirname(BASE), "제작스크립트", "배포.py")
if os.path.isfile(_dep_path):
    dep_spec = importlib.util.spec_from_file_location("dep", _dep_path)
    dep = importlib.util.module_from_spec(dep_spec)
    dep_spec.loader.exec_module(dep)
    items = dep.collect_corrections()
    check("배포.py 가 정정/점수정정.json 을 data_files 로 배포",
          items == [("정정/점수정정.json", "정정/점수정정.json", skel, "data", None)], items)
    info = dep.build_version_json("2.6.0", "테스트", items)
    check("version.json data_files 키 · sha256 포함",
          info["data_files"] == {"정정/점수정정.json": "정정/점수정정.json"}
          and "정정/점수정정.json" in info["sha256"] and info["set_files"] == {})
    bad = os.path.join(TMP, "나쁜정정.json")
    with open(bad, "w", encoding="utf-8") as f:
        json.dump({"corrections": [{"report_id": "x", "new_total": 120}]}, f)
    try:
        dep.check_corrections(bad)
        ok = False
    except ValueError:
        ok = True
    check("배포.py 는 0~100 밖 new_total 을 거른다", ok)
else:
    print("  (배포.py 없음 — 제작 스크립트 검사 건너뜀)")
check("설치 위치 키가 루트/정정/ 로 풀림",
      sj._data_target_path("정정/점수정정.json", sj.BASE_DIR)
      == os.path.join(os.path.dirname(sj.BASE_DIR), "정정", "점수정정.json"))

# --- 배포물은 우리가 만든 세트만 (저작권) -----------------------------------
if os.path.isfile(_dep_path):
    check("배포.py: 자체 제작 세트만 통과하는 관문이 있다",
          dep.is_own_set_file("코코모의고사1회_문제.xlsx") is True
          and dep.is_own_set_file("계산드릴_정답.xlsx") is True
          and dep.is_own_set_file("계산드릴_기대값.json") is True
          and dep.is_own_set_file("2024년상시1회_문제.xlsx") is False
          and dep.is_own_set_file("2026_1회_기대값.json") is False
          and dep.is_own_set_file("컴활2급상시_정답.xlsx") is False,
          str(dep.OWN_SET_PREFIXES))
    set_items = dep.collect_set_files()
    key_items = dep.collect_expected_values()
    dep_names = [os.path.basename(k) for k, _r, _s, _t, _v in set_items + key_items]
    check("배포 대상 세트·기대값이 전부 자체 제작 세트다 (기출 반입 차단)",
          dep_names and all(dep.is_own_set_file(n) for n in dep_names),
          ", ".join(sorted(dep_names)))
    check("배포 세트 목록은 코코 모의고사 1·2회 + 계산드릴뿐",
          sorted({n.split("_")[0] for n in dep_names})
          == ["계산드릴", "코코모의고사1회", "코코모의고사2회"],
          str(sorted({n.split("_")[0] for n in dep_names})))
    _stray = os.path.join(os.path.dirname(BASE), "기대값", "2099기출_기대값.json")
    _made = False
    try:
        if not os.path.exists(_stray):
            with open(_stray, "w", encoding="utf-8") as f:
                json.dump({"sheets": {}}, f)
            _made = True
        try:
            dep.collect_expected_values()
            gate_ok = False
        except SystemExit as e:
            gate_ok = "자체 제작이 아닌" in str(e) and "2099기출" in str(e)
    finally:
        if _made:
            os.remove(_stray)
    check("남의 기대값이 폴더에 섞이면 조용히 빼지 않고 배포를 멈춘다", gate_ok)

    shutil.rmtree(TMP, ignore_errors=True)
    print()
    print(f"자동 재채점 테스트 {N}건 전부 통과")
else:
    print("  (배포.py 없음 — 세트 반입 차단 검사 건너뜀)")
