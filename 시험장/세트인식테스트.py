# -*- coding: utf-8 -*-
"""시험장.py v2.3.1 세트 인식 헤드리스 테스트 (tkinter 불필요).

  1. 파일명 변형 가짜 트리 4종에서 슬롯 해석 (전 토큰/식별 토큰/미발견)
  2. 혼합 확장자(문제 .xlsx + 정답 .xlsm) 그룹핑
  3. 슬롯 매핑 저장 → resolve_day_sets 자동 해석
  4. 진단 텍스트 조립
  5. 오류 로그 append
  6. 문제지 PDF 회차·형·연도 검증 (저장된 잘못된 연결 무시, 애매한 퍼지 매칭 금지)
  7. 기대값 JSON 퍼지 연결 (v2.2.2)
  8. 루트/기대값 폴더 자동 연결 우선순위 (v2.3.0): 세트 폴더 정확 키 > 기대값 폴더
     정확 키 > 토큰 유일 일치, 스캔 루트 밖 key_dirs, 직접 선택, 진단 표시
  9. (v2.3.1) 사용자 PC 진단 트리(2026-09-06) 재현 픽스처 — 슬롯 9개 배정 기대표,
     PDF·기대값 연결, 무관 PDF 접기, 중복 병합/비병합 두 변형
 10. (v2.3.1) 중복 세트 병합: 대표 선택(PDF·기대값 붙은 쪽 > 얕은 폴더), 토큰 합집합,
     세트설정 항목 반영, 진단 표시
 11. (v2.3.1) PDF·기대값 전역 유일 연결: 동점 미연결+사유, 더 잘 맞는 세트가 있으면
     제외, 한 세트 두 PDF 동점, 연도만 겹치는 무관 파일, 토큰 1개 비포함 파일
 12. (v2.3.1) 자동 배포된 루트/모의고사/ 코코 세트가 '코코 1회·2회' 슬롯에 배정
 13. (v2.3.1) 전역 유일 배정: 확실한 슬롯 우선, 동점 미배정, 직접 선택 매핑 최우선,
     match_slot ↔ assign_slots 일관성, 식별 규칙 표
"""
import importlib.util
import json
import os
import sys
import tempfile
import zipfile

BASE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("sj", os.path.join(BASE, "시험장.py"))
sj = importlib.util.module_from_spec(spec)
spec.loader.exec_module(sj)
N = 0
sj.STARTUP_LOG_PATH = os.path.join(tempfile.mkdtemp(), "시작로그.txt")   # 실제 로그 오염 방지
_ISOLATED_KEY_DIR = os.path.join(tempfile.mkdtemp(), "기대값")            # 실제 루트/기대값 폴더 격리
sj.expected_values_dir = lambda base_dir=None: _ISOLATED_KEY_DIR


def check(desc, cond, extra=""):
    global N
    N += 1
    print(f"  [{N:02d}] {desc}: {'OK' if cond else 'FAIL'}" + (f" ({extra})" if extra else ""))
    if not cond:
        raise AssertionError(f"{desc} {extra}")


def touch_xlsx(path, content=None):
    """최소 xlsx(zip) 파일 생성 — 내용은 파일별로 고유(경로 문자열). content를 주면
    그 내용으로(중복 세트 재현용)."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("xl/workbook.xml", '<workbook><sheets><sheet name="기본작업-1"/></sheets></workbook>')
        zf.writestr("xl/id.txt", content if content is not None else path)
    return path


def make_tree(files):
    root = tempfile.mkdtemp()
    for rel in files:
        touch_xlsx(os.path.join(root, rel))
    return root


def touch_pdf(p):
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "wb") as f:
        f.write(b"%PDF-1.4\n")
    return p


def touch_json(p):
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        f.write("{}")
    return p


print("1. 파일명 변형 슬롯 해석")
V1 = make_tree(["세트/2024 A형 문제.xlsx", "세트/2024 A형 정답.xlsm",
                "세트/2024년 상시2회 2급_문제.xlsm", "세트/2024년 상시2회 2급_정답.xlsm"])
V2 = make_tree(["기출/컴활2급 A형 문제.xlsx", "기출/컴활2급 A형 정답.xlsm"])
V3 = make_tree(["A형/A형.xlsx", "A형/A형(정답).xlsm"])
V4 = make_tree(["세트/2024년 상시2회 2급_문제.xlsm", "세트/2024년 상시2회 2급_정답.xlsm"])
for tag, root in (("V1", V1), ("V2", V2), ("V3", V3), ("V4", V4)):
    sets = sj.scan_sets(root, {})
    s, how = sj.match_slot(sets, "2024 A형")
    if tag == "V1":
        check("① '2024 A형 문제.xlsx'+'정답.xlsm' → 전체 토큰 일치", s and s["name"] == "2024 A형" and "전체 토큰" in how, how)
        check("①  혼합 확장자(xlsx 문제 + xlsm 정답) 한 세트로 그룹핑",
              s["problem"].endswith(".xlsx") and s["answer"].endswith(".xlsm"))
    elif tag == "V2":
        check("② '컴활2급 A형 문제.xlsx'(2024 없음) → 식별 토큰 a형 일치",
              s and s["name"] == "컴활2급 A형" and "식별 토큰" in how, how)
    elif tag == "V3":
        check("③ 'A형.xlsx'+'A형(정답).xlsm' → 역할 추정 + 식별 토큰 일치",
              s and s["name"] == "A형" and "식별 토큰" in how, how)
    else:
        check("④ A형 파일 없음 → 미발견(후보 없음)", s is None and "후보 없음" in how, how)
# 연도 충돌 후보는 제외되고(2025 A형), 남은 후보가 유일하면 배정 — v2.3.1
V5 = make_tree(["a/2025 A형 문제.xlsx", "a/2025 A형 정답.xlsx", "b/복원 A형 문제.xlsx", "b/복원 A형 정답.xlsx"])
s, how = sj.match_slot(sj.scan_sets(V5, {}), "2024 A형")
check("2025 A형은 연도 충돌로 제외, 복원 A형만 남아 식별 토큰 일치 (v2.3.1)",
      s is not None and s["name"] == "복원 A형" and "식별 토큰" in how, how)
V5b = make_tree(["a/복원 A형 문제.xlsx", "a/복원 A형 정답.xlsx", "b/기출 A형 문제.xlsx", "b/기출 A형 정답.xlsx"])
s, how = sj.match_slot(sj.scan_sets(V5b, {}), "2024 A형")
check("동점 복수 후보(복원 A형·기출 A형)는 미발견 + 후보 나열 + 직접 선택 안내",
      s is None and "복수 후보" in how and "복원 A형" in how and "기출 A형" in how and "직접 선택" in how, how)
# 다른 슬롯: 24 2급 상시 vs 컴활 2급 상시 vs 2024 상시 1회 구분
V6 = make_tree(["x/24년 2급 상시 문제.xlsx", "x/24년 2급 상시 정답.xlsm",
                "x/컴활 2급 상시 문제.xlsx", "x/컴활 2급 상시 정답.xlsm",
                "x/2024년 상시1회 2급 문제.xlsx", "x/2024년 상시1회 2급 정답.xlsm",
                "x/코코모의고사1회_문제.xlsx", "x/코코모의고사1회_정답.xlsx"])
sets6 = sj.scan_sets(V6, {})
r = {spec: sj.match_slot(sets6, spec) for spec in ("24 2급 상시", "컴활 2급 상시", "2024 상시 1회", "코코 1회", "2024 B형")}
check("'24 2급 상시' → 24년 2급 상시", r["24 2급 상시"][0] and "24년" in r["24 2급 상시"][0]["name"], r["24 2급 상시"][1])
check("'컴활 2급 상시' → 컴활 2급 상시", r["컴활 2급 상시"][0] and r["컴활 2급 상시"][0]["name"].startswith("컴활"), r["컴활 2급 상시"][1])
check("'2024 상시 1회' → 2024년 상시1회 2급", r["2024 상시 1회"][0] and "상시1회" in r["2024 상시 1회"][0]["name"], r["2024 상시 1회"][1])
check("'코코 1회' → 코코모의고사1회", r["코코 1회"][0] and "코코" in r["코코 1회"][0]["name"], r["코코 1회"][1])
check("'2024 B형' → 미발견", r["2024 B형"][0] is None)
check("2자리 연도/컴활 토큰 추출", {"24", "2급", "상시"} <= sj.set_tokens("24년 2급 상시") and "컴활" in sj.set_tokens("컴활 2급 상시")
      and "24" not in sj.set_tokens("2024년 상시1회"))
check("식별 규칙 표: 슬롯 9개 모두 핵심 묶음 보유 + 표시 토큰 + 표 밖 문구는 연도 제외 토큰",
      all("핵심" in v and v["핵심"] for v in sj.SLOT_IDENTITY.values()) and len(sj.SLOT_IDENTITY) == 9
      and sj.slot_identity_tokens("24 2급 상시") == {"24/2024", "2급", "모의/실기/상시"}
      and sj.slot_identity_tokens("2024 상시 1회") == {"2024/24", "1회"}
      and sj.slot_identity_tokens("계산 드릴") == {"계산", "드릴"}
      and sj.SLOT_IDENTITY["컴활 2급 상시"].get("연도제외") is True)

print("2. 2슬롯 plan 해석: 한 슬롯 미발견 + 매핑 대체")
sets4 = sj.scan_sets(V4, {})
# 하루 1세트이고 v3.0.0 부터는 일정표에 고정 세트가 없다(엔진이 보유 세트에서
# 고른다) — 2슬롯 처리 경로는 Day 3 카드에 세트 이름을 직접 넣은 합성 plan 으로
# 확인한다.
check("일정표는 하루 1세트 · 고정 세트 없음 (Day 3 = 자동 선택)",
      sj.plan_for_day(3)["세트"] == [sj.AUTO], str(sj.plan_for_day(3)["세트"]))
plan3 = dict(sj.plan_for_day(3), 세트=["2024 A형", "2024 상시 2회"], 목표들=[70, 70])
plan3["스텝"] = sj.build_day_steps(plan3)
rs = sj.resolve_day_sets(plan3, sets4, [], None)
check("슬롯1(A형) 미발견 + 사유, 슬롯2(상시 2회) 정상",
      rs[0][0] is None and "찾지 못함" in rs[0][1] and "후보 없음" in rs[0][1]
      and rs[1][0] is not None and "상시2회" in rs[1][0]["name"], str([r[1] for r in rs]))
step_exam2 = plan3["스텝"][4]
kind, payload = sj.resolve_step_action(step_exam2, sets4, slot_sets=rs)
check("정상 슬롯(2)의 모의 스텝은 exam", kind == "exam" and "상시2회" in payload["set"]["name"])
kind, payload = sj.resolve_step_action(plan3["스텝"][0], sets4, slot_sets=rs)
check("미발견 슬롯(1)의 모의 스텝은 missing + 직접 선택 안내", kind == "missing" and "다른 세트로 바꾸기" in payload["이유"])
cfgp = os.path.join(tempfile.mkdtemp(), "세트설정.json")
Aset = sj.scan_sets(V2, {})[0]
sj.save_slot_mapping("2024 A형", Aset, cfgp)
mp = sj.load_slot_mapping(cfgp)
check("슬롯 매핑 저장/로드", "2024 A형" in mp and mp["2024 A형"]["problem"] == Aset["problem"])
rs2 = sj.resolve_day_sets(plan3, sets4, [], None, mapping=mp)
check("매핑으로 슬롯1 자동 해석(직접 선택 저장됨) + 슬롯2 유지",
      rs2[0][0] is not None and rs2[0][0]["name"] == "컴활2급 A형" and "직접 선택 저장됨" in rs2[0][1]
      and rs2[1][0]["name"] == rs[1][0]["name"], str([r[1] for r in rs2]))
os.remove(Aset["problem"])
check("매핑 파일이 사라지면 무시하고 다시 매칭 시도", sj.resolve_day_sets(plan3, sets4, [], None, mapping=mp)[0][0] is None)

print("3. 진단 텍스트")
V7 = make_tree(["세트/2024 A형 문제.xlsx", "세트/2024 A형 정답.xlsm", "세트/외톨이_문제.xlsx",
                "세트/2024년 상시2회 2급_문제.xlsm", "세트/2024년 상시2회 2급_정답.xlsm"])
touch_pdf(os.path.join(V7, "세트", "2024 A형 문제지.pdf"))
txt = sj.scan_diagnosis_text(V7, config={"_슬롯매핑": {"코코 1회": {"problem": "/없음/x.xlsx", "answer": "/없음/y.xlsx"}}})
print("----- 진단 텍스트 예시 -----")
print(txt)
print("---------------------------")
lines = txt.splitlines()
check("머리: 스캔 루트 + 4개 섹션", txt.startswith("[코코 시험장 세트 인식 진단]") and V7 in txt
      and all(h in txt for h in ("== 파일 → 정규화 키 → 역할 → 소속 세트 ==", "== 인식된 세트 ==", "== 일정 슬롯 매칭", "== 미소속 Excel 파일")))
check("파일 행: 이름 | 키 | 역할 | 세트", any(l.startswith("2024 A형 문제.xlsx | 2024a형 | 문제 | 2024 A형") for l in lines))
check("정답 xlsm 행이 같은 세트 소속", any(l.startswith("2024 A형 정답.xlsm | 2024a형 | 정답 | 2024 A형") for l in lines))
check("PDF 행 역할 문제지", any("2024 A형 문제지.pdf | 2024a형 | 문제지(PDF) | 2024 A형" in l for l in lines))
check("미소속 파일 표시", any(l.startswith("외톨이_문제.xlsx | 외톨이 | 문제 | 미소속") for l in lines) and "외톨이_문제.xlsx" in txt.split("== 미소속")[1])
check("슬롯 매칭 행: A형 전체 토큰(+PDF·기대값 표시), 상시 2회 전체 토큰, B형 미발견 사유",
      any("슬롯 '2024 A형' → 세트 '2024 A형' (전체 토큰 일치" in l and "PDF 2024 A형 문제지.pdf" in l and "기대값 없음" in l for l in lines)
      and any("슬롯 '2024 상시 2회' → 세트" in l for l in lines)
      and any("슬롯 '2024 B형' → 미발견 (후보 없음" in l for l in lines))
check("저장된 매핑(파일 없음)은 '무시' 표시", any("슬롯 '코코 1회' → 미발견" in l and "직접 선택 저장됨" in l and "파일 없음" in l for l in lines))
check("빈 루트도 안전", "인식된 세트 없음" in sj.scan_diagnosis_text(tempfile.mkdtemp(), config={}))

print("4. 오류 로그")
lp = os.path.join(tempfile.mkdtemp(), "시험장_오류.log")
try:
    raise ValueError("테스트 예외")
except ValueError as e:
    t = sj.log_error("단위 테스트", e, path=lp)
body = open(lp, encoding="utf-8").read()
check("로그 append: 문맥·버전·traceback", "단위 테스트" in body and sj.__version__ in body and "ValueError: 테스트 예외" in body and t in body)
sj.log_error("두 번째", RuntimeError("x"), path=lp)
check("두 번째 기록 append (구분선 2개)", open(lp, encoding="utf-8").read().count("-" * 60) == 2)

print("6. 문제지 PDF 회차 검증 (v2.2.1)")
V8 = make_tree(["기출/2024년 상시1회 2급_문제.xlsm", "기출/2024년 상시1회 2급_정답.xlsm",
                "기출/2024년 상시2회 2급_문제.xlsm", "기출/2024년 상시2회 2급_정답.xlsm"])
pdf1 = touch_pdf(os.path.join(V8, "기출", "2024년 상시1회 2급 문제지.pdf"))   # 같은 키 → 1회
amb = touch_pdf(os.path.join(V8, "기출", "2024 상시 2급 문제지.pdf"))        # 회차 없음 → 1회·2회 동점
sets8 = sj.scan_sets(V8, {})
s1 = next(s for s in sets8 if "상시1회" in s["name"])
s2 = next(s for s in sets8 if "상시2회" in s["name"])
check("같은 키 PDF는 1회 세트에 연결", s1["pdf"] == pdf1)
check("회차 없는 PDF는 1회·2회에 동점이지만 1회는 문제지가 이미 있어 제외 → 2회에 연결 (v2.3.1)",
      s2["pdf"] == amb, str(s2["pdf"]))
check("pdf_conflicts_with_set: 2회 세트 ← 1회 문제지만 충돌",
      sj.pdf_conflicts_with_set(s2, pdf1) and not sj.pdf_conflicts_with_set(s1, pdf1)
      and not sj.pdf_conflicts_with_set(s2, amb) and not sj.pdf_conflicts_with_set(s2, None))
cfg8 = {s2["norm"]: {"name": s2["name"], "problem": s2["problem"], "answer": s2["answer"],
                     "key": None, "pdf": pdf1}}
sets8b = sj.scan_sets(V8, cfg8)
s2b = next(s for s in sets8b if "상시2회" in s["name"])
check("세트설정에 저장된 1회 PDF(회차 충돌)는 무시 + 경고 사유 (자동 연결은 유지)",
      s2b["pdf"] == amb and "상시1회" in (s2b.get("pdf_warning") or ""), str(s2b.get("pdf_warning")))
check("시작 로그에 '문제지 연결 무시' 기록",
      "문제지 연결 무시(회차 불일치, 세트설정)" in open(sj.STARTUP_LOG_PATH, encoding="utf-8").read())
cfg8c = {s2["norm"]: dict(cfg8[s2["norm"]], **{"pdf_확인됨": True})}
s2c = next(s for s in sj.scan_sets(V8, cfg8c) if "상시2회" in s["name"])
check("사용자가 확인한 연결(pdf_확인됨)은 존중", s2c["pdf"] == pdf1 and s2c.get("pdf_확인됨") is True)
# 저장 전용 세트(스캔에 없는 경로)도 같은 검증
V8s = make_tree(["딴곳/2026 모의고사 1회_문제.xlsm", "딴곳/2026 모의고사 1회_정답.xlsm"])
pdf26_2 = touch_pdf(os.path.join(V8s, "딴곳", "2026 모의고사 2회 문제지.pdf"))
saved_cfg = {"2026모의고사1회": {"name": "2026 모의고사 1회",
                              "problem": os.path.join(V8s, "딴곳", "2026 모의고사 1회_문제.xlsm"),
                              "answer": os.path.join(V8s, "딴곳", "2026 모의고사 1회_정답.xlsm"),
                              "pdf": pdf26_2, "key": ["잘못된 타입"]}}
ss = sj.apply_set_config([], saved_cfg)
check("저장 전용 세트: 회차 충돌 PDF 무시, 잘못된 타입 key 무시",
      len(ss) == 1 and ss[0]["pdf"] is None and ss[0]["key"] is None and ss[0].get("pdf_warning") and ss[0]["saved"])
# 슬롯 매핑
m = {"name": s2["name"], "problem": s2["problem"], "answer": s2["answer"], "pdf": pdf1}
ms = sj.set_from_mapping(m)
check("슬롯 매핑의 충돌 PDF 무시 (+경고 사유, direct 표시)",
      ms is not None and ms["pdf"] is None and ms.get("pdf_warning") and ms.get("direct") is True)
check("슬롯 매핑 pdf_확인됨 존중", sj.set_from_mapping(dict(m, **{"pdf_확인됨": True}))["pdf"] == pdf1)
m2 = {"name": s1["name"], "problem": s1["problem"], "answer": s1["answer"], "pdf": pdf1}
check("회차 일치 PDF 매핑은 그대로", sj.set_from_mapping(m2)["pdf"] == pdf1)
check("매핑 pdf 타입 이상(리스트)은 무시하고 세트는 살림",
      sj.set_from_mapping(dict(m2, pdf=["x"]))["pdf"] == pdf1)   # 같은 키 자동 감지로 1회 PDF
txt8 = sj.scan_diagnosis_text(V8, config=cfg8, sets=sets8b)
check("진단 텍스트에 무시 사유(⚠) 표시", "⚠" in txt8 and "무시했습니다" in txt8)
bad_live = dict(s2, pdf=pdf1)
txt8b = sj.scan_diagnosis_text(V8, config={}, sets=[bad_live])
check("진단 텍스트: 현재 연결이 충돌이면 '⚠회차불일치'", "⚠회차불일치" in txt8b)
# 퍼지 매칭 유일성: A형 문제지는 A형에만
V9 = make_tree(["s/2024 A형 문제.xlsx", "s/2024 A형 정답.xlsm", "s/2024 B형 문제.xlsx", "s/2024 B형 정답.xlsm"])
touch_pdf(os.path.join(V9, "s", "2024년 A형 기출 문제지.pdf"))     # 키 다름 → 퍼지 매칭
sets9 = sj.scan_sets(V9, {})
sA = next(s for s in sets9 if "A형" in s["name"])
sB = next(s for s in sets9 if "B형" in s["name"])
check("퍼지 매칭: A형 문제지는 A형에만 (B형은 형 충돌)", sA["pdf"] and sA["pdf"].endswith("기출 문제지.pdf") and sB["pdf"] is None)
check("match_pdf_for_set: 다른 세트와 동점이면 None, 더 잘 맞으면 연결",
      sj.match_pdf_for_set({"2024", "상시", "2회"}, ["/x/2024 상시 문제지.pdf"], others=[{"2024", "상시", "1회"}]) is None
      and sj.match_pdf_for_set({"2024", "상시", "2회"}, ["/x/2024 상시 2회 문제지.pdf"], others=[{"2024", "상시", "1회"}])
      == "/x/2024 상시 2회 문제지.pdf")
check("PDF_MISMATCH_WARNING 문구", "회차가 세트와 다릅니다" in sj.PDF_MISMATCH_WARNING and "[문제지 연결]" in sj.PDF_MISMATCH_WARNING)
check("_token_conflict: 2자리 연도 '24'는 2024와 같은 해, 2026과는 충돌 (v2.3.1)",
      not sj._token_conflict({"24", "2급"}, {"2024", "1회"}) and sj._token_conflict({"24", "2급"}, {"2026", "1회"})
      and sj._year_tokens({"24", "2026", "2급"}) == {"2024", "2026"})

print("7. 기대값 JSON 퍼지 연결 (v2.2.2)")
V10 = make_tree(["k/2026 컴활2급실기_상시복원기출문제 1회_문제.xlsm",
                 "k/2026 컴활2급실기_상시복원기출문제 1회_정답.xlsm",
                 "k/2024 상시 2회_문제.xlsx", "k/2024 상시 2회_정답.xlsm"])
touch_json(os.path.join(V10, "k", "2026_1회_기대값.json"))       # 키 다름 → 토큰(2026·1회)
touch_json(os.path.join(V10, "k", "2024 상시 2회_기대값.json"))   # 같은 키
sets10 = sj.scan_sets(V10, {})
s26 = next(s for s in sets10 if "2026" in s["name"])
s24 = next(s for s in sets10 if "2024" in s["name"])
check("토큰(2026·1회) 퍼지 매칭으로 기대값 JSON 연결",
      bool(s26["key"]) and s26["key"].endswith("2026_1회_기대값.json"), str(s26["key"]))
check("같은 키 기대값은 그대로 연결",
      bool(s24["key"]) and s24["key"].endswith("2024 상시 2회_기대값.json"), str(s24["key"]))
d10 = sj.build_direct_set(s26["problem"], s26["answer"])
check("직접 선택(build_direct_set)도 퍼지 연결",
      bool(d10["key"]) and d10["key"].endswith("2026_1회_기대값.json"), str(d10["key"]))
check("2024 상시 2회 직접 선택은 2026 기대값에 연결 안 함(연도 충돌)",
      sj.build_direct_set(s24["problem"], s24["answer"])["key"].endswith("2024 상시 2회_기대값.json"))
V11 = make_tree(["k/2026 상시 1회_문제.xlsm", "k/2026 상시 1회_정답.xlsm",
                 "k/2026 상시 2회_문제.xlsm", "k/2026 상시 2회_정답.xlsm"])
touch_json(os.path.join(V11, "k", "2026_기대값.json"))            # 회차 없음 → 애매
sets11 = sj.scan_sets(V11, {})
check("회차 없는 기대값은 1회·2회 모두에 맞아 애매 → 연결 안 함 + 동점 사유",
      all(s["key"] is None for s in sets11) and all("2026_기대값.json" in (s.get("key_tie") or "") for s in sets11),
      str([(s["key"], s.get("key_tie")) for s in sets11]))
check("직접 선택도 토큰 1개(연도만) 겹치는 기대값은 연결 안 함",
      sj.build_direct_set(os.path.join(V11, "k", "2026 상시 1회_문제.xlsm"),
                          os.path.join(V11, "k", "2026 상시 1회_정답.xlsm"))["key"] is None)

print("8. 루트/기대값 폴더 자동 연결 (v2.3.0)")
V12 = make_tree(["모의고사/코코모의고사1회_문제.xlsx", "모의고사/코코모의고사1회_정답.xlsx",
                 "기출/코코모의고사2회_문제.xlsx", "기출/코코모의고사2회_정답.xlsx",
                 "세트/2024년 상시1회 2급_문제.xlsm", "세트/2024년 상시1회 2급_정답.xlsm",
                 "세트/2024년 상시2회 2급_문제.xlsm", "세트/2024년 상시2회 2급_정답.xlsm",
                 "세트/2026 컴활2급실기_상시복원기출문제 1회_문제.xlsm",
                 "세트/2026 컴활2급실기_상시복원기출문제 1회_정답.xlsm"])
KD = os.path.join(V12, "기대값")
os.makedirs(KD)
with open(os.path.join(V12, "모의고사", "코코모의고사1회_기대값.json"), "w", encoding="utf-8") as f:
    f.write('{"own": 1}')
for fn in ("코코모의고사1회_기대값.json", "코코모의고사2회_기대값.json",
           "2024상시2회_기대값.json", "2026_1회_기대값.json"):
    touch_json(os.path.join(KD, fn))
sets12 = sj.scan_sets(V12, {}, key_dirs=[KD])
by = {s["name"]: s for s in sets12}
check("세트 폴더의 정확 키가 기대값 폴더의 같은 이름보다 우선 (출처 '세트 폴더')",
      by["코코모의고사1회"]["key"] == os.path.join(V12, "모의고사", "코코모의고사1회_기대값.json")
      and by["코코모의고사1회"]["key_origin"] == "세트 폴더", str(by["코코모의고사1회"]["key"]))
check("세트 폴더에 없으면 기대값 폴더의 정확 키 (출처 '기대값 폴더')",
      by["코코모의고사2회"]["key"] == os.path.join(KD, "코코모의고사2회_기대값.json")
      and by["코코모의고사2회"]["key_origin"] == "기대값 폴더")
check("정확 키가 없으면 토큰(2024·상시·2회) 유일 일치 (출처 '토큰 일치')",
      by["2024년 상시2회 2급"]["key"] == os.path.join(KD, "2024상시2회_기대값.json")
      and by["2024년 상시2회 2급"]["key_origin"] == "토큰 일치")
check("2026 복원 1회 ← 2026_1회_기대값.json (연도·회차 토큰)",
      by["2026 컴활2급실기_상시복원기출문제 1회"]["key"] == os.path.join(KD, "2026_1회_기대값.json"))
check("상시 1회는 회차/연도가 충돌해 어떤 기대값에도 연결되지 않음",
      by["2024년 상시1회 2급"]["key"] is None and by["2024년 상시1회 2급"]["key_origin"] is None)
check("key_link_label: 세트 폴더 정확 키는 파일명만, 그 밖은 '(자동 연결)', 없으면 '없음'",
      sj.key_link_label(by["코코모의고사1회"]) == "코코모의고사1회_기대값.json"
      and sj.key_link_label(by["코코모의고사2회"]) == "코코모의고사2회_기대값.json (자동 연결)"
      and sj.key_link_label(by["2024년 상시2회 2급"]) == "2024상시2회_기대값.json (자동 연결)"
      and sj.key_link_label(by["2024년 상시1회 2급"]) == "없음")
# 기대값 폴더가 스캔 루트 밖에 있어도(--scan-root 지정) 후보에 포함
sets12b = sj.scan_sets(os.path.join(V12, "기출"), {}, key_dirs=[KD])
check("스캔 루트 밖의 key_dirs 기대값도 연결", sets12b[0]["key"] == os.path.join(KD, "코코모의고사2회_기대값.json")
      and sets12b[0]["key_origin"] == "기대값 폴더")
check("key_dirs 없이(기본 expected_values_dir가 없는 환경) 스캔해도 안전",
      sj.scan_sets(os.path.join(V12, "기출"), {}, key_dirs=[os.path.join(V12, "없는폴더")])[0]["key"] is None)
# 직접 선택
d12 = sj.build_direct_set(by["코코모의고사2회"]["problem"], by["코코모의고사2회"]["answer"], key_dirs=[KD])
check("직접 선택: 기대값 폴더의 정확 키", d12["key"] == os.path.join(KD, "코코모의고사2회_기대값.json")
      and d12["key_origin"] == "기대값 폴더")
d12b = sj.build_direct_set(by["2026 컴활2급실기_상시복원기출문제 1회"]["problem"],
                           by["2026 컴활2급실기_상시복원기출문제 1회"]["answer"], key_dirs=[KD])
check("직접 선택: 기대값 폴더의 토큰 일치(2026·1회)", d12b["key"] == os.path.join(KD, "2026_1회_기대값.json")
      and d12b["key_origin"] == "토큰 일치")
d12c = sj.build_direct_set(by["코코모의고사1회"]["problem"], by["코코모의고사1회"]["answer"], key_dirs=[KD])
check("직접 선택: 세트 폴더 정확 키 우선", d12c["key"].startswith(os.path.join(V12, "모의고사")) and d12c["key_origin"] == "세트 폴더")
check("직접 선택: 상시 1회는 연결 안 함",
      sj.build_direct_set(by["2024년 상시1회 2급"]["problem"], by["2024년 상시1회 2급"]["answer"], key_dirs=[KD])["key"] is None)
# 진단 텍스트에 출처 표시
txt12 = sj.scan_diagnosis_text(V12, config={}, sets=sets12)
check("진단 텍스트: 기대값 출처 표시", "코코모의고사2회_기대값.json (기대값 폴더)" in txt12
      and "2024상시2회_기대값.json (토큰 일치)" in txt12
      and "| 코코모의고사1회_기대값.json |" in txt12)
# 저장된 세트설정의 key가 다른 파일이면 '세트설정' 출처
cfg12 = {by["코코모의고사2회"]["norm"]: {"name": "코코모의고사2회", "problem": by["코코모의고사2회"]["problem"],
                                    "answer": by["코코모의고사2회"]["answer"],
                                    "key": os.path.join(V12, "모의고사", "코코모의고사1회_기대값.json"), "pdf": None}}
s12s = next(s for s in sj.scan_sets(V12, cfg12, key_dirs=[KD]) if s["name"] == "코코모의고사2회")
check("세트설정에 저장된 기대값 경로는 그대로 존중 (출처 '세트설정', 자동 연결 표시)",
      s12s["key"].endswith("코코모의고사1회_기대값.json") and s12s["key_origin"] == "세트설정"
      and "(자동 연결)" in sj.key_link_label(s12s))
V13 = make_tree(["k/2026 컴활2급실기_상시복원기출문제 1회_문제.xlsm",
                 "k/2026 컴활2급실기_상시복원기출문제 1회_정답.xlsm"])
KD13 = os.path.join(V13, "기대값")
os.makedirs(KD13)
touch_json(os.path.join(V13, "k", "2026_1회_기대값.json"))      # 세트 폴더 사본
touch_json(os.path.join(KD13, "2026_1회_기대값.json"))          # 배포본 (같은 이름)
s13 = sj.scan_sets(V13, {}, key_dirs=[KD13])[0]
check("같은 이름의 기대값이 세트 폴더·기대값 폴더에 모두 있으면 동점이 아니라 세트 폴더 사본 연결",
      s13["key"] == os.path.join(V13, "k", "2026_1회_기대값.json") and s13["key_origin"] == "토큰 일치", str(s13["key"]))
d13 = sj.build_direct_set(s13["problem"], s13["answer"], key_dirs=[KD13])
check("직접 선택도 세트 폴더 사본 우선", d13["key"] == os.path.join(V13, "k", "2026_1회_기대값.json"))
import importlib.util as _ilu          # 격리 전 원본 함수로 검사
_spec2 = _ilu.spec_from_file_location("sj2", os.path.join(BASE, "시험장.py"))
_sj2 = _ilu.module_from_spec(_spec2)
_spec2.loader.exec_module(_sj2)
check("expected_values_dir: 시험장 폴더 구조 → 상위/기대값, 평면 구조 → 폴더/기대값",
      _sj2.expected_values_dir(os.path.join(V12, "시험장")) == os.path.join(V12, "기대값")
      and _sj2.expected_values_dir(os.path.join(V12, "평면")) == os.path.join(V12, "평면", "기대값"))

print("9. 사용자 PC 진단 트리 재현 (2026-09-06, v2.3.1)")
KIC = "컴활/24년+개정판+2급+실기+한번더+최신기출문제+실습+파일"
DIAG_XLS = [
    "컴활/2026 컴활2급실기_상시복원기출문제 1회_문제.xlsm", "컴활/2026 컴활2급실기_상시복원기출문제 1회_정답.xlsm",
    "컴활/상시기출2회(문제).xlsm", "컴활/상시기출2회(정답).xlsm",
    "컴활/24년 개정판 2급 실기 모의고사 실습 파일/24년 개정판 2급 실기 모의고사(문제).xlsm",
    "컴활/24년 개정판 2급 실기 모의고사 실습 파일/24년 개정판 2급 실기 모의고사(정답).xlsm",
    f"{KIC}/1. 기출(문제) 24년 A형/2024년 상공회의소 샘플 A형(문제).xlsx",
    f"{KIC}/1. 기출(문제) 24년 A형/2024년 상공회의소 샘플 A형(정답).xlsm",
    f"{KIC}/2. 기출(문제) 24년 B형/2024년 상공회의소 샘플 B형(문제).xlsx",
    f"{KIC}/2. 기출(문제) 24년 B형/2024년 상공회의소 샘플 B형(정답).xlsm",
    f"{KIC}/3. 기출(문제) 24년 1회/2024년 기출문제 유형 1회(문제).xlsm",
    f"{KIC}/3. 기출(문제) 24년 1회/2024년 기출문제 유형 1회(정답).xlsm",
    f"{KIC}/4. 기출(문제) 24년 2회/2024년 기출문제 유형 2회(문제).xlsm",
    f"{KIC}/4. 기출(문제) 24년 2회/2024년 기출문제 유형 2회(정답).xlsm",
]
DIAG_PDF = [
    "2024A형_문제지.pdf", "2024B형_문제지.pdf", "2024상시1회_문제지.pdf", "24_2급상시_문제지.pdf",
    "[기공자] 2026 컴활2급실기_상시복원기출문제 1회.pdf",
    "単位修得状況照会 [CampusSquare].pdf", "履修成績照会 [CampusSquare].pdf",
    "국립창원대학교 등록금 고지서.pdf", "아이모2026_주제발표_초안.pdf",
    "컴활2급상시_문제지.pdf", "코코모의고사1회_문제지.pdf", "코코모의고사2회_문제지.pdf",
    "컴활/2024상시2회_문제지.pdf", "컴활/2026 컴활2급실기_상시복원기출문제 문제.pdf", "컴활/계산드릴_문제지.pdf",
]
DIAG_KEYS = ["2024상시2회_기대값.json", "2026_1회_기대값.json", "코코모의고사1회_기대값.json", "코코모의고사2회_기대값.json"]
SLOTS9 = ["2024 상시 1회", "코코 1회", "2024 A형", "2024 상시 2회", "코코 2회", "2024 B형",
          "24 2급 상시", "컴활 2급 상시", "2026 1회"]


def make_diag_tree(same_content=True, with_coco=False):
    """진단 출력의 파일 트리(파일명만 같은 합성 파일). same_content=True면
    '상시기출2회(문제).xlsm'을 '2024년 기출문제 유형 2회(문제).xlsm'과 같은 내용으로."""
    root = tempfile.mkdtemp(prefix="바탕화면_")
    for rel in DIAG_XLS:
        content = "동일내용-2024기출2회문제" if same_content and rel.endswith(
            ("상시기출2회(문제).xlsm", "유형 2회(문제).xlsm")) else None
        touch_xlsx(os.path.join(root, rel), content)
    for rel in DIAG_PDF:
        touch_pdf(os.path.join(root, rel))
    kd = os.path.join(root, "기대값")
    for fn in DIAG_KEYS:
        touch_json(os.path.join(kd, fn))
    if with_coco:          # v2.3.1 set_files 배포 결과: 루트/모의고사/
        for n in (1, 2):
            touch_xlsx(os.path.join(root, "모의고사", f"코코모의고사{n}회_문제.xlsx"))
            touch_xlsx(os.path.join(root, "모의고사", f"코코모의고사{n}회_정답.xlsx"))
            touch_pdf(os.path.join(root, "모의고사", f"코코모의고사{n}회_문제지.pdf"))
    return root, kd


def table(sets, specs=SLOTS9, mapping=None):
    """{슬롯: (세트명, PDF 파일명, 기대값 파일명, 사유)}"""
    out = {}
    for spec, (s, how) in sj.assign_slots(sets, specs, mapping).items():
        if spec in specs:
            out[spec] = ((s["name"] if s else None),
                         (os.path.basename(s["pdf"]) if s and s.get("pdf") else None),
                         (os.path.basename(s["key"]) if s and s.get("key") else None), how)
    return out


D1, KD1 = make_diag_tree(same_content=True)
setsD = sj.scan_sets(D1, {}, key_dirs=[KD1])
names = sorted(s["name"] for s in setsD)
check("세트 6개 인식(중복 병합으로 7→6)", len(setsD) == 6 and "2024년 기출문제 유형 2회" not in names, str(names))
rep = next(s for s in setsD if s["name"] == "상시기출2회")
check("중복 병합: 대표 '상시기출2회'(얕은 폴더) + 중복 목록에 '2024년 기출문제 유형 2회'",
      [d["name"] for d in rep.get("중복") or []] == ["2024년 기출문제 유형 2회"])
check("대표 세트 토큰 = 두 이름·폴더의 합집합(2024·24·2회·기출·상시·컴활)",
      {"2024", "24", "2회", "기출", "상시", "컴활"} <= sj._set_tokens_of(rep), str(sorted(sj._set_tokens_of(rep))))
T = table(setsD)
for line in (f"{k} → {v[0]} · PDF {v[1]} · 기대값 {v[2]} · {v[3]}" for k, v in T.items()):
    print("      " + line)
EXPECT = {
    "2024 상시 1회": ("2024년 기출문제 유형 1회", "2024상시1회_문제지.pdf", None, "식별 토큰 일치"),
    "코코 1회": (None, None, None, "후보 없음"),
    "2024 A형": ("2024년 상공회의소 샘플 A형", "2024A형_문제지.pdf", None, "전체 토큰 일치"),
    "2024 상시 2회": ("상시기출2회", "2024상시2회_문제지.pdf", "2024상시2회_기대값.json", "전체 토큰 일치"),
    "코코 2회": (None, None, None, "후보 없음"),
    "2024 B형": ("2024년 상공회의소 샘플 B형", "2024B형_문제지.pdf", None, "전체 토큰 일치"),
    "24 2급 상시": ("24년 개정판 2급 실기 모의고사", "24_2급상시_문제지.pdf", None, "식별 토큰 일치"),
    "컴활 2급 상시": (None, None, None, "후보 없음"),
    "2026 1회": ("2026 컴활2급실기_상시복원기출문제 1회", "[기공자] 2026 컴활2급실기_상시복원기출문제 1회.pdf",
              "2026_1회_기대값.json", "전체 토큰 일치"),
}
for spec, (nm, pdf, key, how) in EXPECT.items():
    got = T[spec]
    check(f"슬롯 '{spec}' → {nm or '미발견'} / PDF {pdf or '없음'} / 기대값 {key or '없음'} ({how})",
          got[0] == nm and got[1] == pdf and got[2] == key and how in got[3], str(got))
check("문제 1: '2024 상시 1회'가 2026 세트에 붙지 않음(연도 충돌) — 사유에 2026 언급 없음",
      "2026" not in T["2024 상시 1회"][3])
check("문제 2: '컴활 2급 상시'는 2026 세트 제외(연도제외) → 후보 없음, 2026 세트는 '2026 1회' 한 슬롯만",
      T["컴활 2급 상시"][0] is None and sum(1 for v in T.values() if v[0] and v[0].startswith("2026")) == 1)
check("문제 6: 코코 1·2회는 xlsx가 없어 미발견(후보 없음 사유에 코코 토큰 안내)",
      "코코" in T["코코 1회"][3] and "코코" in T["코코 2회"][3])
one_set_one_slot = [v[0] for v in T.values() if v[0]]
check("한 세트는 한 슬롯(배정된 세트명 중복 없음)", len(one_set_one_slot) == len(set(one_set_one_slot)))
attached = {os.path.basename(s["pdf"]) for s in setsD if s.get("pdf")}
check("문제 5: 미소속으로 남는 PDF = 컴활2급상시(세트 없음)·2026 문제.pdf(사본)·계산드릴·코코 2종·아이모·무관 3종",
      attached == {"2024A형_문제지.pdf", "2024B형_문제지.pdf", "2024상시1회_문제지.pdf", "24_2급상시_문제지.pdf",
                   "[기공자] 2026 컴활2급실기_상시복원기출문제 1회.pdf", "2024상시2회_문제지.pdf"}, str(sorted(attached)))
check("연도만 겹치는 '아이모2026_주제발표_초안.pdf'는 2026 세트 후보가 아님",
      sj._file_link_score({"2026", "1회", "상시"}, "/x/아이모2026_주제발표_초안.pdf") is None
      and sj._file_link_score({"2026", "1회", "상시"}, "/x/2026_문제지.pdf") is not None)
txtD = sj.scan_diagnosis_text(D1, config={}, sets=setsD)
print("----- 진단 트리 진단 텍스트 -----")
print(txtD)
print("--------------------------------")
check("진단: 무관 PDF 3개(세트 토큰 없음)는 표에서 빠지고 한 줄로 접힘",
      "무관(세트 토큰 없음) PDF 3개" in txtD and "国立" not in txtD
      and all(n in txtD.split("무관(세트 토큰 없음)")[1].splitlines()[0]
              for n in ("単位修得状況照会 [CampusSquare].pdf", "履修成績照会 [CampusSquare].pdf", "국립창원대학교 등록금 고지서.pdf"))
      and not any(l.startswith("국립창원대학교") for l in txtD.splitlines()))
check("진단: 중복 세트 파일은 '대표 (중복·동일 내용)' 소속, 세트 표 비고에 '중복(동일 내용)'",
      any(l.startswith("2024년 기출문제 유형 2회(문제).xlsm | ") and "상시기출2회 (중복·동일 내용)" in l for l in txtD.splitlines())
      and any(l.startswith("상시기출2회 | ") and "중복(동일 내용): 2024년 기출문제 유형 2회(문제).xlsm" in l for l in txtD.splitlines()))
check("진단: 슬롯 행에 세트·PDF·기대값 함께 표시",
      "슬롯 '2026 1회' → 세트 '2026 컴활2급실기_상시복원기출문제 1회' (전체 토큰 일치" in txtD
      and "· PDF [기공자] 2026 컴활2급실기_상시복원기출문제 1회.pdf · 기대값 2026_1회_기대값.json" in txtD
      and "슬롯 '컴활 2급 상시' → 미발견 (후보 없음" in txtD)
check("진단: 컴활2급상시_문제지.pdf·계산드릴_문제지.pdf는 미소속(문제지 행 유지)",
      any(l.startswith("컴활2급상시_문제지.pdf | ") and "| 미소속 |" in l for l in txtD.splitlines())
      and any(l.startswith("계산드릴_문제지.pdf | ") and "| 미소속 |" in l for l in txtD.splitlines()))
# 적응형 일정의 첫 응시 순서 — v3.0.0: 고정 우선순위 표 없이 **이름순**
pri = [(s["name"], lab) for s, lab in sj.prioritized_sets(setsD)]
check("prioritized_sets: 보유 세트를 이름순으로 (라벨은 'n/전체' 순번)",
      [n for n, _l in pri] == sorted(s["name"] for s in setsD)
      and [l for _n, l in pri] == [f"{i + 1}/{len(setsD)}" for i in range(len(setsD))],
      str(pri))
# 비병합 변형: 내용이 다르면 별개 세트 — '2024 상시 2회'는 2024년 기출 2회, 상시기출2회는 슬롯 없음 + 비슷한 이름 안내
D2, KD2 = make_diag_tree(same_content=False)
setsD2 = sj.scan_sets(D2, {}, key_dirs=[KD2])
T2 = table(setsD2)
check("비병합: 세트 7개, '2024 상시 2회' → 2024년 기출문제 유형 2회(식별: 2024+2회) + PDF·기대값",
      len(setsD2) == 7 and T2["2024 상시 2회"][:3] == ("2024년 기출문제 유형 2회", "2024상시2회_문제지.pdf", "2024상시2회_기대값.json"),
      str(T2["2024 상시 2회"]))
check("비병합: '상시기출2회'는 어느 슬롯에도 없고 '컴활 2급 상시'도 미발견(2급 없음)",
      "상시기출2회" not in [v[0] for v in T2.values()] and T2["컴활 2급 상시"][0] is None)
sim = sj.similar_named_sets(setsD2)
check("비병합: 이름이 비슷한 세트 안내(상시기출2회 ↔ 2024년 기출문제 유형 2회, 공통 2회·기출)",
      any({a["name"], b["name"]} == {"상시기출2회", "2024년 기출문제 유형 2회"} and "2회" in c for a, b, c in sim), str([(a["name"], b["name"], c) for a, b, c in sim]))
txtD2 = sj.scan_diagnosis_text(D2, config={}, sets=setsD2)
check("비병합 진단: '이름이 비슷한 세트' 섹션", "== 이름이 비슷한 세트" in txtD2
      and ("2024년 기출문제 유형 2회 ↔ 상시기출2회" in txtD2 or "상시기출2회 ↔ 2024년 기출문제 유형 2회" in txtD2))
check("비병합: 나머지 슬롯 배정은 병합 변형과 동일",
      all(T2[k][:3] == EXPECT[k][:3] for k in EXPECT if k != "2024 상시 2회"), str({k: T2[k][:3] for k in EXPECT}))

print("10. 중복 세트 병합 (v2.3.1)")
M1 = tempfile.mkdtemp()
touch_xlsx(os.path.join(M1, "깊은/폴더/2024년 상시2회 2급_문제.xlsm"), "같은문제")
touch_xlsx(os.path.join(M1, "깊은/폴더/2024년 상시2회 2급_정답.xlsm"))
touch_xlsx(os.path.join(M1, "상시2회_문제.xlsm"), "같은문제")
touch_xlsx(os.path.join(M1, "상시2회_정답.xlsm"))
pdf_deep = touch_pdf(os.path.join(M1, "깊은/폴더/2024년 상시2회 2급_문제지.pdf"))   # 깊은 쪽에만 같은 키 PDF
setsM = sj.scan_sets(M1, {})
check("문제 파일 내용이 같으면 한 세트", len(setsM) == 1)
check("대표: PDF가 붙은 쪽(깊은 폴더)이 얕은 폴더보다 우선, 중복 목록에 얕은 쪽",
      setsM[0]["name"] == "2024년 상시2회 2급" and setsM[0]["pdf"] == pdf_deep
      and [d["name"] for d in setsM[0]["중복"]] == ["상시2회"], setsM[0]["name"])
M2 = tempfile.mkdtemp()
touch_xlsx(os.path.join(M2, "깊은/폴더/2024년 상시2회 2급_문제.xlsm"), "같은문제")
touch_xlsx(os.path.join(M2, "깊은/폴더/2024년 상시2회 2급_정답.xlsm"))
touch_xlsx(os.path.join(M2, "상시2회_문제.xlsm"), "같은문제")
touch_xlsx(os.path.join(M2, "상시2회_정답.xlsm"))
touch_xlsx(os.path.join(M2, "다른내용_문제.xlsm"))      # 크기 같아도 내용 다르면 별개
touch_xlsx(os.path.join(M2, "다른내용_정답.xlsm"))
setsM2 = sj.scan_sets(M2, {})
repM2 = next(s for s in setsM2 if s.get("중복"))
check("PDF·기대값이 없으면 얕은 폴더가 대표, 내용 다른 세트는 별개",
      len(setsM2) == 2 and repM2["name"] == "상시2회" and [d["name"] for d in repM2["중복"]] == ["2024년 상시2회 2급"])
check("대표 토큰에 중복 세트의 연도·급수 포함 → '2024 상시 2회' 전체 토큰 일치",
      sj.match_slot(setsM2, "2024 상시 2회")[0] is repM2 and "전체 토큰" in sj.match_slot(setsM2, "2024 상시 2회")[1])
deep_norm = repM2["중복"][0]["norm"]
cfgM = {deep_norm: {"name": "2024년 상시2회 2급", "problem": repM2["중복"][0]["problem"],
                    "answer": repM2["중복"][0]["answer"], "key": None, "pdf": pdf_deep}}
setsM2b = sj.scan_sets(M2, cfgM)
repM2b = next(s for s in setsM2b if s.get("중복"))
check("세트설정에 중복 쪽 키로 저장된 항목(PDF)은 대표에 반영되고 별도 세트로 되살아나지 않음",
      len(setsM2b) == 2 and repM2b["pdf"] == pdf_deep)
check("직접 선택 매핑이 중복 쪽 파일을 가리켜도 대표는 다른 슬롯 후보에서 제외",
      sj.assign_slots(setsM2, ["2024 상시 2회", "24 2급 상시"],
                      {"24 2급 상시": {"problem": repM2["중복"][0]["problem"], "answer": repM2["중복"][0]["answer"]}}
                      )["2024 상시 2회"][0] is None)
check("merge_duplicate_sets: 읽을 수 없는 경로는 건너뜀(예외 없음)",
      len(sj.merge_duplicate_sets([{"name": "x", "norm": "x", "dir": "/없음", "problem": "/없음/a.xlsx", "answer": "/없음/b.xlsx"}])) == 1)

print("11. PDF·기대값 전역 유일 연결 (v2.3.1)")
L1 = make_tree(["k/2026 상시 1회_문제.xlsm", "k/2026 상시 1회_정답.xlsm",
                "k/2026 상시 2회_문제.xlsm", "k/2026 상시 2회_정답.xlsm"])
touch_pdf(os.path.join(L1, "k", "2026 상시 문제지.pdf"))          # 회차 없음 → 두 세트 동점
setsL1 = sj.scan_sets(L1, {})
check("한 PDF에 두 세트가 동점(둘 다 문제지 없음) → 연결 안 함 + 양쪽 pdf_tie 사유",
      all(s["pdf"] is None for s in setsL1)
      and all("2026 상시 문제지.pdf ↔ 세트" in (s.get("pdf_tie") or "") and "동점" in s["pdf_tie"] for s in setsL1),
      str([s.get("pdf_tie") for s in setsL1]))
txtL1 = sj.scan_diagnosis_text(L1, config={}, sets=setsL1)
check("진단: 동점 PDF는 '미소속 (동점: …)' + 세트 표 '(동점: …)'",
      any(l.startswith("2026 상시 문제지.pdf | ") and "미소속 (동점:" in l for l in txtL1.splitlines())
      and "- (동점: 2026 상시 문제지.pdf" in txtL1)
L2 = make_tree(["k/컴활 2급 상시 기출 복원_문제.xlsm", "k/컴활 2급 상시 기출 복원_정답.xlsm",
                "k/컴활 상시_문제.xlsm", "k/컴활 상시_정답.xlsm"])
own = touch_pdf(os.path.join(L2, "k", "컴활 2급 상시 기출 복원_문제지.pdf"))   # 같은 키
stray = touch_pdf(os.path.join(L2, "k", "컴활 2급 상시 기출 문제지.pdf"))       # 복원 세트에 4개, 컴활 상시에 2개
setsL2 = sj.scan_sets(L2, {})
sFull = next(s for s in setsL2 if "복원" in s["name"])
sLess = next(s for s in setsL2 if s["name"] == "컴활 상시")
check("문제지가 이미 있는 세트에 더 잘 맞는 PDF는 다른 세트에 붙이지 않음(사본으로 판단)",
      sFull["pdf"] == own and sLess["pdf"] is None, str(sLess["pdf"]))
L3 = make_tree(["k/2025 모의 1회_문제.xlsm", "k/2025 모의 1회_정답.xlsm"])
touch_pdf(os.path.join(L3, "k", "2025 1회 문제지 a.pdf"))
touch_pdf(os.path.join(L3, "k", "2025 1회 문제지 b.pdf"))
sL3 = sj.scan_sets(L3, {})[0]
check("한 세트에 같은 점수의 PDF 둘 → 연결 안 함 + 두 파일 동점 사유",
      sL3["pdf"] is None and "2025 1회 문제지 a.pdf" in sL3["pdf_tie"] and "2025 1회 문제지 b.pdf" in sL3["pdf_tie"], sL3.get("pdf_tie"))
L4 = make_tree(["k/2026 상시 1회_문제.xlsm", "k/2026 상시 1회_정답.xlsm"])
touch_pdf(os.path.join(L4, "k", "아이모2026_발표.pdf"))
sL4 = sj.scan_sets(L4, {})[0]
check("연도만 겹치고 '문제'가 없는 PDF는 후보 아님", sL4["pdf"] is None)
touch_pdf(os.path.join(L4, "k", "2026_문제지.pdf"))
sL4b = sj.scan_sets(L4, {})[0]
check("연도만 겹쳐도 '문제지'이고 세트 토큰에 포함되면 연결", sL4b["pdf"] and sL4b["pdf"].endswith("2026_문제지.pdf"))
L5 = make_tree(["k/2024년 기출문제 유형 1회_문제.xlsm", "k/2024년 기출문제 유형 1회_정답.xlsm"])
touch_json(os.path.join(L5, "k", "코코모의고사1회_기대값.json"))
sL5 = sj.scan_sets(L5, {})[0]
check("토큰 1개('1회')만 겹치고 파일 토큰(코코·모의)이 세트에 없으면 기대값 연결 안 함", sL5["key"] is None, str(sL5["key"]))
L6 = make_tree(["k/2024 상시 2회_문제.xlsm", "k/2024 상시 2회_정답.xlsm",
                "k/상시기출2회_문제.xlsm", "k/상시기출2회_정답.xlsm"])
touch_pdf(os.path.join(L6, "k", "2024상시2회_문제지.pdf"))
setsL6 = sj.scan_sets(L6, {})
sL6a = next(s for s in setsL6 if s["name"] == "2024 상시 2회")
sL6b = next(s for s in setsL6 if s["name"] == "상시기출2회")
check("같은 겹침 수(2)면 연도·회차 겹침이 많은 세트가 이김(2024 상시 2회 > 상시기출2회)",
      sL6a["pdf"] is not None and sL6b["pdf"] is None)
check("link_files_globally: 이미 연결된 파일과 같은 이름의 사본은 후보 제외, 빈 목록 안전",
      sj.link_files_globally([], [], "pdf") is None)

print("12. 자동 배포된 루트/모의고사/ 코코 세트 → '코코 1회·2회' 슬롯 (v2.3.1)")
D3, KD3 = make_diag_tree(same_content=True, with_coco=True)
setsD3 = sj.scan_sets(D3, {}, key_dirs=[KD3])
T3 = table(setsD3)
check("세트 8개(코코 2개 편입)", len(setsD3) == 8)
check("'코코 1회' → 코코모의고사1회 / 세트 폴더 PDF / 기대값 폴더 JSON",
      T3["코코 1회"] [:3] == ("코코모의고사1회", "코코모의고사1회_문제지.pdf", "코코모의고사1회_기대값.json")
      and next(s for s in setsD3 if s["name"] == "코코모의고사1회")["pdf"] == os.path.join(D3, "모의고사", "코코모의고사1회_문제지.pdf")
      and next(s for s in setsD3 if s["name"] == "코코모의고사1회")["key_origin"] == "기대값 폴더", str(T3["코코 1회"]))
check("'코코 2회' → 코코모의고사2회 / PDF / 기대값",
      T3["코코 2회"][:3] == ("코코모의고사2회", "코코모의고사2회_문제지.pdf", "코코모의고사2회_기대값.json"), str(T3["코코 2회"]))
check("코코 편입 후에도 나머지 7개 슬롯 배정 동일",
      all(T3[k][:3] == EXPECT[k][:3] for k in EXPECT if not k.startswith("코코")))
txtD3 = sj.scan_diagnosis_text(D3, config={}, sets=setsD3)
check("진단: 바탕 화면의 코코 PDF 사본은 '(같은 키 사본)' 소속 표시",
      any(l.startswith("코코모의고사1회_문제지.pdf | ") and "코코모의고사1회 (같은 키 사본)" in l and l.rstrip().endswith("| .")
          for l in txtD3.splitlines()), [l for l in txtD3.splitlines() if l.startswith("코코모의고사1회_문제지.pdf")])
check("슬롯 9개 전부 배정(미발견은 '컴활 2급 상시' 하나 — 사용자 폴더에 그 xlsx가 없음)",
      [k for k, v in T3.items() if v[0] is None] == ["컴활 2급 상시"])
# 일정은 세트를 고정하지 않으므로(v3.0.0) 세트 이름을 넣은 슬롯이 전역 배정을 쓰는지 확인
planD3 = dict(sj.plan_for_day(1), 세트=["2024 상시 1회"])
rsD3 = sj.resolve_day_sets(planD3, setsD3, [], None)
check("세트 이름을 지정한 슬롯도 전역 배정 결과 사용",
      rsD3[0][0] is not None and rsD3[0][0]["name"] == "2024년 기출문제 유형 1회" and "식별 토큰" in rsD3[0][1], str(rsD3))

print("13. 전역 유일 배정 규칙 (v2.3.1)")
G1 = make_tree(["컴활/2026 컴활2급실기_상시복원기출문제 1회_문제.xlsm", "컴활/2026 컴활2급실기_상시복원기출문제 1회_정답.xlsm"])
setsG1 = sj.scan_sets(G1, {})
A = sj.assign_slots(setsG1)
check("2026 세트 하나: '2026 1회'만 가져가고 '2024 상시 1회'·'컴활 2급 상시'는 후보 없음(충돌·연도제외)",
      A["2026 1회"][0] is not None and A["2024 상시 1회"][0] is None and A["컴활 2급 상시"][0] is None
      and "후보 없음" in A["2024 상시 1회"][1] and "후보 없음" in A["컴활 2급 상시"][1], str({k: v[1] for k, v in A.items()}))
check("all_slot_specs: 식별 표 9개를 모두 포함하고 중복 없음",
      set(sj.SLOT_IDENTITY) <= set(sj.all_slot_specs()) and len(sj.all_slot_specs()) == len(set(sj.all_slot_specs())))
G2 = make_tree(["k/코코모의고사1회_문제.xlsx", "k/코코모의고사1회_정답.xlsx"])
setsG2 = sj.scan_sets(G2, {})
A2 = sj.assign_slots(setsG2, ["코코 1회", "1회 코코"])
check("같은 세트를 두 슬롯이 같은 점수로 원하면(동점) 둘 다 미배정 + 사유",
      A2["코코 1회"][0] is None and A2["1회 코코"][0] is None and "동점" in A2["코코 1회"][1] and "'1회 코코'" in A2["코코 1회"][1],
      str((A2["코코 1회"][1], A2["1회 코코"][1])))
A2b = sj.assign_slots(setsG2, ["코코 1회", "코코 모의 1회"])
check("점수가 더 높은 슬롯(토큰 3개 겹침)이 이기고 진 슬롯 사유에 '배정됨' 안내",
      A2b["코코 모의 1회"][0] is not None and A2b["코코 1회"][0] is None and "슬롯 '코코 모의 1회'에 배정됨" in A2b["코코 1회"][1],
      A2b["코코 1회"][1])
sc = sj.slot_candidate_score
sG6 = {s["name"]: s for s in sets6}
check("slot_candidate_score: 등급 3(전체) > 2(식별) > 1(부분), 충돌·연도제외·핵심 누락은 None",
      sc("24 2급 상시", sG6["24년 2급 상시"])[0][0] == 3
      and sc("24 2급 상시", sG6["2024년 상시1회 2급"])[0][0] == 3
      and sc("2024 상시 1회", sG6["2024년 상시1회 2급"])[0] > sc("24 2급 상시", sG6["2024년 상시1회 2급"])[0]
      and sc("컴활 2급 상시", sG6["2024년 상시1회 2급"]) is None
      and sc("2026 1회", sG6["2024년 상시1회 2급"]) is None
      and sc("24 2급 상시", sG6["컴활 2급 상시"]) is None)
P = make_tree(["k/24년 2급 문제.xlsx", "k/24년 2급 정답.xlsm"])
sP = sj.scan_sets(P, {})[0]
check("보조 묶음(모의/실기/상시)이 없으면 '부분 일치'로 배정 + 사유",
      sc("24 2급 상시", sP)[0][0] == 1 and "부분 일치" in sj.match_slot([sP], "24 2급 상시")[1], sj.match_slot([sP], "24 2급 상시")[1])
mapG = {"컴활 2급 상시": {"problem": sG6["24년 2급 상시"]["problem"], "answer": sG6["24년 2급 상시"]["answer"]}}
A3 = sj.assign_slots(sets6, SLOTS9, mapG)
check("직접 선택 매핑 최우선: 매핑된 세트는 다른 슬롯 후보에서 제외되고 사유에 파일명",
      A3["컴활 2급 상시"][0]["name"] == "24년 2급 상시" and "직접 선택 저장됨 (24년 2급 상시 문제.xlsx)" in A3["컴활 2급 상시"][1]
      and A3["24 2급 상시"][0] is None and "슬롯 '2024 상시 1회'에 배정됨" in A3["24 2급 상시"][1]
      and A3["2024 상시 1회"][0]["name"] == "2024년 상시1회 2급", str({k: (v[0] and v[0]["name"], v[1]) for k, v in A3.items()}))
check("match_slot ↔ assign_slots 일관성(모든 슬롯)",
      all(sj.match_slot(setsD3, k)[0] is sj.assign_slots(setsD3, SLOTS9)[k][0] for k in SLOTS9))
check("match_slot(mapping)도 매핑 반영", sj.match_slot(sets6, "컴활 2급 상시", mapG)[0]["name"] == "24년 2급 상시")

print()
print(f"세트 인식 테스트 {N}건 전부 통과")
