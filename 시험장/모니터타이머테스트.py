# -*- coding: utf-8 -*-
"""시험장.py v3.2.0 — 우측 상단 타이머 · 다중 모니터 배치 헤드리스 테스트.

화면(display)도 tkinter 도 Windows API 도 없는 환경에서 돌아갑니다. GUI 를
띄우지 않고 검증할 수 있도록 기능을 두 층으로 나눠 두었고, 여기서는 그
**순수 계산 층**과 **방어선**을 확인합니다.

  1. 모니터 목록 정규화 — 0/1/2/3대, 정렬, 주 모니터 기본값, 깨진 항목 버리기
  2. 사람이 읽는 이름표 ('모니터 2 (2560×1440, 오른쪽)')
  3. 설정 UI 노출 기준 (2대 이상일 때만) · 기본 선택
  4. 타이머 위치 계산 — 우측 상단 기준, 저장된 자리 복원, 화면 밖이면 되돌리기
  5. 설정 저장·읽기 (세트설정.json `_설정`) 왕복
  6. 창 배치용 좌표 계산 · 새로 뜬 창 고르기
  7. 플랫폼이 Windows 가 아니면 안전하게 비활성화
  8. Windows API 가 실패해도 예외가 새어 나오지 않음 (가짜 ctypes 로 강제)
  9. GUI 층 구조 — 오버레이 정리 경로·설정 창 칸이 코드에 있는지 (ast)
 10. 기존 시간 경고 규칙(색·ALERTS)을 오버레이가 다시 정의하지 않음

GUI 가 실제로 어떻게 보이고 Excel 창이 옮겨지는지는 이 테스트로 알 수 없습니다
— Windows PC 에서 직접 확인해야 합니다.
"""
import ast
import importlib.util
import inspect
import io
import json
import os
import sys
import tempfile
import types

BASE = os.path.dirname(os.path.abspath(__file__))
SOURCE = os.path.join(BASE, "시험장.py")
spec = importlib.util.spec_from_file_location("sj", SOURCE)
sj = importlib.util.module_from_spec(spec)
spec.loader.exec_module(sj)
N = 0
TMP = tempfile.mkdtemp(prefix="coco_monitor_")
sj.SET_CONFIG_PATH = os.path.join(TMP, "세트설정.json")
sj.STARTUP_LOG_PATH = os.path.join(TMP, "시작로그.txt")
sj.ERROR_LOG_PATH = os.path.join(TMP, "오류.log")


def check(desc, cond, extra=""):
    global N
    N += 1
    print(f"  [{N:03d}] {desc}: {'OK' if cond else 'FAIL'}"
          + (f" ({extra})" if extra else ""))
    if not cond:
        raise AssertionError(f"{desc} {extra}")


def cfg():
    try:
        with io.open(sj.SET_CONFIG_PATH, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def reset_cfg(data=None):
    with io.open(sj.SET_CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(data if data is not None else {}, f, ensure_ascii=False)


def raw(left, top, w, h, primary=False, taskbar=0):
    """가짜 모니터 한 대 (작업 표시줄만큼 작업영역을 줄인다)."""
    return {"rect": (left, top, left + w, top + h),
            "work": (left, top, left + w, top + h - taskbar),
            "primary": primary}


reset_cfg()

check("버전 3.2.0 이상", tuple(int(x) for x in sj.__version__.split(".")) >= (3, 2, 0),
      sj.__version__)

# ---------------------------------------------------------------------------
print("1. 모니터 목록 정규화 (0/1/2/3대)")
# ---------------------------------------------------------------------------
check("모니터 0대 → 빈 목록", sj.normalize_monitors([]) == [])
check("None 이어도 빈 목록", sj.normalize_monitors(None) == [])

one = sj.normalize_monitors([raw(0, 0, 1920, 1080, primary=True, taskbar=40)])
check("모니터 1대: 길이 1 · index 0 · 주 모니터", len(one) == 1
      and one[0]["index"] == 0 and one[0]["primary"] is True)
check("1대: 폭·높이는 rcMonitor 기준", (one[0]["width"], one[0]["height"]) == (1920, 1080))
check("1대: 작업영역은 rcWork (작업 표시줄 제외)", one[0]["work"] == (0, 0, 1920, 1040))
check("1대: 위치 낱말 없음 (고를 것이 없음)", one[0]["position"] == "")

# 주 모니터가 오른쪽에 있고, 보조가 왼쪽 음수 좌표에 있는 흔한 배치
two = sj.normalize_monitors([raw(0, 0, 1920, 1080, primary=True, taskbar=40),
                             raw(-2560, 0, 2560, 1440)])
check("모니터 2대: 왼→오 정렬 (음수 좌표 보조가 먼저)",
      [m["rect"][0] for m in two] == [-2560, 0])
check("2대: index 는 정렬 뒤 0,1", [m["index"] for m in two] == [0, 1])
check("2대: 주 모니터 표식 유지 (오른쪽이 주 모니터)",
      [m["primary"] for m in two] == [False, True])
check("2대: 위치 낱말 왼쪽/오른쪽",
      [m["position"] for m in two] == ["왼쪽", "오른쪽"])

three = sj.normalize_monitors([raw(1920, 0, 1920, 1080),
                               raw(0, 0, 1920, 1080, primary=True),
                               raw(3840, 0, 1280, 1024)])
check("모니터 3대: 왼쪽/가운데/오른쪽",
      [m["position"] for m in three] == ["왼쪽", "가운데", "오른쪽"])
check("3대: 주 모니터는 가장 왼쪽 것",
      [m["primary"] for m in three] == [True, False, False])

four = sj.normalize_monitors([raw(i * 1920, 0, 1920, 1080) for i in range(4)])
check("모니터 4대: '왼쪽에서 N번째' 로 셈",
      [m["position"] for m in four]
      == ["왼쪽에서 1번째", "왼쪽에서 2번째", "왼쪽에서 3번째", "왼쪽에서 4번째"])

stacked = sj.normalize_monitors([raw(0, 1080, 1920, 1080),
                                 raw(0, 0, 1920, 1080, primary=True)])
check("세로로 쌓은 2대 → 위/아래",
      [m["position"] for m in stacked] == ["위", "아래"]
      and [m["rect"][1] for m in stacked] == [0, 1080])

check("주 모니터 표식이 하나도 없으면 첫 번째를 주 모니터로",
      sj.normalize_monitors([raw(0, 0, 1920, 1080),
                             raw(1920, 0, 1920, 1080)])[0]["primary"] is True)

junk = sj.normalize_monitors([
    {"rect": (0, 0, 1920, 1080), "primary": True},       # work 없음 → rect 로
    {"rect": (10, 10, 5, 5)},                            # 폭/높이 0 이하 → 버림
    {"rect": ("a", "b", "c", "d")},                      # 숫자 아님 → 버림
    {"work": (0, 0, 100, 100)},                          # rect 없음 → 버림
    "모니터",                                             # dict 아님 → 버림
    None,
])
check("깨진 항목은 버리고 성한 것만 남김", len(junk) == 1)
check("work 가 없으면 rect 로 채움", junk[0]["work"] == junk[0]["rect"])

check("rect_size", sj.rect_size((-100, 50, 1820, 1130)) == (1920, 1080))

# ---------------------------------------------------------------------------
print("2. 사람이 읽는 이름표")
# ---------------------------------------------------------------------------
check("이름표 번호는 1부터 · 주 모니터 · 해상도 · 위치",
      one[0]["label"] == "모니터 1 (주 모니터, 1920×1080)", one[0]["label"])
check("2대 이름표 (보조)", two[0]["label"] == "모니터 1 (2560×1440, 왼쪽)",
      two[0]["label"])
check("2대 이름표 (주 모니터)",
      two[1]["label"] == "모니터 2 (주 모니터, 1920×1080, 오른쪽)", two[1]["label"])
check("이름표는 모두 서로 다름 (콤보 상자에서 구분 가능)",
      len({m["label"] for m in three}) == 3)

# ---------------------------------------------------------------------------
print("3. 설정 UI 노출 기준 · 기본 선택")
# ---------------------------------------------------------------------------
check("모니터 0대 → UI 숨김", sj.monitor_ui_visible([]) is False)
check("모니터 1대 → UI 숨김 (고를 것이 없음)", sj.monitor_ui_visible(one) is False)
check("모니터 2대 → UI 노출", sj.monitor_ui_visible(two) is True)
check("모니터 3대 → UI 노출", sj.monitor_ui_visible(three) is True)
check("None → 숨김", sj.monitor_ui_visible(None) is False)

check("기본 선택(모니터 없음): 전부 0",
      sj.default_monitor_choices([]) == {"pdf": 0, "excel": 0, "timer": 0})
check("기본 선택(1대): 전부 같은 화면",
      sj.default_monitor_choices(one) == {"pdf": 0, "excel": 0, "timer": 0})
d2 = sj.default_monitor_choices(two)
check("기본 선택(2대): Excel=주 모니터(1) · 문제지=남는 화면(0)",
      d2["excel"] == 1 and d2["pdf"] == 0, str(d2))
check("기본 선택: 타이머는 Excel 과 같은 화면 (작업하는 쪽에 남은 시간)",
      d2["timer"] == d2["excel"])
d3 = sj.default_monitor_choices(three)
check("기본 선택(3대): Excel=주 모니터(0) · 문제지=다음 화면(1)",
      d3 == {"pdf": 1, "excel": 0, "timer": 0}, str(d3))

check("pick_monitor: 정상 번호", sj.pick_monitor(two, 0) is two[0])
check("pick_monitor: 범위 밖 → 주 모니터", sj.pick_monitor(two, 9) is two[1])
check("pick_monitor: 음수 → 주 모니터", sj.pick_monitor(two, -1) is two[1])
check("pick_monitor: None → 주 모니터", sj.pick_monitor(two, None) is two[1])
check("pick_monitor: 숫자 아님 → 주 모니터", sj.pick_monitor(two, "오른쪽") is two[1])
check("pick_monitor: 목록이 비면 None (배치하지 않음)",
      sj.pick_monitor([], 0) is None)
check("primary_monitor: 표식 없으면 첫 번째",
      sj.primary_monitor([{"rect": (0, 0, 1, 1)}])["rect"] == (0, 0, 1, 1))
check("primary_monitor: 빈 목록이면 None", sj.primary_monitor([]) is None)

# ---------------------------------------------------------------------------
print("4. 타이머 위치 계산 (우측 상단 · 저장된 자리 복원)")
# ---------------------------------------------------------------------------
SIZE = (200, 100)
check("우측 상단: x = 오른쪽 - 폭 - 여백, y = 위 + 여백",
      sj.overlay_top_right((0, 0, 1920, 1040), SIZE, 16) == (1704, 16))
check("우측 상단: 음수 좌표 모니터에서도 그 모니터 안",
      sj.overlay_top_right((-2560, 0, 0, 1400), SIZE, 16) == (-216, 16))
check("우측 상단: 창보다 좁은 모니터면 왼쪽 끝을 넘지 않음",
      sj.overlay_top_right((0, 0, 100, 100), SIZE, 16) == (0, 16))

check("저장된 자리 없음 → 고른 모니터(0=왼쪽)의 우측 상단",
      sj.overlay_position(two, 0, SIZE, None, 16) == (-216, 16),
      str(sj.overlay_position(two, 0, SIZE, None, 16)))
check("저장된 자리 없음 → 고른 모니터(1=주 모니터) 우측 상단 (작업영역 기준)",
      sj.overlay_position(two, 1, SIZE, None, 16) == (1704, 16))
check("저장된 자리가 화면 안이면 그 자리를 복원",
      sj.overlay_position(two, 1, SIZE, (300, 400), 16) == (300, 400))
check("저장된 자리가 어느 모니터에도 안 걸치면 우측 상단으로 되돌림",
      sj.overlay_position(two, 1, SIZE, (99999, 99999), 16) == (1704, 16))
check("저장된 자리가 깨졌으면(문자열) 우측 상단",
      sj.overlay_position(two, 1, SIZE, ("x", "y"), 16) == (1704, 16))
check("모니터를 하나도 모르면(비 Windows) 저장된 자리는 그대로 존중",
      sj.overlay_position([], 0, SIZE, (120, 34), 16) == (120, 34))
check("모니터도 저장된 자리도 없으면 None (tk 화면 크기로 대신 계산)",
      sj.overlay_position([], 0, SIZE, None, 16) is None)

check("tk 화면 크기 대체 계산: 우측 상단",
      sj.overlay_position_for_screen(1366, 768, SIZE, None, 16) == (1150, 16))
check("tk 화면 크기 대체 계산: 저장된 자리가 화면 밖이면 우측 상단",
      sj.overlay_position_for_screen(1366, 768, SIZE, (5000, 5000), 16)
      == (1150, 16))

check("걸침 판정: 모니터 안", sj.point_on_any_monitor(100, 100, two, SIZE) is True)
check("걸침 판정: 모니터 밖", sj.point_on_any_monitor(9000, 9000, two, SIZE) is False)
check("걸침 판정: 모서리에 40px 만 걸쳐도 보이는 것으로",
      sj.point_on_any_monitor(1920 - 40, 0, two, SIZE) is True)
check("걸침 판정: 39px 만 걸치면 밖으로 봄",
      sj.point_on_any_monitor(1920 - 39, 0, two, SIZE) is False)
check("걸침 판정: 모니터 목록이 비면 False",
      sj.point_on_any_monitor(0, 0, [], SIZE) is False)

check("clamp: 화면 안이면 그대로", sj.clamp_overlay_position(300, 300, two, SIZE)
      == (300, 300))
clamped = sj.clamp_overlay_position(9000, 9000, two, SIZE)
check("clamp: 화면 밖이면 가장 가까운 모니터 작업영역 안으로",
      sj.point_on_any_monitor(clamped[0], clamped[1], two, SIZE) is True,
      str(clamped))
check("clamp: 오른쪽 아래로 넘긴 경우 주 모니터 작업영역 끝",
      clamped == (1720, 940), str(clamped))
check("clamp: 모니터를 모르면 손댄 자리 그대로",
      sj.clamp_overlay_position(7, 9, [], SIZE) == (7, 9))

# ---------------------------------------------------------------------------
print("5. 설정 저장·읽기 (세트설정.json `_설정`)")
# ---------------------------------------------------------------------------
reset_cfg()
check("타이머 오버레이 기본값은 켜짐", sj.timer_overlay_enabled() is True)
check("끄면 False 로 저장", sj.set_timer_overlay_enabled(False) is False
      and sj.timer_overlay_enabled() is False)
check("저장 위치는 _설정.타이머오버레이",
      cfg()["_설정"][sj.TIMER_OVERLAY_SETTING] is False)
check("다시 켜면 True", sj.set_timer_overlay_enabled(True) is True
      and sj.timer_overlay_enabled() is True)
check("cfg 를 직접 줘도 읽힘 (파일 다시 안 읽음)",
      sj.timer_overlay_enabled(cfg={"_설정": {sj.TIMER_OVERLAY_SETTING: False}})
      is False)
check("값이 None 이면 켜짐으로 (설정이 비어 있던 옛 파일)",
      sj.timer_overlay_enabled(cfg={"_설정": {sj.TIMER_OVERLAY_SETTING: None}})
      is True)

check("타이머 자리: 기본은 없음", sj.saved_overlay_pos() is None)
check("자리 저장 왕복", sj.save_overlay_pos(1704, 16) == (1704, 16)
      and sj.saved_overlay_pos() == (1704, 16))
check("자리는 [x, y] 로 저장",
      cfg()["_설정"][sj.TIMER_OVERLAY_POS_SETTING] == [1704, 16])
check("옛 {x:..,y:..} 모양도 읽힘",
      sj.saved_overlay_pos(cfg={"_설정": {sj.TIMER_OVERLAY_POS_SETTING:
                                          {"x": 5, "y": 6}}}) == (5, 6))
check("깨진 자리는 None (죽지 않음)",
      sj.saved_overlay_pos(cfg={"_설정": {sj.TIMER_OVERLAY_POS_SETTING: "x"}})
      is None)
check("자리 지우기 ([타이머 위치 초기화])",
      sj.save_overlay_pos(None, None) is None and sj.saved_overlay_pos() is None)
check("숫자가 아니면 저장하지 않음", sj.save_overlay_pos("a", "b") is None)

reset_cfg()
check("모니터 배치: 고르기 전에는 전부 None",
      sj.monitor_assignment() == {"pdf": None, "excel": None, "timer": None})
saved = sj.save_monitor_assignment(pdf=0, excel=1, timer=1)
check("모니터 배치 저장 왕복", saved == {"pdf": 0, "excel": 1, "timer": 1}
      and sj.monitor_assignment() == saved)
check("저장 키는 _설정.문제지모니터/엑셀모니터/타이머모니터",
      cfg()["_설정"][sj.MONITOR_PDF_SETTING] == 0
      and cfg()["_설정"][sj.MONITOR_EXCEL_SETTING] == 1
      and cfg()["_설정"][sj.MONITOR_TIMER_SETTING] == 1)
check("값을 주지 않은 항목은 그대로 둠",
      sj.save_monitor_assignment(timer=0) == {"pdf": 0, "excel": 1, "timer": 0})
check("숫자가 아닌 값은 무시",
      sj.save_monitor_assignment(pdf="왼쪽")["pdf"] == 0)
check("설정이 깨져 있으면(문자열) None 으로 읽음",
      sj.monitor_assignment(cfg={"_설정": {sj.MONITOR_PDF_SETTING: "x"}})["pdf"]
      is None)

sj.save_monitor_assignment(pdf=0, excel=1, timer=1)
check("배치 해석: 저장값이 범위 안이면 저장값",
      sj.resolve_monitor_assignment(two) == {"pdf": 0, "excel": 1, "timer": 1})
check("배치 해석: 모니터를 뺐으면(범위 밖) 조용히 기본값으로",
      sj.resolve_monitor_assignment(one) == {"pdf": 0, "excel": 0, "timer": 0})
check("배치 해석: 모니터를 하나도 모르면 기본값 0",
      sj.resolve_monitor_assignment([]) == {"pdf": 0, "excel": 0, "timer": 0})
sj.save_monitor_assignment(pdf=2, excel=2, timer=2)
check("배치 해석: 3대에서 2번 화면을 고른 그대로",
      sj.resolve_monitor_assignment(three) == {"pdf": 2, "excel": 2, "timer": 2})
check("배치 해석: 같은 설정으로 2대만 남으면 기본값",
      sj.resolve_monitor_assignment(two) == d2)
check("기본값으로 되돌리기", sj.clear_monitor_assignment()
      == {"pdf": None, "excel": None, "timer": None}
      and sj.MONITOR_PDF_SETTING not in cfg().get("_설정", {}))
check("되돌린 뒤 해석하면 기본 배치", sj.resolve_monitor_assignment(two) == d2)

# 설정이 통째로 깨져도 죽지 않는다
reset_cfg({"_설정": "망가짐"})
check("_설정 이 dict 가 아니어도 기본값으로 읽힘",
      sj.timer_overlay_enabled() is True and sj.saved_overlay_pos() is None
      and sj.monitor_assignment() == {"pdf": None, "excel": None, "timer": None})
with io.open(sj.SET_CONFIG_PATH, "w", encoding="utf-8") as f:
    f.write("{깨진 json")
check("세트설정.json 이 깨져도 기본값",
      sj.timer_overlay_enabled() is True
      and sj.resolve_monitor_assignment(two) == d2)
reset_cfg()

# ---------------------------------------------------------------------------
print("6. 창 배치용 좌표 계산 · 새로 뜬 창 고르기")
# ---------------------------------------------------------------------------
check("창 배치 좌표 = 작업영역 (작업 표시줄 제외)",
      sj.window_rect_for_monitor(two[1]) == (0, 0, 1920, 1040))
check("창 배치 좌표: 보조 모니터 (음수 좌표)",
      sj.window_rect_for_monitor(two[0]) == (-2560, 0, 2560, 1440))
check("창 배치 좌표: 여백을 주면 안쪽으로",
      sj.window_rect_for_monitor(two[1], margin=20) == (20, 20, 1880, 1000))
check("창 배치 좌표: 모니터가 없으면 None",
      sj.window_rect_for_monitor(None) is None)
check("창 배치 좌표: rect 도 work 도 없으면 None",
      sj.window_rect_for_monitor({}) is None)

check("새로 뜬 창만 고름 (순서 유지)",
      sj.pick_new_windows([1, 2, 3], [3, 2, 7, 1, 9]) == [7, 9])
check("새 창이 없으면 빈 목록", sj.pick_new_windows([1, 2], [2, 1]) == [])
check("before 가 비면 전부 새 창", sj.pick_new_windows([], [4, 5]) == [4, 5])
check("None 이어도 죽지 않음", sj.pick_new_windows(None, None) == [])

# ---------------------------------------------------------------------------
print("7. 플랫폼이 Windows 가 아니면 안전하게 비활성화")
# ---------------------------------------------------------------------------
check("여기는 Windows 가 아님 (테스트 전제)", sys.platform != "win32", sys.platform)
check("monitors_supported() False", sj.monitors_supported() is False)
check("모니터 열거는 빈 목록", sj.enum_monitors_raw() == [])
check("모니터 목록도 빈 목록 → 설정 UI 도 안 뜸",
      sj.list_monitors() == [] and sj.monitor_ui_visible(sj.list_monitors()) is False)
check("창 열거는 빈 목록", sj.top_level_windows() == [])
check("Excel 창 찾기는 빈 목록", sj.excel_windows() == []
      and sj.excel_windows(4242) == []
      and sj.excel_windows(4242, strict=True) == [])
check("창 옮기기는 False (아무 일도 안 함)",
      sj.move_window_to_monitor(12345, two[1]) is False)
check("Excel 배치는 (0, 사유)", sj.place_excel_on_monitor(two[1])[0] == 0)
check("문제지 배치는 (0, 'Windows 아님')",
      sj.place_new_window_on_monitor(two[1], []) == (0, "Windows 아님"))
check("모니터를 안 주면 (0, '모니터 정보 없음')",
      sj.place_excel_on_monitor(None) == (0, "모니터 정보 없음")
      and sj.place_new_window_on_monitor(None, []) == (0, "모니터 정보 없음"))
check("Excel 창 클래스 상수는 XLMAIN", sj.EXCEL_WINDOW_CLASS == "XLMAIN")
_place_src = inspect.getsource(sj.place_excel_on_monitor)
check("Excel 배치는 우리가 띄운 프로세스를 먼저 대조한다",
      "strict=True" in _place_src)
check("프로세스를 모르면 맨 앞 창 하나만 — 남이 열어 둔 통합 문서는 그대로",
      "[:1]" in _place_src)

# ---------------------------------------------------------------------------
print("8. Windows API 가 실패해도 예외가 새어 나오지 않음")
# ---------------------------------------------------------------------------


class _Boom:
    """무엇을 건드려도 터지는 가짜 user32 (API 실패 재현)."""

    def __getattr__(self, name):
        raise OSError(f"가짜 실패: {name}")


fake_ctypes = types.ModuleType("ctypes")
fake_ctypes.windll = _Boom()
real_ctypes = sys.modules.get("ctypes")
real_platform = sys.platform
errors_before = os.path.getsize(sj.ERROR_LOG_PATH) \
    if os.path.isfile(sj.ERROR_LOG_PATH) else 0
sys.modules["ctypes"] = fake_ctypes
sys.platform = "win32"
try:
    check("가짜 Windows 에서는 지원한다고 봄 (여기서부터 API 가 터진다)",
          sj.monitors_supported() is True)
    check("모니터 열거 실패 → 빈 목록", sj.enum_monitors_raw() == [])
    check("모니터 목록 실패 → 빈 목록", sj.list_monitors() == [])
    check("창 열거 실패 → 빈 목록", sj.top_level_windows() == [])
    check("Excel 창 찾기 실패 → 빈 목록", sj.excel_windows(7) == []
          and sj.excel_windows(7, strict=True) == [])
    check("창 옮기기 실패 → False", sj.move_window_to_monitor(1, two[1]) is False)
    check("Excel 배치 실패 → (0, 사유)",
          sj.place_excel_on_monitor(two[1], 7)[0] == 0)
    check("문제지 배치 실패 → (0, 사유)",
          sj.place_new_window_on_monitor(two[1], [1, 2])[0] == 0)
    check("설정 UI 는 모니터를 못 읽으면 안 뜸",
          sj.monitor_ui_visible(sj.list_monitors()) is False)
    check("배치 해석은 그래도 기본값을 돌려줌",
          sj.resolve_monitor_assignment(sj.list_monitors())
          == {"pdf": 0, "excel": 0, "timer": 0})
finally:
    sys.platform = real_platform
    if real_ctypes is not None:
        sys.modules["ctypes"] = real_ctypes
    else:                                          # pragma: no cover
        sys.modules.pop("ctypes", None)
check("실패는 삼키지 않고 오류 로그에 남긴다 (조용히 죽지는 않음)",
      os.path.isfile(sj.ERROR_LOG_PATH)
      and os.path.getsize(sj.ERROR_LOG_PATH) > errors_before)
check("가짜 ctypes 를 되돌린 뒤 정상 동작", sj.monitors_supported() is False
      and sj.list_monitors() == [])

# 이 환경에는 tkinter 가 없을 수 있다 — 그래도 import 단계에서 죽지 않았다
check("tkinter 가 없어도 모듈이 그대로 import 됨 (HAS_TK 로만 갈라짐)",
      isinstance(sj.HAS_TK, bool))

# ---------------------------------------------------------------------------
print("9. GUI 층 구조 (ast — 화면 없이 확인할 수 있는 만큼)")
# ---------------------------------------------------------------------------
tree = ast.parse(io.open(SOURCE, encoding="utf-8").read())
classes = {}


def walk(node):
    for child in ast.iter_child_nodes(node):
        if isinstance(child, ast.ClassDef):
            classes[child.name] = child
        walk(child)


walk(tree)
check("TimerOverlay 클래스가 있다", "TimerOverlay" in classes)
over = classes["TimerOverlay"]
over_methods = {n.name for n in over.body if isinstance(n, ast.FunctionDef)}
for name in ("place_self", "_grab", "_drag_to", "_drop", "update_time",
             "close"):
    check(f"TimerOverlay.{name}", name in over_methods)
check("TimerOverlay.close_all (남은 창 일괄 정리)",
      any(isinstance(n, ast.FunctionDef) and n.name == "close_all"
          for n in over.body))
over_src = ast.unparse(over)
check("테두리 없는 항상 위 창 (overrideredirect + topmost)",
      "overrideredirect" in over_src and "-topmost" in over_src)
check("끌어서 옮기기 바인딩 (<B1-Motion>)", "<B1-Motion>" in over_src)
check("옮긴 자리를 설정에 저장 (save_overlay_pos)",
      "save_overlay_pos" in over_src)
check("자리 계산은 순수 함수에 맡김 (overlay_position)",
      "overlay_position" in over_src)

timer = classes["TimerWindow"]
timer_methods = {n.name for n in timer.body if isinstance(n, ast.FunctionDef)}
for name in ("_open_overlay", "_sync_overlay", "close_overlay", "destroy"):
    check(f"TimerWindow.{name}", name in timer_methods)
timer_src = ast.unparse(timer)
check("타이머 창이 닫힐 때 오버레이도 닫힌다 (destroy → close_overlay)",
      "close_overlay" in ast.unparse(
          next(n for n in timer.body
               if isinstance(n, ast.FunctionDef) and n.name == "destroy")))
check("채점이 끝나도 오버레이를 닫는다",
      "close_overlay" in ast.unparse(
          next(n for n in timer.body
               if isinstance(n, ast.FunctionDef) and n.name == "_grading_done")))
check("오버레이는 설정으로 끌 수 있다 (timer_overlay_enabled)",
      "timer_overlay_enabled" in timer_src)
check("1초마다 오버레이에 남은 시간을 옮긴다",
      "_sync_overlay" in ast.unparse(
          next(n for n in timer.body
               if isinstance(n, ast.FunctionDef) and n.name == "_tick")))

app = classes["ExamApp"]
app_methods = {n.name for n in app.body if isinstance(n, ast.FunctionDef)}
for name in ("_plan_monitor_layout", "_place_excel_window",
             "_place_pdf_window"):
    check(f"ExamApp.{name}", name in app_methods)
check("시험이 끝나면 남은 오버레이를 쓸어 담는다 (유령 창 방지)",
      "TimerOverlay.close_all" in ast.unparse(
          next(n for n in app.body
               if isinstance(n, ast.FunctionDef) and n.name == "exam_closed")))
pdf_src = ast.unparse(next(n for n in app.body if isinstance(n, ast.FunctionDef)
                           and n.name == "_place_pdf_window"))
check("문제지 창을 못 잡으면 한 줄 안내만 남기고 넘어간다",
      "show_toast" in pdf_src and "startup_log" in pdf_src)

settings = classes["SettingsDialog"]
set_methods = {n.name for n in settings.body if isinstance(n, ast.FunctionDef)}
for name in ("_build_monitor_section", "_monitor_choice", "apply_monitors",
             "reset_monitors", "_toggle_overlay", "reset_overlay_pos"):
    check(f"SettingsDialog.{name}", name in set_methods)
init_src = ast.unparse(next(n for n in settings.body
                            if isinstance(n, ast.FunctionDef)
                            and n.name == "__init__"))
check("모니터 칸은 2대 이상일 때만 만든다 (monitor_ui_visible 로 가름)",
      "monitor_ui_visible" in init_src and "_build_monitor_section" in init_src)
rows_node = next(n for n in settings.body if isinstance(n, ast.Assign)
                 and getattr(n.targets[0], "id", "") == "MONITOR_ROWS")
row_keys = [t.elts[0].value for t in rows_node.value.elts]
row_texts = [t.elts[1].value for t in rows_node.value.elts]
check("설정 창이 고를 항목은 문제지·Excel·타이머 셋",
      row_keys == ["pdf", "excel", "timer"], str(row_keys))
check("항목 문구에 무엇을 띄우는지 적혀 있다",
      "문제지" in row_texts[0] and "Excel" in row_texts[1]
      and "타이머" in row_texts[2], str(row_texts))

# ---------------------------------------------------------------------------
print("10. 기존 시간 경고 규칙을 오버레이가 다시 정의하지 않음")
# ---------------------------------------------------------------------------
color_src = ast.unparse(next(n for n in timer.body
                             if isinstance(n, ast.FunctionDef)
                             and n.name == "_color"))
check("타이머 창의 경고 기준은 그대로 (5분 빨강 · 10분 호박)",
      "300" in color_src and "600" in color_src
      and "#FF8B80" in color_src and "#F2B24C" in color_src)
alerts = next(n for n in timer.body if isinstance(n, ast.Assign)
              and getattr(n.targets[0], "id", "") == "ALERTS")
check("소리 경고 기준도 그대로 (10분·5분·1분)",
      [e.value for e in alerts.value.elts] == [600, 300, 60],
      ast.unparse(alerts.value))
check("오버레이는 시간 경고 기준을 스스로 정하지 않는다 (중복 정의 금지)",
      "300" not in over_src and "600" not in over_src)
check("오버레이 색은 타이머 창이 넘겨 준다 (_color 를 그대로 씀)",
      "_color()" in ast.unparse(
          next(n for n in timer.body if isinstance(n, ast.FunctionDef)
               and n.name == "_sync_overlay")))

print()
print(f"모니터·타이머 테스트 {N}건 전부 통과")
