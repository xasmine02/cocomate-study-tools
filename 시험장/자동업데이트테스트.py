# -*- coding: utf-8 -*-
"""시험장.py v2.3.1 자동 업데이트 헤드리스 테스트 (tkinter 불필요).

로컬 http.server 스레드로 version.json과 파일을 서빙해 실제 다운로드 경로를
검증합니다.
  1. version.json 조회·버전 비교·오프라인 처리
  2. 정상 흐름: files(프로그램)+data_files(기대값) 다운로드 → sha256·문법·JSON
     검증 → .bak 백업 → 교체 → 안내 기록 → 재시작 안내 소비
  3. 검증 실패 롤백: sha256 불일치 / JSON 깨짐 / py 문법 오류 / 파일 버전 불일치
     → 기존 파일 그대로, .new 잔여 없음
  4. 같은 버전에서 기대값만 동기화(data_files: 없으면 받고, 해시 다르면 갱신,
     원격에서 사라진 파일은 지우지 않음)
  5. UpdateCoordinator: 시험 중 연기 → 종료 후 적용, 토글 off면 서버 미접속,
     같은 실행 중복 확인 방지, 적용 후 버전 미상승(재시작 반복) 방지
  6. 2.2.3 호환: 구버전 apply_update 로직으로 새 version.json을 적용해도
     files만 처리하고 data_files/sha256은 무시(실패 없음)
  7. 코코채점.pyw v1.7.0: 같은 규칙의 기대값 자동 연결(detect_key)과 묻지 않는
     자동 업데이트(run_auto_update: applied / data / latest / disabled / offline)
  8. (v2.3.1) set_files: 코코 모의고사 세트 파일(xlsx·pdf)을 sha256만으로 검증해
     <루트>/모의고사/에 내려받음(새 버전 스테이징 + 같은 버전 동기화), sha256
     없음/불일치/형식 오류는 실패, 2.3.0의 data_files 검증(UTF-8·json)에 바이너리를
     넣으면 실패하므로 set_files 분리가 필요한 이유, 2.3.0 이하·코코채점은 set_files 무시
"""
import hashlib
import http.server
import importlib.util
import json
import os
import shutil
import sys
import tempfile
import threading
import urllib.parse

BASE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("sj", os.path.join(BASE, "시험장.py"))
sj = importlib.util.module_from_spec(spec)
spec.loader.exec_module(sj)
N = 0
TMP = tempfile.mkdtemp(prefix="coco_upd_")
sj.STARTUP_LOG_PATH = os.path.join(TMP, "시작로그.txt")
sj.ERROR_LOG_PATH = os.path.join(TMP, "오류.log")


def check(desc, cond, extra=""):
    global N
    N += 1
    print(f"  [{N:02d}] {desc}: {'OK' if cond else 'FAIL'}" + (f" ({extra})" if extra else ""))
    if not cond:
        raise AssertionError(f"{desc} {extra}")


def _raises(fn):
    try:
        fn()
    except RuntimeError:
        return True
    return False


def sha(data):
    return hashlib.sha256(data if isinstance(data, bytes) else data.encode("utf-8")).hexdigest()


def read(p):
    with open(p, encoding="utf-8") as f:
        return f.read()


def write(p, text):
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        f.write(text)
    return p


# ---------------- 로컬 서버 (메모리 파일 맵) ----------------
SERVED = {}          # 경로 -> bytes
HITS = []            # 요청 경로 기록


class Handler(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        path = urllib.parse.unquote(self.path.lstrip("/"))
        HITS.append(path)
        data = SERVED.get(path)
        if data is None:
            self.send_response(404)
            self.end_headers()
            return
        self.send_response(200)
        self.send_header("Content-Type", "application/octet-stream")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, *a):
        pass


srv = http.server.ThreadingHTTPServer(("127.0.0.1", 0), Handler)
threading.Thread(target=srv.serve_forever, daemon=True).start()
BASE_URL = f"http://127.0.0.1:{srv.server_address[1]}/"


def serve(files, version="2.9.0", notes="테스트 릴리스", sha_map=True, extra=None):
    """files: {저장소 상대경로: 텍스트 또는 bytes}. version.json을 자동 생성해 서빙.
    .py/.pyw → files, .json → data_files, 그 밖(xlsx·pdf) → set_files(있을 때만)."""
    SERVED.clear()
    prog = {k: v for k, v in files.items() if k.endswith((".py", ".pyw"))}
    data = {k: v for k, v in files.items() if k.endswith(".json")}
    sets = {k: v for k, v in files.items() if k not in prog and k not in data}
    info = {"version": version, "notes": notes,
            "files": {("시험장/" + k if k in ("시험장.py", "코코시험장.pyw") else "채점/" + k): k
                      for k in prog},
            "data_files": {k: k for k in data}}
    if sets:
        info["set_files"] = {k: k for k in sets}
    if sha_map:
        info["sha256"] = {k: sha(v) for k, v in files.items()}
    if extra:
        info.update(extra)
    for k, v in files.items():
        SERVED[k] = v if isinstance(v, bytes) else v.encode("utf-8")
    SERVED["version.json"] = json.dumps(info, ensure_ascii=False).encode("utf-8")
    return info


def fresh_install(version="2.3.0"):
    """<루트>/시험장/시험장.py + 코코시험장.pyw, <루트>/채점/grade.py + 코코채점.pyw."""
    root = tempfile.mkdtemp(prefix="coco_inst_", dir=TMP)
    base = os.path.join(root, "시험장")
    write(os.path.join(base, "시험장.py"), f'__version__ = "{version}"\nprint("old")\n')
    write(os.path.join(base, "코코시험장.pyw"), '__version__ = "1.8.0"\n')
    write(os.path.join(root, "채점", "grade.py"), '__version__ = "2.0.2"\n')
    write(os.path.join(root, "채점", "코코채점.pyw"), '__version__ = "1.6.0"\n')
    return root, base


def snapshot(root):
    """{상대경로: 내용} — UTF-8 텍스트면 str, 바이너리(xlsx·pdf)면 bytes."""
    out = {}
    for dp, dn, fns in os.walk(root):
        for fn in fns:
            p = os.path.join(dp, fn)
            with open(p, "rb") as f:
                raw = f.read()
            try:
                out[os.path.relpath(p, root)] = raw.decode("utf-8")
            except UnicodeDecodeError:
                out[os.path.relpath(p, root)] = raw
    return out


NEW_FILES = {
    "시험장.py": '__version__ = "2.9.0"\nprint("new")\n',
    "코코시험장.pyw": '__version__ = "1.9.0"\n',
    "grade.py": '__version__ = "2.1.0"\n',
    "코코채점.pyw": '__version__ = "1.7.0"\n',
    "기대값/코코모의고사1회_기대값.json": '{"points": {"계산작업": 40}}\n',
    "기대값/2026_1회_기대값.json": '{"groups": {"계산작업": [["D14"]]}}\n',
}

print("1. version.json 조회·비교")
SERVED.clear()
check("서빙 파일 없음(404) → None", sj.fetch_update_info(BASE_URL, timeout=2) is None)
check("연결 불가(오프라인) → None (예외 없음)",
      sj.fetch_update_info("http://127.0.0.1:1/", timeout=1) is None)
info = serve(NEW_FILES)
got = sj.fetch_update_info(BASE_URL, timeout=2)
check("version.json 조회 + files/data_files/sha256 dict 보장",
      got and got["version"] == "2.9.0" and len(got["files"]) == 4 and len(got["data_files"]) == 2
      and len(got["sha256"]) == 6)
check("버전 비교: 2.9.0 > 2.3.0 / 2.3.0 == 2.3.0 아님",
      sj.update_available(got, "2.3.0") and not sj.update_available(got, "2.9.0")
      and not sj.update_available(got, "2.10.0"))
SERVED["version.json"] = b"{not json"
check("깨진 version.json → None", sj.fetch_update_info(BASE_URL, timeout=2) is None)
check("version_marker 파싱", sj.version_marker('# x\n__version__ = "2.3.0"\n') == "2.3.0"
      and sj.version_marker("no marker") is None)

print("2. 정상 업데이트 흐름 (files + data_files, sha256 검증, 백업, 안내 기록)")
info = serve(NEW_FILES)
root, base = fresh_install()
before = snapshot(root)
ok, msg, staged = sj.stage_update(info, BASE_URL, base_dir=base, timeout=5)
check("스테이징 성공: 6개 .new", ok and len(staged) == 6, msg)
check("스테이징 중에는 기존 파일 그대로 + .new 존재",
      snapshot(root)[os.path.join("시험장", "시험장.py")] == before[os.path.join("시험장", "시험장.py")]
      and os.path.isfile(os.path.join(base, "시험장.py.new"))
      and os.path.isfile(os.path.join(root, "기대값", "2026_1회_기대값.json.new")))
ok, msg = sj.commit_staged(staged, info["version"])
after = snapshot(root)
check("교체 성공 메시지", ok and "프로그램 4개 + 기대값 2개" in msg and "v2.9.0" in msg, msg)
check("프로그램 4개 교체", after[os.path.join("시험장", "시험장.py")] == NEW_FILES["시험장.py"]
      and after[os.path.join("채점", "grade.py")] == NEW_FILES["grade.py"]
      and after[os.path.join("채점", "코코채점.pyw")] == NEW_FILES["코코채점.pyw"]
      and after[os.path.join("시험장", "코코시험장.pyw")] == NEW_FILES["코코시험장.pyw"])
check("기대값 2개는 <루트>/기대값/ 에 생성",
      after.get(os.path.join("기대값", "코코모의고사1회_기대값.json")) == NEW_FILES["기대값/코코모의고사1회_기대값.json"]
      and after.get(os.path.join("기대값", "2026_1회_기대값.json")) == NEW_FILES["기대값/2026_1회_기대값.json"])
check(".bak 백업 4개(기존 파일만) + .new 잔여 없음",
      after[os.path.join("시험장", "시험장.py.bak")] == before[os.path.join("시험장", "시험장.py")]
      and after[os.path.join("채점", "grade.py.bak")] == before[os.path.join("채점", "grade.py")]
      and not any(k.endswith(".new") for k in after)
      and not any(k.endswith(".bak") for k in after if k.startswith("기대값")))
check("expected_values_dir: 시험장/ 구조 → <루트>/기대값",
      sj.expected_values_dir(base) == os.path.join(root, "기대값"))
flat = tempfile.mkdtemp(prefix="flat_", dir=TMP)
check("expected_values_dir: 평면 구조 → <폴더>/기대값",
      sj.expected_values_dir(flat) == os.path.join(flat, "기대값"))
np_ = os.path.join(TMP, "notice.json")
sj.write_update_notice("2.3.0", "2.9.0", "테스트 릴리스", path=np_)
check("안내 기록: 다른 버전이 읽으면 None", sj.consume_update_notice("2.3.0", path=np_) is None)
d = sj.consume_update_notice("2.9.0", path=np_)
check("안내 기록: 새 버전이 1회 소비", d and d["from"] == "2.3.0" and d["notes"] == "테스트 릴리스"
      and sj.consume_update_notice("2.9.0", path=np_) is None
      and sj.read_update_notice(np_)["shown"] is True)
# apply_update(호환 API) 한 번에
root2, base2 = fresh_install()
ok, msg = sj.apply_update(info, BASE_URL, base_dir=base2, timeout=5)
check("apply_update(구 API) = stage + commit", ok and read(os.path.join(base2, "시험장.py")) == NEW_FILES["시험장.py"])
# sha256 맵 없이도 동작
info_nosha = serve(NEW_FILES, sha_map=False)
root3, base3 = fresh_install()
ok, msg = sj.apply_update(info_nosha, BASE_URL, base_dir=base3, timeout=5)
check("sha256 맵 없으면 해시 검증 생략하고 적용", ok, msg)

print("3. 검증 실패 → 전체 롤백")


def rollback_case(desc, files, version="2.9.0", sha_override=None, expect_msg=""):
    info = serve(files, version=version)
    if sha_override:
        info["sha256"].update(sha_override)
        SERVED["version.json"] = json.dumps(info, ensure_ascii=False).encode("utf-8")
    root, base = fresh_install()
    before = snapshot(root)
    ok, msg = sj.apply_update(info, BASE_URL, base_dir=base, timeout=5)
    after = snapshot(root)
    check(desc, (not ok) and after == before and (expect_msg in msg), msg)


rollback_case("sha256 불일치 → 적용 안 함", NEW_FILES, sha_override={"grade.py": "00" * 32}, expect_msg="sha256 불일치")
rollback_case("기대값 JSON 깨짐 → 프로그램 포함 전체 롤백",
              dict(NEW_FILES, **{"기대값/2026_1회_기대값.json": "{broken"}), expect_msg="JSON 형식 오류")
rollback_case("py 문법 오류 → 전체 롤백",
              dict(NEW_FILES, **{"grade.py": '__version__ = "2.1.0"\ndef (:\n'}), expect_msg="업데이트 실패")
rollback_case("__version__ 표식 없는 프로그램 파일 → 롤백",
              dict(NEW_FILES, **{"코코채점.pyw": "print(1)\n"}), expect_msg="__version__")
rollback_case("시험장.py 버전이 version.json과 다르면(재시작 반복 위험) 롤백",
              dict(NEW_FILES, **{"시험장.py": '__version__ = "2.8.0"\n'}), expect_msg="재시작 반복 방지")
# 다운로드 중 404 (파일 목록에 있는데 서버에 없음)
info = serve(NEW_FILES)
del SERVED["기대값/2026_1회_기대값.json"]
root, base = fresh_install()
before = snapshot(root)
ok, msg = sj.apply_update(info, BASE_URL, base_dir=base, timeout=5)
check("파일 하나 404 → 전체 롤백 + .new 정리", not ok and snapshot(root) == before, msg)
check("빈 목록 → 실패", sj.apply_update({"version": "9"}, BASE_URL, base_dir=base)[0] is False)
# 교체 단계 실패 롤백 (os.replace 강제 실패)
info = serve(NEW_FILES)
root, base = fresh_install()
before = snapshot(root)
ok, msg, staged = sj.stage_update(info, BASE_URL, base_dir=base, timeout=5)
_orig_replace = sj.os.replace
calls = [0]


def flaky_replace(a, b):
    calls[0] += 1
    if calls[0] == 3:
        raise OSError("교체 실패 주입")
    return _orig_replace(a, b)


sj.os.replace = flaky_replace
try:
    ok, msg = sj.commit_staged(staged, "2.9.0")
finally:
    sj.os.replace = _orig_replace
after = {k: v for k, v in snapshot(root).items() if not k.endswith(".bak")}
check("교체 도중 실패 → 이미 바꾼 파일 .bak로 복원 + 새로 만든 파일 제거",
      not ok and after == before and "교체 실패 주입" in msg, msg)

print("4. 같은 버전: 기대값만 동기화")
info = serve(NEW_FILES, version="2.3.0")
root, base = fresh_install("2.3.0")
todo = sj.data_files_to_sync(info, base_dir=base)
check("로컬에 없는 기대값 2개가 동기화 대상", len(todo) == 2)
ok, n, msg = sj.sync_data_files(info, BASE_URL, base_dir=base, timeout=5)
check("기대값 2개 내려받음(프로그램은 건드리지 않음)", ok and n == 2
      and read(os.path.join(root, "기대값", "2026_1회_기대값.json")) == NEW_FILES["기대값/2026_1회_기대값.json"]
      and read(os.path.join(base, "시험장.py")).startswith('__version__ = "2.3.0"'), msg)
check("다시 확인하면 대상 없음", sj.data_files_to_sync(info, base_dir=base) == []
      and sj.sync_data_files(info, BASE_URL, base_dir=base)[1] == 0)
changed = dict(NEW_FILES, **{"기대값/2026_1회_기대값.json": '{"groups": {"계산작업": [["D15"]]}}\n'})
info = serve(changed, version="2.3.0")
check("원격 해시가 바뀐 기대값만 대상", [t[1] for t in sj.data_files_to_sync(info, base_dir=base)] == ["기대값/2026_1회_기대값.json"])
ok, n, msg = sj.sync_data_files(info, BASE_URL, base_dir=base, timeout=5)
check("바뀐 기대값 1개 갱신 + .bak", ok and n == 1
      and read(os.path.join(root, "기대값", "2026_1회_기대값.json")) == changed["기대값/2026_1회_기대값.json"]
      and os.path.isfile(os.path.join(root, "기대값", "2026_1회_기대값.json.bak")))
removed = {k: v for k, v in changed.items() if k != "기대값/코코모의고사1회_기대값.json"}
info = serve(removed, version="2.3.0")
ok, n, msg = sj.sync_data_files(info, BASE_URL, base_dir=base, timeout=5)
check("원격에서 사라진 기대값은 로컬에서 지우지 않음", ok and n == 0
      and os.path.isfile(os.path.join(root, "기대값", "코코모의고사1회_기대값.json")))
info = serve(dict(changed, **{"기대값/2026_1회_기대값.json": "{broken"}), version="2.3.0")
ok, n, msg = sj.sync_data_files(info, BASE_URL, base_dir=base, timeout=5)
check("깨진 기대값은 받지 않음(기존 유지)", not ok and n == 0
      and read(os.path.join(root, "기대값", "2026_1회_기대값.json")) == changed["기대값/2026_1회_기대값.json"], msg)
nosha = serve(NEW_FILES, version="2.3.0", sha_map=False)
check("sha256 맵 없으면 로컬에 없는 파일만 대상",
      [t[1] for t in sj.data_files_to_sync(nosha, base_dir=base)] == [])

print("5. UpdateCoordinator (상태 기계)")
logs = []


def make(root_base, enabled=True, current="2.3.0"):
    root, base = root_base
    return sj.UpdateCoordinator(BASE_URL, base_dir=base, enabled=lambda: enabled,
                                current=current, log=logs.append,
                                notice_path=os.path.join(root, "notice.json"))


info = serve(NEW_FILES)
HITS.clear()
inst = fresh_install()
c_off = make(inst, enabled=False)
check("토글 off → 서버 미접속(disabled) + 로그", c_off.check() == "disabled" and HITS == []
      and any("꺼짐" in l for l in logs))
check("토글 off여도 수동(force)은 확인·스테이징", c_off.check(force=True) == "staged" and c_off.pending
      and "version.json" in HITS)
c_off.pending = None
inst = fresh_install()
c = make(inst)
HITS.clear()
check("자동 확인: 새 버전 → staged (다운로드·검증 완료, 아직 미적용)",
      c.check() == "staged" and len(c.pending) == 6
      and read(os.path.join(inst[1], "시험장.py")).startswith('__version__ = "2.3.0"'))
check("같은 실행에서 중복 확인 방지", c.check() == "already" and HITS.count("version.json") == 1)
check("시험 중 → deferred (파일 그대로, pending 유지)",
      c.apply_pending(exam_running=True) == "deferred" and c.pending
      and read(os.path.join(inst[1], "시험장.py")).startswith('__version__ = "2.3.0"')
      and any("시험 진행 중" in l for l in logs))
check("시험 종료 후 → applied + 안내 기록", c.apply_pending(exam_running=False) == "applied"
      and c.pending is None and c.applied_version == "2.9.0"
      and read(os.path.join(inst[1], "시험장.py")) == NEW_FILES["시험장.py"]
      and sj.read_update_notice(c.notice_path)["to"] == "2.9.0")
check("적용할 것이 없으면 none", c.apply_pending() == "none")
# 재시작된 새 인스턴스: 버전이 2.9.0이면 안내 소비 + latest
c2 = make(inst, current="2.9.0")
check("재시작 후(v2.9.0): 안내 소비 + 최신", sj.consume_update_notice("2.9.0", path=c2.notice_path)["notes"] == "테스트 릴리스"
      and c2.check() == "latest")
# 재시작 반복 방지: 적용했는데 실행 버전이 그대로면 자동 적용 중단
c3 = make(inst, current="2.3.0")
check("적용 후 버전 미상승 → stuck (자동 재적용 안 함)", c3.check() == "stuck" and c3.pending is None)
check("stuck이어도 수동(force)은 재시도", make(inst, current="2.3.0").check(force=True) == "staged")
# 오프라인 / 실패 / data
c4 = sj.UpdateCoordinator("http://127.0.0.1:1/", base_dir=inst[1], enabled=lambda: True,
                          current="2.3.0", log=logs.append, notice_path=os.path.join(inst[0], "n2.json"))
check("오프라인 → offline (조용히)", c4.check(timeout=1) == "offline" and c4.pending is None)
serve(dict(NEW_FILES, **{"grade.py": "def (:\n"}))
inst5 = fresh_install()
c5 = make(inst5)
check("검증 실패 → failed + 상세", c5.check() == "failed" and "__version__" in c5.detail and c5.pending is None)
serve(NEW_FILES, version="2.3.0")
inst6 = fresh_install()
c6 = make(inst6)
check("같은 버전 + 기대값 없음 → data (2개 동기화)", c6.check() == "data" and c6.data_synced == 2
      and os.path.isfile(os.path.join(inst6[0], "기대값", "2026_1회_기대값.json")))
c6b = make(inst6)
check("기대값까지 최신이면 latest", c6b.check() == "latest")
check("시작 로그에 전 과정 기록", any("새 버전 v2.9.0 발견" in l for l in logs)
      and any("검증 통과" in l for l in logs) and any("업데이트 적용" in l for l in logs)
      and any("기대값 자동 갱신" in l for l in logs) and any("최신 버전" in l for l in logs)
      and any("서버 응답 없음" in l for l in logs))
# 자동업데이트 설정 기본값/토글 (세트설정.json _설정)
cfgp = os.path.join(TMP, "세트설정.json")
check("자동업데이트 기본 true", sj.auto_update_enabled(cfgp) is True)
sj.set_app_setting(sj.AUTO_UPDATE_SETTING, False, path=cfgp)
check("토글 off 저장/읽기", sj.auto_update_enabled(cfgp) is False
      and json.load(open(cfgp, encoding="utf-8"))["_설정"]["자동업데이트"] is False)

print("6. 2.2.3 호환: 구버전 apply_update 로직으로 새 version.json 적용")


def legacy_apply_update(info, base_url, base_dir, timeout=5):
    """2.2.3 시험장.py apply_update 그대로 (files만, __version__ 표식 + py_compile)."""
    import py_compile
    import urllib.request
    files = (info or {}).get("files") or {}
    if not files:
        return False, "업데이트 파일 목록이 비어 있습니다."
    staged = []
    replaced = []
    try:
        for rel_key, repo_rel in files.items():
            target = sj._update_target_path(rel_key, base_dir)
            url = base_url + urllib.parse.quote(str(repo_rel))
            req = urllib.request.Request(url, headers={"User-Agent": "cocomate-updater"})
            with urllib.request.urlopen(req, timeout=timeout) as r:
                data = r.read()
            text = data.decode("utf-8")
            if "__version__" not in text:
                raise RuntimeError(f"{repo_rel}: __version__ 표식이 없어 배포 파일이 아닌 것으로 판단")
            tmp = target + ".new"
            os.makedirs(os.path.dirname(target) or ".", exist_ok=True)
            with open(tmp, "wb") as f:
                f.write(data)
            staged.append((target, tmp))
            py_compile.compile(tmp, cfile=tmp + "c", doraise=True)
            os.remove(tmp + "c")
        for target, tmp in staged:
            if os.path.isfile(target):
                shutil.copy2(target, target + ".bak")
            os.replace(tmp, target)
            replaced.append(target)
        return True, f"{len(replaced)}개 파일을 v{info.get('version')}(으)로 업데이트했습니다."
    except Exception as e:
        return False, f"업데이트 실패(기존 버전 유지): {e}"


info = serve(NEW_FILES)
root, base = fresh_install("2.2.3")
ok, msg = legacy_apply_update(info, BASE_URL, base)
check("2.2.3 apply_update: files 4개 적용 성공(data_files·sha256 무시)", ok and "4개 파일" in msg
      and read(os.path.join(base, "시험장.py")) == NEW_FILES["시험장.py"]
      and not os.path.isdir(os.path.join(root, "기대값")), msg)
bad = dict(info)
bad["files"] = dict(info["files"], **{"기대값/2026_1회_기대값.json": "기대값/2026_1회_기대값.json"})
root, base = fresh_install("2.2.3")
ok, msg = legacy_apply_update(bad, BASE_URL, base)
check("(대조) JSON을 files에 넣으면 2.2.3은 실패 → data_files 분리가 필요한 이유",
      not ok and "__version__" in msg, msg)
# 새 인스턴스(2.3.0 코드)가 2.2.3이 남긴 상태에서 기대값을 채움
root, base = fresh_install("2.2.3")
legacy_apply_update(info, BASE_URL, base)
c7 = sj.UpdateCoordinator(BASE_URL, base_dir=base, enabled=lambda: True, current="2.9.0",
                          log=logs.append, notice_path=os.path.join(root, "n.json"))
check("2.2.3이 프로그램만 올린 뒤 재시작된 새 버전이 기대값을 자동 동기화",
      c7.check() == "data" and c7.data_synced == 2)

# 실제 배포 파일로 version.json 호환 확인 (공개 저장소 클론이 있으면)
# 개발 저장소에서는 옆 폴더의 공개 클론을, 공개 배포본에서는 자기 자신을 본다.
_pub_cands = [os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(BASE))),
                           "cocomate-study-tools", "version.json"),
              os.path.join(BASE, "version.json")]
pub = next((c for c in _pub_cands if os.path.isfile(c)), _pub_cands[0])
if os.path.isfile(pub):
    vj = json.load(open(pub, encoding="utf-8"))
    check("공개 version.json: files에는 .py/.pyw만, JSON은 data_files에만",
          all(k.endswith((".py", ".pyw")) for k in vj.get("files", {}).values())
          and all(k.endswith(".json") for k in vj.get("data_files", {}).values()), str(list(vj.get("files", {}).values())))

print("7. 코코채점.pyw (단독 채점 래퍼) — 기대값 연결 + 자동 업데이트")
import importlib.machinery
# 개발 저장소는 <루트>/채점/코코채점.pyw, 배포본은 스위트와 같은 폴더에 있다.
_kc_cands = [os.path.join(os.path.dirname(BASE), "채점", "코코채점.pyw"),
             os.path.join(BASE, "코코채점.pyw")]
_kc_path = next((c for c in _kc_cands if os.path.isfile(c)), None)
if _kc_path is None:
    raise SystemExit("코코채점.pyw 를 찾지 못했습니다: " + " / ".join(_kc_cands))
_kc_spec = importlib.util.spec_from_loader("kc", importlib.machinery.SourceFileLoader("kc", _kc_path))
kc = importlib.util.module_from_spec(_kc_spec)
_kc_spec.loader.exec_module(kc)
kc.LOG_PATH = os.path.join(TMP, "코코채점_로그.txt")
kc.INSTALLED_VERSION_PATH = os.path.join(TMP, ".installed_version")
KROOT = tempfile.mkdtemp(prefix="kc_", dir=TMP)
KD = os.path.join(KROOT, "기대값")
for rel in ("모의고사/코코모의고사1회_문제.xlsx", "모의고사/코코모의고사1회_기대값.json",
            "기출/코코모의고사2회_문제.xlsx", "세트/2024년 상시2회 2급_문제.xlsm",
            "세트/2024년 상시1회 2급_문제.xlsm", "세트/2026 컴활2급실기_상시복원기출문제 1회_문제.xlsm",
            "세트/2026_1회_기대값.json",
            "기대값/코코모의고사1회_기대값.json", "기대값/코코모의고사2회_기대값.json",
            "기대값/2024상시2회_기대값.json", "기대값/2026_1회_기대값.json"):
    write(os.path.join(KROOT, *rel.split("/")), "{}")
P = lambda *a: os.path.join(KROOT, *a)
check("코코채점 detect_key: 문제 파일 폴더의 정확 키 우선",
      kc.detect_key(P("모의고사", "코코모의고사1회_문제.xlsx"), key_dirs=[KD])
      == (P("모의고사", "코코모의고사1회_기대값.json"), "문제 파일 폴더"))
check("코코채점 detect_key: 기대값 폴더의 정확 키",
      kc.detect_key(P("기출", "코코모의고사2회_문제.xlsx"), key_dirs=[KD])
      == (P("기대값", "코코모의고사2회_기대값.json"), "기대값 폴더"))
check("코코채점 detect_key: 토큰(2024·상시·2회) 유일 일치",
      kc.detect_key(P("세트", "2024년 상시2회 2급_문제.xlsm"), key_dirs=[KD])
      == (P("기대값", "2024상시2회_기대값.json"), "토큰 일치"))
check("코코채점 detect_key: 같은 이름이 두 곳에 있으면 문제 파일 폴더 쪽",
      kc.detect_key(P("세트", "2026 컴활2급실기_상시복원기출문제 1회_문제.xlsm"), key_dirs=[KD])
      == (P("세트", "2026_1회_기대값.json"), "토큰 일치"))
check("코코채점 detect_key: 상시 1회는 충돌로 미연결",
      kc.detect_key(P("세트", "2024년 상시1회 2급_문제.xlsm"), key_dirs=[KD]) == (None, None))
check("코코채점 detect_key: 기대값 폴더 없어도 안전",
      kc.detect_key(P("세트", "2024년 상시2회 2급_문제.xlsm"), key_dirs=[P("없음")]) == (None, None))
check("코코채점 expected_values_dir: 채점/ 구조 → 루트/기대값",
      kc.expected_values_dir(P("채점")) == P("기대값"))
# 자동 업데이트 (묻지 않음)
info = serve(NEW_FILES)
root, base = fresh_install("2.3.0")
kbase = os.path.join(root, "채점")
check("코코채점 current_version: 시험장.py __version__ 기준", kc.current_version(kbase) == "2.3.0")
sj.set_app_setting(sj.AUTO_UPDATE_SETTING, False, path=os.path.join(base, "세트설정.json"))
check("코코채점 토글 off(시험장/세트설정.json) → disabled",
      kc.auto_update_enabled(kbase) is False and kc.run_auto_update(BASE_URL, kbase)[0] == "disabled")
sj.set_app_setting(sj.AUTO_UPDATE_SETTING, True, path=os.path.join(base, "세트설정.json"))
state, note = kc.run_auto_update(BASE_URL, kbase)
check("코코채점 자동 업데이트: 새 버전 → applied (프로그램 4개 + 기대값 2개)",
      state == "applied" and "v2.9.0" in note
      and read(os.path.join(root, "채점", "grade.py")) == NEW_FILES["grade.py"]
      and read(os.path.join(base, "시험장.py")) == NEW_FILES["시험장.py"]
      and read(os.path.join(root, "기대값", "2026_1회_기대값.json")) == NEW_FILES["기대값/2026_1회_기대값.json"], note)
check("코코채점: 적용 후 현재 버전 2.9.0 → latest", kc.current_version(kbase) == "2.9.0"
      and kc.run_auto_update(BASE_URL, kbase)[0] == "latest")
serve(dict(NEW_FILES, **{"기대값/2026_1회_기대값.json": '{"x": 2}'}))
state, note = kc.run_auto_update(BASE_URL, kbase)
check("코코채점: 같은 버전 + 바뀐 기대값 → data 1개", state == "data" and "1개" in note, note)
check("코코채점: 오프라인 → offline", kc.run_auto_update("http://127.0.0.1:1/", kbase, timeout=1)[0] == "offline")
serve(dict(NEW_FILES, **{"grade.py": "def (:\n"}), version="3.0.0")
before = snapshot(root)
state, note = kc.run_auto_update(BASE_URL, kbase)
check("코코채점: 검증 실패 → failed + 롤백", state == "failed" and snapshot(root) == before, note)
# 시험장.py 없는 채점 전용 설치: .installed_version 기록으로 반복 적용 방지
serve(NEW_FILES)
solo = tempfile.mkdtemp(prefix="solo_", dir=TMP)
write(os.path.join(solo, "grade.py"), '__version__ = "2.0.2"\n')
write(os.path.join(solo, "코코채점.pyw"), '__version__ = "1.6.0"\n')
kc.INSTALLED_VERSION_PATH = os.path.join(solo, ".installed_version")
check("채점 전용 설치: 시험장.py 없으면 0.0.0 → 1회 적용 후 기록", kc.current_version(solo) == "0.0.0"
      and kc.run_auto_update(BASE_URL, solo)[0] == "applied" and kc.current_version(solo) == "2.9.0"
      and kc.run_auto_update(BASE_URL, solo)[0] == "latest")
check("코코채점 로그 기록", "새 버전 v2.9.0 발견" in read(kc.LOG_PATH) and "업데이트 적용" in read(kc.LOG_PATH))

print("8. set_files: 코코 모의고사 세트 파일(xlsx·pdf) 자동 배포 (v2.3.1)")
import io
import zipfile


def fake_xlsx(tag):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("xl/workbook.xml", f"<workbook>{tag}</workbook>")
        zf.writestr("bin.dat", bytes(range(256)))          # UTF-8로 디코드되지 않는 바이트
    return buf.getvalue()


SET_FILES = {
    "모의고사/코코모의고사1회_문제.xlsx": fake_xlsx("1문제"),
    "모의고사/코코모의고사1회_정답.xlsx": fake_xlsx("1정답"),
    "모의고사/코코모의고사1회_문제지.pdf": b"%PDF-1.4\nround1 \xff\xfe\n",
    "모의고사/코코모의고사2회_문제.xlsx": fake_xlsx("2문제"),
    "모의고사/코코모의고사2회_정답.xlsx": fake_xlsx("2정답"),
    "모의고사/코코모의고사2회_문제지.pdf": b"%PDF-1.4\nround2 \xff\xfe\n",
}
ALL_FILES = dict(NEW_FILES, **SET_FILES)


def readb(p):
    with open(p, "rb") as f:
        return f.read()


info = serve(ALL_FILES)
got = sj.fetch_update_info(BASE_URL, timeout=2)
check("version.json 조회: set_files 6개 dict 보장(+files 4·data_files 2)",
      got and len(got["set_files"]) == 6 and len(got["files"]) == 4 and len(got["data_files"]) == 2)
root, base = fresh_install()
check("_data_target_path: 모의고사/… → <루트>/모의고사/…, 기대값/… → <루트>/기대값/…, 폴더 없는 키 → 기대값/",
      sj._data_target_path("모의고사/코코모의고사1회_문제.xlsx", base) == os.path.join(root, "모의고사", "코코모의고사1회_문제.xlsx")
      and sj._data_target_path("기대값/x_기대값.json", base) == os.path.join(root, "기대값", "x_기대값.json")
      and sj._data_target_path("x_기대값.json", base) == os.path.join(root, "기대값", "x_기대값.json"))
flat = tempfile.mkdtemp(prefix="flat_", dir=TMP)
check("평면 구조: <폴더>/모의고사/…", sj._data_target_path("모의고사/a.pdf", flat) == os.path.join(flat, "모의고사", "a.pdf"))
ok, msg, staged = sj.stage_update(got, BASE_URL, base_dir=base, timeout=5)
check("새 버전 스테이징: 프로그램 4 + 기대값 2 + 세트 파일 6 = 12개 .new", ok and len(staged) == 12, msg)
ok, msg = sj.commit_staged(staged, got["version"])
check("교체 메시지에 '세트 파일 6개'", ok and "프로그램 4개 + 기대값 2개 + 세트 파일 6개" in msg, msg)
check("세트 파일 6개가 <루트>/모의고사/에 바이트 그대로 생성",
      all(readb(os.path.join(root, *k.split("/"))) == v for k, v in SET_FILES.items())
      and not os.path.isdir(os.path.join(root, "기대값", "모의고사")))
# 같은 버전: 세트 파일만 동기화
info = serve(ALL_FILES, version="2.3.1")
root, base = fresh_install("2.3.1")
todo = sj.data_files_to_sync(info, base_dir=base)
check("같은 버전: 로컬에 없는 기대값 2 + 세트 파일 6 = 8개가 동기화 대상", len(todo) == 8
      and sum(1 for t in todo if t[1].startswith("모의고사/")) == 6)
ok, n, msg = sj.sync_data_files(info, BASE_URL, base_dir=base, timeout=5)
check("동기화: 8개 내려받음(프로그램은 그대로) + 세트 파일 내용 일치", ok and n == 8
      and readb(os.path.join(root, "모의고사", "코코모의고사2회_문제지.pdf")) == SET_FILES["모의고사/코코모의고사2회_문제지.pdf"]
      and read(os.path.join(base, "시험장.py")).startswith('__version__ = "2.3.1"'), msg)
check("다시 확인하면 대상 없음(sha256 일치)", sj.data_files_to_sync(info, base_dir=base) == [])
changed_set = dict(ALL_FILES, **{"모의고사/코코모의고사1회_문제.xlsx": fake_xlsx("1문제-수정")})
info = serve(changed_set, version="2.3.1")
ok, n, msg = sj.sync_data_files(info, BASE_URL, base_dir=base, timeout=5)
check("원격 해시가 바뀐 세트 파일 1개만 갱신 + .bak", ok and n == 1
      and readb(os.path.join(root, "모의고사", "코코모의고사1회_문제.xlsx")) == changed_set["모의고사/코코모의고사1회_문제.xlsx"]
      and os.path.isfile(os.path.join(root, "모의고사", "코코모의고사1회_문제.xlsx.bak")), msg)
# 검증 실패
before = snapshot(root)
bad = serve(changed_set, version="2.3.1")
bad["sha256"]["모의고사/코코모의고사2회_정답.xlsx"] = "0" * 64
SERVED["version.json"] = json.dumps(bad, ensure_ascii=False).encode("utf-8")
os.remove(os.path.join(root, "모의고사", "코코모의고사2회_정답.xlsx"))
ok, n, msg = sj.sync_data_files(bad, BASE_URL, base_dir=base, timeout=5)
check("sha256 불일치 → 실패, 아무 파일도 안 바뀜", not ok and n == 0 and "sha256 불일치" in msg
      and not os.path.isfile(os.path.join(root, "모의고사", "코코모의고사2회_정답.xlsx"))
      and not any(k.endswith(".new") for k in snapshot(root)), msg)
nosha = serve(changed_set, version="2.3.1", sha_map=False)
ok, n, msg = sj.sync_data_files(nosha, BASE_URL, base_dir=base, timeout=5)
check("sha256 맵 없음 → 바이너리는 검증 불가로 실패(로컬에 없는 파일이 대상이었음)",
      not ok and "sha256 항목이 없어" in msg, msg)
wrong_magic = serve(dict(changed_set, **{"모의고사/코코모의고사2회_정답.xlsx": b"<html>not a zip</html>"}), version="2.3.1")
ok, n, msg = sj.sync_data_files(wrong_magic, BASE_URL, base_dir=base, timeout=5)
check("해시는 맞아도 xlsx(zip) 머리가 아니면 실패", not ok and "형식이 아닙니다" in msg, msg)
wrong_pdf = serve(dict(changed_set, **{"모의고사/코코모의고사2회_정답.xlsx": fake_xlsx("2정답"),
                                       "모의고사/코코모의고사2회_문제지.pdf": b"not a pdf"}), version="2.3.1")
os.remove(os.path.join(root, "모의고사", "코코모의고사2회_문제지.pdf"))
ok, n, msg = sj.sync_data_files(wrong_pdf, BASE_URL, base_dir=base, timeout=5)
check("PDF 머리(%PDF)가 아니면 실패", not ok and ".pdf 파일 형식이 아닙니다" in msg, msg)
# 2.3.0의 data_files 검증(UTF-8 디코드 + json.loads)에 바이너리를 넣으면 실패 → set_files 분리 이유


def legacy_230_verify(data):
    text = data.decode("utf-8")          # 2.3.0 _verify_download 그대로
    json.loads(text)


try:
    legacy_230_verify(SET_FILES["모의고사/코코모의고사1회_문제.xlsx"])
    legacy_ok = True
except Exception:
    legacy_ok = False
check("(대조) xlsx를 2.3.0 data_files 검증에 넣으면 실패 → 바이너리는 set_files에만 두는 이유", not legacy_ok)
check("2.3.1 _verify_download: 바이너리는 sha256+머리만 검사(None 반환), JSON은 텍스트 반환",
      sj._verify_download("모의고사/a.pdf", b"%PDF-1.4\n\xff", "data",
                          {"sha256": {"모의고사/a.pdf": sha(b"%PDF-1.4\n\xff")}}) is None
      and sj._verify_download("기대값/a.json", b"{}", "data", {}) == "{}")
# UpdateCoordinator: 같은 버전 + 세트 파일 없음 → data
info = serve(ALL_FILES, version="2.3.1")
root, base = fresh_install("2.3.1")
c8 = sj.UpdateCoordinator(BASE_URL, base_dir=base, enabled=lambda: True, current="2.3.1",
                          log=logs.append, notice_path=os.path.join(root, "n.json"))
check("UpdateCoordinator: 같은 버전 + 세트 파일 없음 → data (8개 동기화)", c8.check() == "data" and c8.data_synced == 8)
info = serve(ALL_FILES, version="2.9.0")
root, base = fresh_install("2.3.1")
c9 = sj.UpdateCoordinator(BASE_URL, base_dir=base, enabled=lambda: True, current="2.3.1",
                          log=logs.append, notice_path=os.path.join(root, "n.json"))
check("UpdateCoordinator: 새 버전 로그에 '세트 파일 6개'", c9.check() == "staged"
      and any("프로그램 4개 + 기대값 2개 + 세트 파일 6개" in l for l in logs), str(logs[-2:]))
# 2.2.3·코코채점 1.7.0은 set_files를 무시(프로그램·기대값만 적용, 실패 없음)
root, base = fresh_install("2.2.3")
ok, msg = legacy_apply_update(info, BASE_URL, base)
check("2.2.3 apply_update: set_files가 있어도 files 4개만 적용(무시)", ok and "4개 파일" in msg
      and not os.path.isdir(os.path.join(root, "모의고사")), msg)
root, base = fresh_install("2.3.1")
kbase = os.path.join(root, "채점")
kc.INSTALLED_VERSION_PATH = os.path.join(TMP, ".installed_version2")
state, note = kc.run_auto_update(BASE_URL, kbase)
check("코코채점 1.7.0: set_files 무시하고 프로그램+기대값만 applied(모의고사 폴더 안 만듦)",
      state == "applied" and not os.path.isdir(os.path.join(root, "모의고사"))
      and read(os.path.join(root, "기대값", "2026_1회_기대값.json")) == NEW_FILES["기대값/2026_1회_기대값.json"], note)
if os.path.isfile(pub):
    vj = json.load(open(pub, encoding="utf-8"))
    # 세트 폴더는 배포.py SET_DIRS 와 같음 (모의고사/ · 드릴/) — 새 세트 폴더를
    # 배포 대상에 넣으면 여기에도 같이 추가한다.
    SET_REPO_DIRS = ("모의고사/", "드릴/")
    check("공개 version.json: set_files는 xlsx/pdf(모의고사/·드릴/)와 루틴.html(2.4.0+)만, data_files는 JSON만, sha256에 전부 포함",
          all((k.endswith((".xlsx", ".pdf")) and k.startswith(SET_REPO_DIRS)) or k == "루틴.html"
              for k in vj.get("set_files", {}).values())
          and all(k.endswith(".json") for k in vj.get("data_files", {}).values())
          and all(k in vj.get("sha256", {}) for k in list(vj.get("set_files", {}).values()) + list(vj.get("data_files", {}).values())),
          str(list(vj.get("set_files", {}).values())))


print("8. 경로 키 위조 방어 — version.json 이 시키는 위치에 파일을 쓰므로")
# 원격 version.json 이 조작되면 설치 폴더 밖에 쓰려 들 수 있다.
# 어떤 키가 와도 설치 루트를 벗어나지 않거나 거부돼야 한다.
_pt_base = os.path.join(TMP, "설치", "시험장")
os.makedirs(_pt_base, exist_ok=True)
_pt_root = os.path.dirname(_pt_base)
_evil = ["../../../../../../tmp/PWNED.json",
         "..\\..\\Windows\\evil.json",
         "C:\\Windows\\evil.json",
         "/etc/cron.d/evil",
         "정정/../../../../PWNED.json",
         "기대값/..\\..\\evil.json"]
for _fn in ("_data_target_path", "_update_target_path"):
    for _k in _evil:
        try:
            _p = os.path.abspath(getattr(sj, _fn)(_k, _pt_base))
            _inside = _p == _pt_root or _p.startswith(_pt_root + os.sep)
        except RuntimeError:
            _inside = True          # 거부도 정답
        check(f"{_fn}: 설치 폴더를 벗어나지 않는다", _inside, _k)
check("드라이브 문자가 든 키는 거부한다",
      _raises(lambda: sj._safe_rel_parts("C:\\x.json")))
check("비어 있는 키는 거부한다", _raises(lambda: sj._safe_rel_parts("../..")))
check("정상 키는 그대로 통과한다",
      sj._safe_rel_parts("기대값/a_기대값.json") == ["기대값", "a_기대값.json"])

srv.shutdown()
shutil.rmtree(TMP, ignore_errors=True)
print()
print(f"자동 업데이트 테스트 {N}건 전부 통과")
