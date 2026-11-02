"""Windows 에서 skill-creator 스크립트를 돌리기 위한 보정 (PYTHONPATH 로 연결하면 모든 파이썬 프로세스가 자동으로 읽는다).

skill-creator 의 scripts/run_eval.py 는 `claude -p` 출력 파이프를 select.select 로 기다리는데,
Windows 의 select 는 소켓만 받아서 파이프를 넘기면 OSError(WinError 10093/10038)로 죽는다.
run_eval 은 여러 자식 프로세스에서 돌기 때문에 부모만 고쳐서는 안 되고, 여기처럼 시작 시점에 고쳐야 한다.

파이프가 섞여 들어오면 PeekNamedPipe 로 읽을 데이터가 있는지 확인하는 방식으로 대신 기다린다.
소켓만 들어오면 원래 select 를 그대로 쓴다.
"""
import sys

if sys.platform == "win32":
    import ctypes
    import msvcrt
    import select
    import time
    from ctypes import wintypes

    _orig_select = select.select
    _peek = ctypes.windll.kernel32.PeekNamedPipe
    _peek.argtypes = [wintypes.HANDLE, ctypes.c_void_p, wintypes.DWORD, ctypes.c_void_p,
                      ctypes.POINTER(wintypes.DWORD), ctypes.c_void_p]
    _peek.restype = wintypes.BOOL

    def _pipe_ready(f) -> bool:
        fd = f if isinstance(f, int) else f.fileno()
        avail = wintypes.DWORD(0)
        if not _peek(msvcrt.get_osfhandle(fd), None, 0, None, ctypes.byref(avail), None):
            return True  # 파이프가 닫힘(EOF) - 읽으면 b"" 가 나오므로 호출한 쪽이 끝을 알아챈다
        return avail.value > 0

    def _select_windows(rlist, wlist, xlist, timeout=None):
        try:
            return _orig_select(rlist, wlist, xlist, timeout)
        except OSError:
            pass  # 파이프가 섞여 있음 - 아래에서 직접 기다린다
        deadline = None if timeout is None else time.monotonic() + timeout
        while True:
            ready = [f for f in rlist if _pipe_ready(f)]
            if ready or (deadline is not None and time.monotonic() >= deadline):
                return ready, [], []
            time.sleep(0.05)

    select.select = _select_windows
