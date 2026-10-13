"""Per-preparation progress and cooperative cancellation; no global job state."""
from contextvars import ContextVar
from contextlib import contextmanager


class PreparationCancelled(RuntimeError):
    pass


_current = ContextVar('xlmerge_preparation', default=None)


@contextmanager
def preparation_context(report, cancelled):
    token = _current.set((report, cancelled))
    try:
        yield
    finally:
        _current.reset(token)


def checkpoint(**progress):
    context = _current.get()
    if context is None:
        return
    report, cancelled = context
    if cancelled.is_set():
        raise PreparationCancelled('解析已中断，工作树未修改。')
    if progress:
        report(progress)


def cancellable():
    return _current.get() is not None
