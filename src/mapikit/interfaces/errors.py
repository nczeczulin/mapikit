from pywintypes import com_error
from win32com.mapi import mapi, mapiutil

__all__ = ['extended_errors']


class extended_errors:
    """A context wrapper that calls GetLastError.

    Args:
        obj (PyIUnknown, optional): An interface object with a GetLastError method.
        check_all (bool, optional): If False, only calls GetLastError
            when code is ``MAPI_E_EXTENDED_ERROR``.
        flags (int, optional): Passed to the GetLastError method.

    Attributes:
        obj (PyIUnknown): An interface object with a GetLastError method.
        flags (int, optional): Passed to the GetLastError method.
    """

    __slots__ = ('obj', 'check_all', 'flags')

    def __init__(self, obj=None, check_all=True, flags=0):
        self.obj = obj
        self.check_all = check_all
        self.flags = flags

    def _annotate_exc(self, exc):
        if type(exc) is com_error and not hasattr(exc, 'exerror'):
            try:
                exc.exerror = None
                exc.args = (exc.hresult, mapiutil.mapiErrorTable.get(exc.hresult, exc.strerror), *exc.args[2:])
                exc.strerror = exc.args[1]
                if self.obj and (self.check_all or exc.hresult == mapi.MAPI_E_EXTENDED_ERROR):
                    try:
                        exc.exerror = self.obj.GetLastError(exc.hresult, self.flags)
                    except AttributeError:
                        pass
                    except com_error as e:
                        if e.hresult == mapi.MAPI_E_BAD_CHARWIDTH:
                            try:
                                exc.exerror = self.obj.GetLastError(
                                    exc.hresult, self.flags ^ mapi.MAPI_UNICODE)
                            except com_error:
                                pass
                    finally:
                        if exc.exerror:
                            exc.args = (*exc.args, exc.exerror)
            except:  # noqa: E722
                self = None  # force decrementing wrapped obj ref count
                raise

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            self._annotate_exc(exc_val)
        finally:
            self.obj = None
