import pythoncom
from typing import Dict

from ..callwrapper import CallWrapper
from .errors import extended_errors

__all__ = ['IUnknown']


class IUnknown:
    """Interface Wrapper"""

    __slots__ = ('_raw',)

    _raw_typemap: Dict[type, type] = {}

    def __init__(self, obj):
        if self._raw_typemap[type(obj)] is not type(self):
            raise TypeError(f'{type(obj).__name__} object is not a {type(self).__name__} object')
        self._raw = obj

    @property
    def raw(self):
        """object:"""
        self._raise_if_released()
        return self._raw

    @property
    def released(self):
        """bool:"""
        return self._raw is None

    def release(self):
        """Release resource."""
        self._raw = None

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.release()

    def _raise_if_released(self):
        if self._raw is None:
            raise ValueError('operation on released object')

    def __getattr__(self, name):
        func = getattr(self.raw, name)
        return CallWrapper(func, self._result_handler, extended_errors(self)._annotate_exc)

    def _result_handler(self, result):
        try:
            return self._raw_typemap[type(result)](result)
        except KeyError:
            pass

        return result

    @classmethod
    def __init_subclass__(cls, rawtype, **kwargs):
        super().__init_subclass__(**kwargs)
        cls._raw_typemap[rawtype] = cls


IUnknown._raw_typemap[pythoncom.TypeIIDs[pythoncom.IID_IUnknown]] = IUnknown
