import codecs
import collections

import pythoncom
from pywintypes import com_error
from win32com.mapi import mapi, mapitags

from .base import IUnknown
from .errors import extended_errors
from ..macros import PROP_TYPE
from ..structures import SRestriction

__all__ = [
    'IMAPISession',
    'IMAPIProp',
    'IProfSect',
    'IMessage',
    'IMsgStore',
    'IAttach',
    'IMailUser',
    'IAddrBook',
    'IMAPIContainer',
    'IMAPIFolder',
    'IDistList',
    'IMAPITable',
    'IProfAdmin',
    'IMsgServiceAdmin',
    'IMsgServiceAdmin2',
]


class IMAPISession(IUnknown, rawtype=pythoncom.TypeIIDs[mapi.IID_IMAPISession]):
    __slots__ = ()

    def release(self):
        if not self.released:
            self.Logoff(0, 0, 0)
        super().release()


class IMAPIProp(IUnknown, rawtype=pythoncom.TypeIIDs[mapi.IID_IMAPIProp]):
    __slots__ = ()
    _unicode_stream_reader = codecs.getreader('utf-16le')
    _unicode_stream_writer = codecs.getwriter('utf-16le')

    def __getitem__(self, key):
        try:
            with extended_errors(self):
                return mapi.HrGetOneProp(self.raw, key)[1]
        except com_error as e:
            if e.hresult == mapi.MAPI_E_NOT_FOUND:
                raise KeyError(key) from None
            elif e.hresult == mapi.MAPI_E_NOT_ENOUGH_MEMORY:
                if PROP_TYPE(key) in (mapitags.PT_BINARY, mapitags.PT_UNICODE, mapitags.PT_STRING8):
                    try:
                        with self.OpenProperty(key, pythoncom.IID_IStream, 0, 0) as s:
                            if PROP_TYPE(key) == mapitags.PT_UNICODE:
                                with self._unicode_stream_reader(s) as r:
                                    return r.read()
                            else:
                                return s.read()
                    except Exception as e:
                        raise e from None
            raise

    def __setitem__(self, key, value):
        try:
            with extended_errors(self):
                return mapi.HrSetOneProp(self.raw, (key, value))
        except com_error as e:
            if e.hresult == mapi.MAPI_E_NOT_FOUND:
                raise KeyError(key) from None
            elif e.hresult == mapi.MAPI_E_NOT_ENOUGH_MEMORY:
                if PROP_TYPE(key) in (mapitags.PT_BINARY, mapitags.PT_UNICODE, mapitags.PT_STRING8):
                    try:
                        with self.OpenProperty(
                                key, pythoncom.IID_IStream, 0, mapi.MAPI_CREATE | mapi.MAPI_MODIFY) as s:
                            if PROP_TYPE(key) == mapitags.PT_UNICODE:
                                with self._unicode_stream_writer('utf-16le')(s) as w:
                                    w.write(value)
                            else:
                                s.write(value)
                    except Exception as e:
                        raise e from None
                    else:
                        return
            raise

    def __delitem__(self, key):
        with extended_errors(self):
            ret, probs = self.DeleteProps((key,), True)
            if probs and probs[0][2] == mapi.MAPI_E_NOT_FOUND:  # ensure consistent missing key behavior
                raise KeyError(key)

    def __contains__(self, key):
        try:
            with extended_errors(self):
                mapi.HrGetOneProp(self.raw, key)
                return True
        except com_error as e:
            if e.hresult == mapi.MAPI_E_NOT_FOUND:
                return False
            elif e.hresult == mapi.MAPI_E_NOT_ENOUGH_MEMORY:
                return True
            else:
                raise

    def __iter__(self):
        raise TypeError(f'{repr(type(self).__name__)} object is not iterable')

    def get(self, proptag, default=None):
        try:
            return self[proptag]
        except KeyError:
            return default


class IProfSect(IMAPIProp, rawtype=pythoncom.TypeIIDs[mapi.IID_IProfSect]):
    __slots__ = ()


class IMessage(IMAPIProp, rawtype=pythoncom.TypeIIDs[mapi.IID_IMessage]):
    __slots__ = ()


class IMsgStore(IMAPIProp, rawtype=pythoncom.TypeIIDs[mapi.IID_IMsgStore]):
    __slots__ = ()

    def release(self):
        if not self.released:
            pass  # placeholder for when StoreLogoff is added to pywin
        super().release()


class IAttach(IMAPIProp, rawtype=pythoncom.TypeIIDs[mapi.IID_IAttachment]):
    __slots__ = ()


class IMailUser(IMAPIProp, rawtype=pythoncom.TypeIIDs[mapi.IID_IMailUser]):
    __slots__ = ()


class IAddrBook(IMAPIProp, rawtype=pythoncom.TypeIIDs[mapi.IID_IAddrBook]):
    __slots__ = ()


class IMAPIContainer(IMAPIProp, rawtype=pythoncom.TypeIIDs[mapi.IID_IMAPIContainer]):
    __slots__ = ()


class IMAPIFolder(IMAPIContainer, rawtype=pythoncom.TypeIIDs[mapi.IID_IMAPIFolder]):
    __slots__ = ()

    def folders(self):
        with self.GetHierarchyTable(mapi.MAPI_UNICODE) as table:
            yield from table

    def contents(self):
        with self.GetContentsTable(mapi.MAPI_UNICODE) as table:
            yield from table


class IDistList(IMAPIContainer, rawtype=pythoncom.TypeIIDs[mapi.IID_IDistList]):
    __slots__ = ()


class IMAPITable(IUnknown, collections.abc.Iterable, rawtype=pythoncom.TypeIIDs[mapi.IID_IMAPITable]):
    __slots__ = ('_iter_prefetch',)

    @property
    def iter_prefetch(self):
        """int: Get value"""
        return self._iter_prefetch

    @iter_prefetch.setter
    def iter_prefetch(self, value):
        """int: Set value"""
        self._iter_prefetch = value

    def __init__(self, obj, iter_prefetch=1000):
        super().__init__(obj)
        self._iter_prefetch = iter_prefetch

    def __iter__(self):
        while True:
            rows = self.QueryRows(self.iter_prefetch, 0)
            if not rows:
                break
            for row in rows:
                yield row

    def search(self, res, origin=mapi.BOOKMARK_BEGINNING, backward=False):
        bookmark = origin
        while True:
            try:
                self.FindRow(res, bookmark, mapi.DIR_BACKWARD if backward else 0)
                bookmark = mapi.BOOKMARK_CURRENT
                # TODO: Add MAPI_E_BUSY handling for QueryRows
                rows = self.QueryRows(1, -1 if backward else 0)
                yield rows[0]
            except com_error as e:
                if e.hresult == mapi.MAPI_E_NOT_FOUND:
                    return
                raise


class IProfAdmin(IUnknown, collections.abc.Container, collections.abc.Iterable,
                 rawtype=pythoncom.TypeIIDs[mapi.IID_IProfAdmin]):
    __slots__ = ()

    @property
    def default(self):
        """str: Get the default profile name.

        Raises:
            KeyError: The default profile name was not found.
        """
        with self.GetProfileTable(0) as table:
            table.SetColumns((mapitags.PR_DISPLAY_NAME_A,), mapi.TBL_BATCH)
            res = SRestriction.res_exist(mapitags.PR_DEFAULT_PROFILE)
            res &= SRestriction.res_property(mapi.RELOP_EQ, mapitags.PR_DEFAULT_PROFILE, True)

            for row in table.search(res):
                return row[0][1].decode('mbcs')
            else:
                raise LookupError('PR_DEFAULT_PROFILE') from None

    @default.setter
    def default(self, name):
        self.SetDefaultProfile(name, 0)

    def __contains__(self, name):
        try:
            with self.GetProfileTable(0) as table:
                table.FindRow(
                    SRestriction.res_property(mapi.RELOP_EQ, mapitags.PR_DISPLAY_NAME_A, name),
                    mapi.BOOKMARK_BEGINNING, 0
                )
        except com_error as e:
            if e.hresult == mapi.MAPI_E_NOT_FOUND:
                return False
            raise
        else:
            return True

    def __iter__(self):
        with self.GetProfileTable(0) as table:
            table.SetColumns((mapitags.PR_DISPLAY_NAME_A,), mapi.TBL_BATCH)
            yield from (t[0][1].decode('mbcs') for t in table)


class IMsgServiceAdmin(IUnknown, rawtype=pythoncom.TypeIIDs[mapi.IID_IMsgServiceAdmin]):
    __slots__ = ()


class IMsgServiceAdmin2(IMsgServiceAdmin, rawtype=pythoncom.TypeIIDs[mapi.IID_IMsgServiceAdmin2]):
    __slots__ = ()
