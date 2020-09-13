import os
from datetime import datetime
from contextlib import ContextDecorator
from typing import Tuple

from pywintypes import com_error
from win32com.mapi import mapi, mapitags

from .interfaces import IMAPISession, IMsgStore, extended_errors
from .structures import SRestriction
from .functions import MAPIAdminProfiles, MAPILogonEx

__all__ = ['mapi_initialize', 'logon_temp_profile', 'open_default_store', 'open_pst_file']


class mapi_initialize(ContextDecorator):

    def __init__(self, init: Tuple[int, int] = (mapi.MAPI_INIT_VERSION, mapi.MAPI_MULTITHREAD_NOTIFICATIONS)) -> None:
        self._init = init  # MAPIINIT_0

    def __enter__(self):
        with extended_errors():
            mapi.MAPIInitialize(self._init)

    def __exit__(self, *exc):
        with extended_errors():
            mapi.MAPIUninitialize()


def logon_temp_profile(flags: int = mapi.MAPI_EXTENDED | mapi.MAPI_NO_MAIL) -> IMAPISession:
    name = 'MAPIKit.TempProfile.{:.3f}[{}]'.format(datetime.utcnow().timestamp(), os.getpid())
    with MAPIAdminProfiles() as profile_admin:
        profile_admin.CreateProfile(name, None, 0, 0)
        try:
            return MAPILogonEx(0, name, None, flags)
        finally:
            profile_admin.DeleteProfile(name, 0)


def open_default_store(session: IMAPISession, uiParam: int = 0,
                       flags: int = mapi.MAPI_BEST_ACCESS | mapi.MDB_NO_DIALOG) -> IMsgStore:
    with session.GetMsgStoresTable(0) as table:
        table.SetColumns((mapitags.PR_ENTRYID,), mapi.TBL_BATCH)
        res = SRestriction.res_exist(mapitags.PR_DEFAULT_STORE)
        res &= SRestriction.res_property(mapi.RELOP_EQ, mapitags.PR_DEFAULT_STORE, True)

        if (row := next(table.search(res), None)) is None:
            raise LookupError('PR_DEFAULT_STORE')

        return session.OpenMsgStore(uiParam, row[0][1], None, flags)


def open_pst_file(session: IMAPISession, pst_path: str, uiParam: int = 0,
                  flags: int = mapi.MDB_TEMPORARY | mapi.MAPI_BEST_ACCESS | mapi.MDB_NO_DIALOG) -> IMsgStore:

    admin = session.AdminServices().QueryInterface(mapi.IID_IMsgServiceAdmin2)
    uid = admin.CreateMsgServiceEx('MSUPST MS', None, 0, mapi.SERVICE_NO_RESTART_WARNING)

    try:
        admin.ConfigureMsgService(uid, 0, 0, ((mapitags.PR_PST_PATH_W, pst_path),))

        with session.GetMsgStoresTable(0) as table:
            table.SetColumns((mapitags.PR_ENTRYID,), mapi.TBL_BATCH)
            res = SRestriction.res_exist(mapitags.PR_SERVICE_UID)
            res &= SRestriction.res_property(mapi.RELOP_EQ, mapitags.PR_SERVICE_UID, bytes(uid))

            if (row := next(table.search(res), None)) is None:
                raise LookupError('PR_SERVICE_UID')

        return session.OpenMsgStore(uiParam, row[0][1], None, flags)

    except com_error:
        admin.DeleteMsgService(uid)
        raise
