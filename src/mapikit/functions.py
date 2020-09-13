from win32com.mapi import mapi

from .interfaces import IMAPISession, IProfAdmin, extended_errors

__all__ = ['MAPILogonEx', 'MAPIAdminProfiles']


def MAPILogonEx(uiParam: int, profileName: str, password: str = None,
                flags: int = mapi.MAPI_EXTENDED | mapi.MAPI_NEW_SESSION | mapi.MAPI_EXPLICIT_PROFILE) -> IMAPISession:

    with extended_errors():
        return IMAPISession(mapi.MAPILogonEx(uiParam, profileName, password, flags))


def MAPIAdminProfiles(flags: int = 0) -> IProfAdmin:
    with extended_errors():
        return IProfAdmin(mapi.MAPIAdminProfiles(flags))
