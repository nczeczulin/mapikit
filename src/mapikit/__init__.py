"""
mapikit: Extended MAPI Made Easier
==================================
"""

from pkg_resources import get_distribution, DistributionNotFound
try:
    __version__ = get_distribution(__name__).version
except DistributionNotFound:
    # package is not installed
    pass
finally:
    del get_distribution, DistributionNotFound

from win32com.mapi import mapi, mapitags, mapiutil

from . import interfaces  # noqa: F401
from .macros import *  # noqa: F401, F403
from .functions import *  # noqa: F401, F403
from .utils import *  # noqa F401, F403

mapitags.PR_PROFILE_USER_SMTP_EMAIL_ADDRESS = mapitags.PROP_TAG(mapitags.PT_TSTRING, 0x6641)
mapitags.PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W = mapitags.PROP_TAG(mapitags.PT_UNICODE, 0x6641)
mapitags.PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_A = mapitags.PROP_TAG(mapitags.PT_STRING8, 0x6641)

mapitags.PR_INTERNET_MESSAGE_ID = mapitags.PROP_TAG(mapitags.PT_TSTRING, 0x1035)
mapitags.PR_INTERNET_MESSAGE_ID_W = mapitags.PROP_TAG(mapitags.PT_UNICODE, 0x1035)
mapitags.PR_INTERNET_MESSAGE_ID_A = mapitags.PROP_TAG(mapitags.PT_STRING8, 0x1035)

# Ensure mapiutil tables are initialized
mapiutil.GetPropTagName(mapitags.PR_BODY)
mapiutil.GetMapiTypeName(mapitags.PT_UNICODE)
mapiutil.GetScodeString(mapi.MAPI_E_NOT_FOUND)

mapiutil.mapiErrorTable[-2130696378] = 'MAIL_E_NAMENOTFOUND'
mapiutil.mapiErrorTable[-2147219456] = 'SYNC_E_OBJECT_DELETED'
mapiutil.mapiErrorTable[-2147219455] = 'SYNC_E_IGNORE'
mapiutil.mapiErrorTable[-2147219454] = 'SYNC_E_CONFLICT'
mapiutil.mapiErrorTable[-2147219453] = 'SYNC_E_NO_PARENT'
mapiutil.mapiErrorTable[-2147219452] = 'SYNC_E_CYCLE'
mapiutil.mapiErrorTable[-2147219451] = 'SYNC_E_UNSYNCHRONIZED'
mapiutil.mapiErrorTable[264224] = 'SYNC_W_PROGRESS'
mapiutil.mapiErrorTable[264225] = 'SYNC_W_CLIENT_CHANGE_NEWER'


mapi.FOLDER_IPM_SUBTREE_VALID = 0x00000001
mapi.FOLDER_IPM_INBOX_VALID = 0x00000002
mapi.FOLDER_IPM_OUTBOX_VALID = 0x00000004
mapi.FOLDER_IPM_WASTEBASKET_VALID = 0x00000008
mapi.FOLDER_IPM_SENTMAIL_VALID = 0x00000010
mapi.FOLDER_VIEWS_VALID = 0x00000020
mapi.FOLDER_COMMON_VIEWS_VALID = 0x00000040
mapi.FOLDER_FINDER_VALID = 0x00000080

mapi.SERVICE_NO_RESTART_WARNING = 0x00000080
