from win32com.mapi import mapitags

from mapikit import macros


def test_PROP_TYPE():
    assert macros.PROP_TYPE(mapitags.PR_MESSAGE_CLASS_W) == mapitags.PT_UNICODE


def test_PROP_ID():
    assert macros.PROP_ID(mapitags.PR_MESSAGE_CLASS_W) == 26


def test_PROP_TYPE_AND_ID():
    assert macros.PROP_TYPE_AND_ID(mapitags.PR_MESSAGE_CLASS_W) == (mapitags.PT_UNICODE, 26)


def test_PROP_TAG():
    assert macros.PROP_TAG(mapitags.PT_UNICODE, 26) == mapitags.PR_MESSAGE_CLASS_W


def test_CHANGE_PROP_TYPE():
    assert macros.CHANGE_PROP_TYPE(mapitags.PR_MESSAGE_CLASS_W, mapitags.PT_STRING8) == mapitags.PR_MESSAGE_CLASS_A
