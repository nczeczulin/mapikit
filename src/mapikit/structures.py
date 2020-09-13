import os
from typing import NamedTuple

from win32com.mapi import mapi, mapiutil

__all__ = ['SRestriction']


class SRestriction(NamedTuple):
    rt: int
    res: tuple

    _RES_TABLE = dict((v, k) for (k, v) in mapi.__dict__.items() if k.startswith('RES_'))  # type: ignore
    _RELOP_TABLE = dict((v, k) for (k, v) in mapi.__dict__.items() if k.startswith('RELOP_'))  # type: ignore
    _FL_TABLE = dict((v, k) for (k, v) in mapi.__dict__.items() if k.startswith('FL_'))  # type: ignore
    _BMR_TABLE = dict((v, k) for (k, v) in mapi.__dict__.items() if k.startswith('BMR_'))  # type: ignore

    def __add__(self, other):
        return NotImplemented

    def __and__(self, other):
        if not isinstance(other, SRestriction):
            return NotImplemented
        return self._logical_res(mapi.RES_AND, self, other)

    def __or__(self, other):
        if not isinstance(other, SRestriction):
            return NotImplemented
        return self._logical_res(mapi.RES_OR, self, other)

    @classmethod
    def _logical_res(cls, rt, *res):

        modified = list()
        for r in res:
            modified.extend(r.res) if rt == r.rt else modified.append(r)

        return cls(rt, tuple(modified))

    @classmethod
    def res_not(cls, res):
        return cls(mapi.RES_NOT, (res,))

    @classmethod
    def res_property(cls, relop, proptag, value):
        return cls(mapi.RES_PROPERTY, (relop, proptag, (proptag, value)))

    @classmethod
    def res_exist(cls, proptag):
        return cls(mapi.RES_EXIST, (proptag,))

    @classmethod
    def res_content(cls, fuzzylevel, proptag, value):
        return cls(mapi.RES_CONTENT, (fuzzylevel, proptag, (proptag, value)))

    @classmethod
    def res_bitmask(cls, relbmr, proptag, mask):
        return cls(mapi.RES_BITMASK, (relbmr, proptag, mask))

    def pformat(self, indent=4, linesep=os.linesep):

        def _pformat(res, depth=0):
            output = list()
            output.append(f'{" "*indent*depth}{type(res).__name__}: {self._RES_TABLE[res.rt]}')
            if res.rt in (mapi.RES_AND, mapi.RES_OR, mapi.RES_NOT):
                output.append(f'{linesep}')
                for r in res.res:
                    output.append(_pformat(r, depth+1))
            else:
                if res.rt == mapi.RES_PROPERTY:
                    relop = self._RELOP_TABLE[res.res[0]]
                    proptag = mapiutil.GetPropTagName(res.res[1])
                    value = res.res[2][1]
                    output.append(f' {relop=} {proptag=} {value=}')
                elif res.rt == mapi.RES_EXIST:
                    proptag = mapiutil.GetPropTagName(res.res[0])
                    output.append(f' {proptag=}')
                elif res.rt == mapi.RES_CONTENT:
                    fuzzylevel = self._FL_TABLE[res.res[0]]
                    proptag = mapiutil.GetPropTagName(res.res[1])
                    value = res.res[2][1]
                    output.append(f' {fuzzylevel=} {proptag=} {value=}')
                elif res.rt == mapi.RES_BITMASK:
                    relbmr = self._BMR_TABLE[res.res[0]]
                    proptag = mapiutil.GetPropTagName(res.res[1])
                    mask = res.res[2]
                    output.append(f' {relbmr=} {proptag=} {mask=}')

                output.append(f'{linesep}')
            return ''.join(output)

        return _pformat(self, 0)

    def pprint(self, indent=4):
        print(self.pformat(indent))
