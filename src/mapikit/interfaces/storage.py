import pythoncom
from win32com import storagecon

from .base import IUnknown

__all__ = ['IStream']


class IStream(IUnknown, rawtype=pythoncom.TypeIIDs[pythoncom.IID_IStream]):
    __slots__ = ()

    def read(self, size=-1):
        if size == -1:
            read_size = self.raw.Stat()[2] - self.raw.Seek(0, storagecon.STREAM_SEEK_CUR)
            if read_size <= 0:
                return b''
            return self.raw.Read(read_size)
        else:
            return self.raw.Read(size)

    def write(self, data):
        return self.raw.Write(data)

    def close(self):
        self.release()

    def __len__(self):
        return self.raw.Stat()[2]
