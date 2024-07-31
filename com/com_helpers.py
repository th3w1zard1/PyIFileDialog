from __future__ import annotations

import sys

from ctypes import POINTER, byref, c_uint, windll
from typing import TYPE_CHECKING, Generic, TypeVar

from com.com_types import GUID
from com.interfaces import IUnknown
from hresult import S_FALSE, S_OK, decode_hresult

if TYPE_CHECKING:
    from ctypes import _CData, _Pointer as PointerType
    from types import TracebackType

    T = TypeVar("T", bound=_CData)
if not TYPE_CHECKING:
    PointerType = POINTER(c_uint).__class__
    T = TypeVar("T")

class COMInitializeContext(Generic[T]):
    def __init__(self, clsid: GUID | None = None, interface: type[T | IUnknown] | None = None):
        self.clsid: GUID | None = clsid
        self.interface: type[T | IUnknown] | None = interface
        self._should_uninitialize: bool = False

    def __enter__(self) -> T | IUnknown | None:
        print("windll.ole32.CoInitialize()")
        hr = windll.ole32.CoInitialize(None)
        if hr == S_FALSE:
            print("COM library already initialized.", file=sys.stderr)
        elif hr != S_OK:
            raise OSError(f"CoInitialize failed! {hr}")

        self._should_uninitialize = True
        if self.interface is not None and self.clsid is not None:
            p: PointerType[T | IUnknown] = POINTER(self.interface)()
            iid: GUID | None = getattr(self.interface, "_iid_", None)
            if iid is None or not isinstance(iid, GUID.guid_ducktypes()):
                raise OSError("Incorrect interface definition")
            hr = windll.ole32.CoCreateInstance(byref(self.clsid), None, 1, byref(iid), byref(p))
            if hr != S_OK:
                raise OSError(f"CoCreateInstance failed! {decode_hresult(hr)}")
            return p.contents
        return None

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_val: BaseException | None,
        exc_tb: TracebackType | None,
    ):
        if self._should_uninitialize:
            print("windll.ole32.CoUninitialize()")
            windll.ole32.CoUninitialize()
