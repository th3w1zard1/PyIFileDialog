from __future__ import annotations

from ctypes import (
    CFUNCTYPE,
    POINTER,
    WINFUNCTYPE,
    Structure,
    byref,
    c_int,
    c_uint,
    c_void_p,
    c_wchar_p,
    pointer,
    windll,
    wintypes,
)
from typing import TYPE_CHECKING, Any, ClassVar, Sequence

from com_types import GUID
from hresult import HRESULT

if TYPE_CHECKING:
    from ctypes import _FuncPointer, _Pointer

REFIID: type[_Pointer[GUID]] = POINTER(GUID)
REFGUID: type[_Pointer[GUID]] = POINTER(GUID)

# Define ULONG and LONG
ULONG = wintypes.DWORD
LONG = wintypes.LONG

# Define basic Windows types
LPVOID = c_void_p
HWND = c_void_p
LPWSTR = c_wchar_p
LPCWSTR = c_wchar_p
UINT = c_uint
BOOL = c_int
DWORD = wintypes.DWORD
HMODULE = c_void_p

# Define HRESULT values
S_OK = HRESULT(0)
E_NOINTERFACE = HRESULT(0x80004002)

# Define CLSCTX_INPROC_SERVER
CLSCTX_INPROC_SERVER = 0x1


# Define SUCCEEDED and FAILED macros
def SUCCEEDED(hr):
    return hr >= 0


def FAILED(hr):
    return hr < 0



# Define InterlockedIncrement manually
def InterlockedIncrement(var_ptr):
    windll.kernel32.EnterCriticalSection(byref(var_ptr))
    var_ptr.contents.value += 1
    result = var_ptr.contents.value
    windll.kernel32.LeaveCriticalSection(byref(var_ptr))
    return result

# Define InterlockedDecrement manually
def InterlockedDecrement(var_ptr):
    windll.kernel32.EnterCriticalSection(byref(var_ptr))
    var_ptr.contents.value -= 1
    result = var_ptr.contents.value
    windll.kernel32.LeaveCriticalSection(byref(var_ptr))
    return result


class COMBase(Structure):
    _fields_: Sequence[tuple[str, Any]] = []
    lpVtbl: _Pointer[Structure]
    _methods_: ClassVar[list[tuple[str, Sequence[type | str]]]] = []
    _base_: type = Structure

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        #if 'lpVtbl' in self._fields_:
        #    setattr(self, 'lpVtbl', (self.lpVtbl._type_)())

    @classmethod
    def __init_subclass__(cls, **kwargs):
        super().__init_subclass__(**kwargs)
        if "_methods_" in cls.__dict__ and cls._methods_:
            cls._finalize_methods()

    @classmethod
    def _finalize_methods(cls):
        vtable_fields = []
        for method_name, method_types in cls._methods_:
            updated_types = tuple(
                POINTER(cls) if t == "SELF" else
                POINTER(POINTER(cls)) if t == "POINTER(SELF)" else
                t
                for t in method_types
            )
            func_type: type[_FuncPointer] = CFUNCTYPE(*updated_types)
            vtable_fields.append((method_name, c_void_p))

        # Define the VTable as a subclass of Structure
        vtable_class = type(f"{cls.__name__}VTable", (Structure,), {"_fields_": vtable_fields})
        cls.VTable = vtable_class
        cls._fields_ = [("lpVtbl", POINTER(vtable_class))]

        # Set the finalized function pointers in the VTable
        for method_name, method_types in cls._methods_:
            updated_types = tuple(
                POINTER(cls) if t == "SELF" else
                POINTER(POINTER(cls)) if t == "POINTER(SELF)" else
                t
                for t in method_types
            )
            setattr(cls.VTable, method_name, CFUNCTYPE(*updated_types))

    def __setattr__(self, name, value):
        if name == "_methods_" and value:
            self.__class__._methods_ = value
            self.__class__._finalize_methods()
        super().__setattr__(name, value)

    def call(self, method_name: str, *args):
        func_type: Sequence[type] | None = None
        for name, types in self._methods_:
            if name == method_name:
                func_type = types
                break
        else:
            raise AttributeError(f"Method '{method_name}' not found in class {self.__class__.__name__}")

        new_cfunctype: type[_FuncPointer] = WINFUNCTYPE(func_type[0], *([POINTER(type(self)), *func_type[1:]]))
        func_ptr: _FuncPointer = getattr(self.lpVtbl.contents, method_name)
        new_func_ptr: _FuncPointer = new_cfunctype(func_ptr)

        return new_func_ptr(self, *args)

    def query_interface(self, interface_id: GUID):
        assert isinstance(self, IUnknown)
        p_interface: _Pointer[IUnknown] = pointer(self.__class__())
        hr = self.call("QueryInterface", byref(self), byref(interface_id), byref(p_interface))
        if hr != S_OK:
            raise HRESULT(hr).exception("QueryInterface call failed!")
        return p_interface

    def add_ref(self):
        return self.call("AddRef")

    def release(self):
        return self.call("Release")

# Define IUnknown interface
IID_IUnknown = GUID("{00000000-0000-0000-C000-000000000046}")

class IUnknown(COMBase):
    _iid_: GUID = IID_IUnknown
    _methods_: ClassVar[list[tuple[str, Sequence[type | str]]]] = [
        ("QueryInterface", (HRESULT, "SELF", REFIID, POINTER(LPVOID))),
        ("AddRef", (ULONG, "SELF")),
        ("Release", (ULONG, "SELF")),
    ]

LPUNKNOWN: type[_Pointer[IUnknown]] = POINTER(IUnknown)
