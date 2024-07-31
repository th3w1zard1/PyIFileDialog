from __future__ import annotations

from ctypes import (
    POINTER,
    WINFUNCTYPE,
    Structure,
    byref,
    c_char_p,
    c_int,
    c_uint,
    c_ulong,
    c_void_p,
    windll,
)
from ctypes.wintypes import HWND, LPCWSTR, LPWSTR
from typing import TYPE_CHECKING, ClassVar, Sequence

from com_types import GUID
from hresult import HRESULT
from iunknown import COMBase, IUnknown, IUnknownVTable

if TYPE_CHECKING:
    import os

    from ctypes import _CData, _FuncPointer, _NamedFuncPointer, _Pointer

# COM HRESULT definitions
SFGAO_FILESYSTEM = 0x40000000
SFGAO_FOLDER = 0x20000000
SIGDN_FILESYSPATH = c_ulong(0x80058000)
SIGDN_NORMALDISPLAY = c_ulong(0x00000000)
SICHINT_CANONICAL = 0x10000000
IID_IFileDialog = GUID("{4292C689-9298-4A37-B8F0-03835B7B6EE6}")
IID_IShellItem = GUID("{43826D1E-E718-42EE-BC55-A1E261C37BFE}")
IID_IShellItemArray = GUID("{B63EA76D-1F85-456F-A19C-48159EFA858B}")
IID_IShellItemFilter = GUID("{2659B475-EEB8-48B7-8F07-B378810F48CF}")
IID_IModalWindow = GUID("{B4DB1657-70D7-485E-8E3E-6FCB5A5C1802}")
IID_IFileSaveDialog = GUID("{84BCCD23-5FDE-4CDB-AEAA-AE8CD7A4D575}")
IID_IFileOpenDialog = GUID("{D57C7288-D4AD-4768-BE02-9D969532D960}")
IID_IFileDialogEvents = GUID("{973510DB-7D7F-452B-8975-74A85828D354}")
CLSID_FileDialog = GUID("{3D9C8F03-50D4-4E40-BB11-70E74D3F10F3}")
CLSID_FileOpenDialog = GUID("{DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7}")
CLSID_FileSaveDialog = GUID("{C0B4E2F3-BA21-4773-8DBA-335EC946EB8B}")
CLSID_IShellLibrary = GUID("{D9B3211D-E57F-4426-AAEF-30A806ADD397}")


class COMDLG_FILTERSPEC(Structure):  # noqa: N801
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]] = [
        ("pszName", LPCWSTR),
        ("pszSpec", LPCWSTR),
    ]


# COM interface and vtable definitions
class IModalWindow(COMBase):
    _iid_: GUID = IID_IModalWindow
    _methods_: ClassVar[list[tuple[str, Sequence[type | str]]]] = [
        ("Show", (HRESULT, "SELF", HWND)),
    ]
    _base_ = IUnknown.VTable


class IShellItem(COMBase):
    _iid_: GUID = IID_IShellItem
    _methods_: ClassVar[list[tuple[str, Sequence[type | str]]]] = [
        ("BindToHandler", (HRESULT, "SELF", POINTER(IUnknown), POINTER(GUID), POINTER(GUID), POINTER(c_void_p))),
        ("GetParent", (HRESULT, "SELF", "POINTER(SELF)")),
        ("GetDisplayName", (HRESULT, "SELF", c_int, POINTER(LPWSTR))),
        ("GetAttributes", (HRESULT, "SELF", c_ulong, POINTER(c_ulong))),
        ("Compare", (HRESULT, "SELF", "SELF", c_int, POINTER(c_int))),
    ]
    _base_ = IUnknown.VTable

    @classmethod
    def from_path(cls, path: os.PathLike | str) -> _Pointer[IShellItem]:
        item = POINTER(IShellItem)()
        hr = SHCreateItemFromParsingName(path, None, byref(IID_IShellItem), byref(item))
        if hr != 0:
            raise HRESULT(hr).exception("SHCreateItemFromParsingName failed!")
        return item


class IShellItemArray(COMBase):
    _iid_: GUID = IID_IShellItemArray
    _methods_: ClassVar[list[tuple[str, Sequence[type | str]]]] = [
        ("GetCount", (HRESULT, "SELF", POINTER(c_uint))),
        ("GetItemAt", (HRESULT, "SELF", c_uint, "POINTER(SELF)")),
        ("EnumItems", (HRESULT, "SELF", POINTER(POINTER(IUnknown)))),
    ]
    _base_ = IUnknown.VTable


class IEnumShellItems(COMBase):
    _iid_ = GUID("{70629033-E363-4A28-A567-0DB78006E6D7}")
    _methods_: ClassVar[list[tuple[str, Sequence[type | str]]]] = [
        ("Next", (HRESULT, "SELF", c_ulong, POINTER(POINTER(IShellItem)), POINTER(c_ulong))),
        ("Skip", (HRESULT, "SELF", c_ulong)),
        ("Reset", (HRESULT, "SELF")),
        ("Clone", (HRESULT, "SELF", "POINTER(SELF)")),
    ]
    _base_ = IUnknown.VTable


class IShellItemFilter(COMBase):
    _iid_: GUID = IID_IShellItemFilter
    lpVtbl: _Pointer[IShellItemFilterVTable]


class IShellItemFilterVTable(IUnknownVTable):
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]] = [
        *IUnknownVTable._fields_,
        ("IncludeItem", WINFUNCTYPE(HRESULT, POINTER(IShellItemFilter), POINTER(IShellItem))),
        ("GetEnumFlagsForItem", WINFUNCTYPE( HRESULT, POINTER(IShellItemFilter), POINTER(IShellItem), POINTER(c_ulong))),
    ]


IShellItemFilter._fields_ = [("lpVtbl", POINTER(IShellItemFilterVTable))]

class IFileDialog(COMBase):
    _iid_: GUID = IID_IFileDialog
    _base_ = IModalWindow.VTable
    _methods_: ClassVar[list[tuple[str, Sequence[type | str]]]]
class IFileDialogEvents(COMBase):
    _iid_: GUID = IID_IFileDialogEvents
    _methods_: ClassVar[list[tuple[str, Sequence[type | str]]]] = [
        ("OnFileOk", (HRESULT, "SELF", POINTER(IFileDialog))),
        ("OnFolderChanging", (HRESULT, "SELF", POINTER(IFileDialog), POINTER(IShellItem))),
        ("OnFolderChange", (HRESULT, "SELF", POINTER(IFileDialog))),
        ("OnSelectionChange", (HRESULT, "SELF", POINTER(IFileDialog))),
        ("OnShareViolation", (HRESULT, "SELF", POINTER(IFileDialog), POINTER(IShellItem), POINTER(c_int))),
        ("OnTypeChange", (HRESULT, "SELF", POINTER(IFileDialog))),
        ("OnOverwrite", (HRESULT, "SELF", POINTER(IFileDialog), POINTER(IShellItem), POINTER(c_int))),
    ]
    _base_ = IUnknown.VTable
IFileDialog._methods_ = [
    ("SetFileTypes", (HRESULT, "SELF", c_uint, POINTER(COMDLG_FILTERSPEC))),
    ("SetFileTypeIndex", (HRESULT, "SELF", c_uint)),
    ("GetFileTypeIndex", (HRESULT, "SELF", POINTER(c_uint))),
    ("Advise", (HRESULT, "SELF", POINTER(IFileDialogEvents), POINTER(c_uint))),
    ("Unadvise", (HRESULT, "SELF", c_uint)),
    ("SetOptions", (HRESULT, "SELF", c_uint)),
    ("GetOptions", (HRESULT, "SELF", POINTER(c_uint))),
    ("SetDefaultFolder", (HRESULT, "SELF", POINTER(IShellItem))),
    ("SetFolder", (HRESULT, "SELF", POINTER(IShellItem))),
    ("GetFolder", (HRESULT, "SELF", POINTER(POINTER(IShellItem)))),
    ("GetCurrentSelection", (HRESULT, "SELF", POINTER(POINTER(IShellItem)))),
    ("SetFileName", (HRESULT, "SELF", LPCWSTR)),
    ("GetFileName", (HRESULT, "SELF", POINTER(LPWSTR))),
    ("SetTitle", (HRESULT, "SELF", LPCWSTR)),
    ("SetOkButtonLabel", (HRESULT, "SELF", LPCWSTR)),
    ("SetFileNameLabel", (HRESULT, "SELF", LPCWSTR)),
    ("GetResult", (HRESULT, "SELF", POINTER(POINTER(IShellItem)))),
    ("AddPlace", (HRESULT, "SELF", POINTER(IShellItem), c_int)),
    ("SetDefaultExtension", (HRESULT, "SELF", LPCWSTR)),
    ("Close", (HRESULT, "SELF", HRESULT)),
    ("SetClientGuid", (HRESULT, "SELF", POINTER(GUID))),
    ("ClearClientData", (HRESULT, "SELF")),
    ("SetFilter", (HRESULT, "SELF", POINTER(IShellItemFilter))),
]

class IFileOpenDialog(COMBase):
    _iid_: GUID = GUID("{D57C7288-D4AD-4768-BE02-9D969532D960}")
    _methods_: ClassVar[list[tuple[str, Sequence[type | str]]]] = [
        ("GetResults", (HRESULT, "SELF", POINTER(POINTER(IShellItemArray)))),
        ("GetSelectedItems", (HRESULT, "SELF", POINTER(POINTER(IShellItemArray)))),
    ]
    _base_ = IFileDialog.VTable

class IFileSaveDialog(COMBase):
    _iid_: GUID = GUID("{84BCCD23-5FDE-4CDB-AEA4-AF64B83D78AB}")
    _methods_: ClassVar[list[tuple[str, Sequence[type | str]]]] = [
        ("SetSaveAsItem", (HRESULT, "SELF", POINTER(IShellItem))),
        ("SetProperties", (HRESULT, "SELF", POINTER(IUnknown))),
        ("SetCollectedProperties", (HRESULT, "SELF", POINTER(IUnknown), c_int)),
        ("GetProperties", (HRESULT, "SELF", POINTER(POINTER(IUnknown)))),
        ("ApplyProperties", (HRESULT, "SELF", POINTER(IShellItem), POINTER(IUnknown), HWND, POINTER(IUnknown))),
    ]
    _base_ = IFileDialog.VTable


# Constants
CLSCTX_INPROC_SERVER = 1
FOS_OVERWRITEPROMPT = 0x2
FOS_STRICTFILETYPES = 0x4
FOS_NOCHANGEDIR = 0x8
FOS_PICKFOLDERS = 0x20
FOS_FORCEFILESYSTEM = 0x40
FOS_ALLNONSTORAGEITEMS = 0x80
FOS_NOVALIDATE = 0x100
FOS_ALLOWMULTISELECT = 0x200
FOS_PATHMUSTEXIST = 0x800
FOS_FILEMUSTEXIST = 0x1000
FOS_CREATEPROMPT = 0x2000
FOS_SHAREAWARE = 0x4000
FOS_NOREADONLYRETURN = 0x8000
FOS_NOTESTFILECREATE = 0x10000
FOS_HIDEMRUPLACES = 0x20000
FOS_HIDEPINNEDPLACES = 0x40000
FOS_NODEREFERENCELINKS = 0x100000
FOS_DONTADDTORECENT = 0x2000000
FOS_FORCESHOWHIDDEN = 0x10000000
FOS_DEFAULTNOMINIMODE = 0x20000000
FOS_FORCEPREVIEWPANEON = 0x40000000


# Enums
class SIGDN(c_int):
    SIGDN_NORMALDISPLAY = 0x00000000
    SIGDN_PARENTRELATIVEPARSING = 0x80018001
    SIGDN_PARENTRELATIVEFORADDRESSBAR = 0x8001C001
    SIGDN_DESKTOPABSOLUTEPARSING = 0x80028000
    SIGDN_PARENTRELATIVEEDITING = 0x80031001
    SIGDN_DESKTOPABSOLUTEEDITING = 0x8004C000
    SIGDN_FILESYSPATH = 0x80058000
    SIGDN_URL = 0x80068000


class FDAP(c_int):
    FDAP_BOTTOM = 0x00000000
    FDAP_TOP = 0x00000001


class SIATTRIBFLAGS(c_int):
    SIATTRIBFLAGS_AND = 0x1
    SIATTRIBFLAGS_OR = 0x2
    SIATTRIBFLAGS_APPCOMPAT = 0x3
    SIATTRIBFLAGS_MASK = 0x3


# Function prototypes
SHCreateItemFromParsingName: _NamedFuncPointer = windll.shell32.SHCreateItemFromParsingName
SHCreateItemFromParsingName.argtypes = [
    LPCWSTR,
    c_void_p,
    POINTER(GUID),
    POINTER(POINTER(IShellItem)),
]
SHCreateItemFromParsingName.restype = HRESULT


class COMFunctionPointers:
    def __init__(self):
        self.hOle32: _FuncPointer
        self.hShell32: _FuncPointer
        self.pCoInitialize: _FuncPointer
        self.pCoUninitialize: _FuncPointer
        self.pCoCreateInstance: _FuncPointer
        self.pCoTaskMemFree: _FuncPointer
        self.pSHCreateItemFromParsingName: _FuncPointer

    @staticmethod
    def load_library(dll_name: str) -> _FuncPointer:
        windll.kernel32.LoadLibraryW.argtypes = [LPCWSTR]
        windll.kernel32.LoadLibraryW.restype = c_void_p
        handle = windll.kernel32.LoadLibraryW(dll_name)
        if not handle:
            raise ValueError(f"Unable to load library: {dll_name}")
        return handle

    @staticmethod
    def resolve_function(handle: _FuncPointer, func: bytes, func_type: type[_FuncPointer]) -> _FuncPointer:
        windll.kernel32.GetProcAddress.argtypes = [c_void_p, c_char_p]
        windll.kernel32.GetProcAddress.restype = c_void_p
        address = windll.kernel32.GetProcAddress(handle, func)
        assert address is not None
        return func_type(address)
