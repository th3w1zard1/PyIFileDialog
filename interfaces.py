from __future__ import annotations

from ctypes import (
    POINTER,
    WINFUNCTYPE,
    Structure,
    byref,
    c_int,
    c_uint,
    c_ulong,
    c_void_p,
    windll,
)
from ctypes.wintypes import BOOL, HWND, LPCWSTR, LPWSTR
from typing import TYPE_CHECKING, Sequence

from com_types import GUID
from hresult import HRESULT

if TYPE_CHECKING:
    import os

    from ctypes import _CData, _Pointer

# COM HRESULT definitions
E_NOINTERFACE = HRESULT(0x80004002)

SFGAO_FILESYSTEM = 0x40000000
SFGAO_FOLDER = 0x20000000
SIGDN_FILESYSPATH = c_ulong(0x80058000)
SIGDN_NORMALDISPLAY = c_ulong(0x00000000)
SICHINT_CANONICAL = 0x10000000
IID_IFileDialog = GUID("{4292C689-9298-4A37-B8F0-03835B7B6EE6}")
IID_IFileSaveDialog = GUID("{84BCCD23-5FDE-4CDB-AEAA-AE8CD7A4D575}")
IID_IFileOpenDialog = GUID("{D57C7288-D4AD-4768-BE02-9D969532D960}")
CLSID_FileDialog = GUID("{3D9C8F03-50D4-4E40-BB11-70E74D3F10F3}")
CLSID_FileOpenDialog = GUID("{DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7}")
CLSID_FileSaveDialog = GUID("{C0B4E2F3-BA21-4773-8DBA-335EC946EB8B}")
CLSID_IShellLibrary = GUID("{D9B3211D-E57F-4426-AAEF-30A806ADD397}")


class COMDLG_FILTERSPEC(Structure):
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]] = [
        ("pszName", LPCWSTR),
        ("pszSpec", LPCWSTR),
    ]


# COM interface and vtable definitions

class IUnknown(Structure):
    _iid_: GUID = GUID("{00000000-0000-0000-C000-000000000046}")

class IUnknownVTable(Structure):
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]] = [
        ("QueryInterface", WINFUNCTYPE(HRESULT, POINTER(IUnknown), POINTER(GUID), POINTER(POINTER(IUnknown)))),
        ("AddRef", WINFUNCTYPE(c_ulong, POINTER(IUnknown))),
        ("Release", WINFUNCTYPE(c_ulong, POINTER(IUnknown))),
    ]

IUnknown._fields_ =  [("lpVtbl", POINTER(IUnknownVTable))]


class IModalWindow(Structure):
    _iid_ = GUID("{B4DB1657-70D7-485E-8E3E-6FCB5A5C1802}")

class IModalWindowVTable(Structure):
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]] = [
        ("QueryInterface", WINFUNCTYPE(HRESULT, POINTER(IModalWindow), POINTER(GUID), POINTER(POINTER(IUnknown)))),
        ("AddRef", WINFUNCTYPE(c_ulong, POINTER(IModalWindow))),
        ("Release", WINFUNCTYPE(c_ulong, POINTER(IModalWindow))),
        ("Show", WINFUNCTYPE(HRESULT, POINTER(IModalWindow), HWND)),
    ]

IModalWindow._fields_ =  [("lpVtbl", POINTER(IModalWindowVTable))]


class IShellItem(Structure):
    _iid_ = GUID("{43826D1E-E718-42EE-BC55-A1E261C37BFE}")

    @classmethod
    def from_path(cls, path: os.PathLike | str) -> _Pointer[IShellItem]:
        item = POINTER(IShellItem)()
        hr = SHCreateItemFromParsingName(path, None, byref(IShellItem._iid_), byref(item))
        if hr != 0:
            raise OSError(f"SHCreateItemFromParsingName failed! HRESULT: {hr}")
        return item

class IShellItemVTable(Structure):
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]] = [
        ("QueryInterface", WINFUNCTYPE(HRESULT, POINTER(IShellItem), POINTER(GUID), POINTER(POINTER(IUnknown)))),
        ("AddRef", WINFUNCTYPE(c_ulong, POINTER(IShellItem))),
        ("Release", WINFUNCTYPE(c_ulong, POINTER(IShellItem))),
        ("BindToHandler", WINFUNCTYPE(HRESULT, POINTER(IShellItem), POINTER(IUnknown), POINTER(GUID), POINTER(GUID), POINTER(c_void_p))),
        ("GetParent", WINFUNCTYPE(HRESULT, POINTER(IShellItem), POINTER(POINTER(IShellItem)))),
        ("GetDisplayName", WINFUNCTYPE(HRESULT, POINTER(IShellItem), c_int, POINTER(LPWSTR))),
        ("GetAttributes", WINFUNCTYPE(HRESULT, POINTER(IShellItem), c_ulong, POINTER(c_ulong))),
        ("Compare", WINFUNCTYPE(HRESULT, POINTER(IShellItem), POINTER(IShellItem), c_int, POINTER(c_int))),
    ]

IShellItem._fields_ =  [("lpVtbl", POINTER(IShellItemVTable))]


class IShellItemArray(Structure):
    _iid_ = GUID("{B63EA76D-1F85-456F-A19C-48159EFA858B}")
class IShellItemArrayVTable(Structure):
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]] = [
        ("QueryInterface", WINFUNCTYPE(HRESULT, POINTER(IShellItemArray), POINTER(GUID), POINTER(POINTER(IUnknown)))),
        ("AddRef", WINFUNCTYPE(c_ulong, POINTER(IShellItemArray))),
        ("Release", WINFUNCTYPE(c_ulong, POINTER(IShellItemArray))),
        ("GetCount", WINFUNCTYPE(HRESULT, POINTER(IShellItemArray), POINTER(c_uint))),
        ("GetItemAt", WINFUNCTYPE(HRESULT, POINTER(IShellItemArray), c_uint, POINTER(POINTER(IShellItem)))),
        ("EnumItems", WINFUNCTYPE(HRESULT, POINTER(IShellItemArray), POINTER(POINTER(IUnknown)))),
    ]
IShellItemArray._fields_ =  [("lpVtbl", POINTER(IShellItemArrayVTable))]


class IEnumShellItems(Structure):
    _iid_ = GUID("{70629033-E363-4A28-A567-0DB78006E6D7}")
class IEnumShellItemsVTable(Structure):
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]] = [
        ("QueryInterface", WINFUNCTYPE(HRESULT, POINTER(IEnumShellItems), POINTER(GUID), POINTER(POINTER(IUnknown)))),
        ("AddRef", WINFUNCTYPE(c_ulong, POINTER(IEnumShellItems))),
        ("Release", WINFUNCTYPE(c_ulong, POINTER(IEnumShellItems))),
        ("Next", WINFUNCTYPE(HRESULT, POINTER(IEnumShellItems), c_ulong, POINTER(POINTER(IShellItem)), POINTER(c_ulong))),
        ("Skip", WINFUNCTYPE(HRESULT, POINTER(IEnumShellItems), c_ulong)),
        ("Reset", WINFUNCTYPE(HRESULT, POINTER(IEnumShellItems))),
        ("Clone", WINFUNCTYPE(HRESULT, POINTER(IEnumShellItems), POINTER(POINTER(IEnumShellItems)))),
    ]
IEnumShellItems._fields_ = [("lpVtbl", POINTER(IEnumShellItemsVTable))]


class IShellItemFilter(Structure):
    _iid_ = GUID("{2659B475-EEB8-48B7-8F07-B378810F48CF}")
class IShellItemFilterVTable(Structure):
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]] = [
        ("QueryInterface", WINFUNCTYPE(HRESULT, POINTER(IShellItemFilter), POINTER(GUID), POINTER(POINTER(IUnknown)))),
        ("AddRef", WINFUNCTYPE(c_ulong, POINTER(IShellItemFilter))),
        ("Release", WINFUNCTYPE(c_ulong, POINTER(IShellItemFilter))),
        ("IncludeItem", WINFUNCTYPE(HRESULT, POINTER(IShellItemFilter), POINTER(IShellItem))),
        ("GetEnumFlagsForItem", WINFUNCTYPE(HRESULT, POINTER(IShellItemFilter), POINTER(IShellItem), POINTER(c_ulong))),
    ]
IShellItemFilter._fields_ =  [("lpVtbl", POINTER(IShellItemFilterVTable))]





class IFileDialog(Structure):
    _iid_ = GUID("{4292C689-9298-4A37-B8F0-03835B7B6EE6}")
class IFileDialogEvents(Structure):
    _iid_ = GUID("{973510DB-7D7F-452B-8975-74A85828D354}")

class IFileDialogVTable(Structure):
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]] = [
        ("QueryInterface", WINFUNCTYPE(HRESULT, POINTER(IFileDialog), POINTER(GUID), POINTER(POINTER(IUnknown)))),
        ("AddRef", WINFUNCTYPE(c_ulong, POINTER(IFileDialog))),
        ("Release", WINFUNCTYPE(c_ulong, POINTER(IFileDialog))),
        ("Show", WINFUNCTYPE(HRESULT, POINTER(IFileDialog), HWND)),
        ("SetFileTypes", WINFUNCTYPE(HRESULT, POINTER(IFileDialog), c_uint, POINTER(COMDLG_FILTERSPEC))),
        ("SetFileTypeIndex", WINFUNCTYPE(HRESULT, POINTER(IFileDialog), c_uint)),
        ("GetFileTypeIndex", WINFUNCTYPE(HRESULT, POINTER(IFileDialog), POINTER(c_uint))),
        ("Advise", WINFUNCTYPE(HRESULT, POINTER(IFileDialog), POINTER(IFileDialogEvents), POINTER(c_uint))),
        ("Unadvise", WINFUNCTYPE(HRESULT, POINTER(IFileDialog), c_uint)),
        ("SetOptions", WINFUNCTYPE(HRESULT, POINTER(IFileDialog), c_uint)),
        ("GetOptions", WINFUNCTYPE(HRESULT, POINTER(IFileDialog), POINTER(c_uint))),
        ("SetDefaultFolder", WINFUNCTYPE(HRESULT, POINTER(IFileDialog), POINTER(IShellItem))),
        ("SetFolder", WINFUNCTYPE(HRESULT, POINTER(IFileDialog), POINTER(IShellItem))),
        ("GetFolder", WINFUNCTYPE(HRESULT, POINTER(IFileDialog), POINTER(POINTER(IShellItem)))),
        ("GetCurrentSelection", WINFUNCTYPE(HRESULT, POINTER(IFileDialog), POINTER(POINTER(IShellItem)))),
        ("SetFileName", WINFUNCTYPE(HRESULT, POINTER(IFileDialog), LPCWSTR)),
        ("GetFileName", WINFUNCTYPE(HRESULT, POINTER(IFileDialog), POINTER(LPWSTR))),
        ("SetTitle", WINFUNCTYPE(HRESULT, POINTER(IFileDialog), LPCWSTR)),
        ("SetOkButtonLabel", WINFUNCTYPE(HRESULT, POINTER(IFileDialog), LPCWSTR)),
        ("SetFileNameLabel", WINFUNCTYPE(HRESULT, POINTER(IFileDialog), LPCWSTR)),
        ("GetResult", WINFUNCTYPE(HRESULT, POINTER(IFileDialog), POINTER(POINTER(IShellItem)))),
        ("AddPlace", WINFUNCTYPE(HRESULT, POINTER(IFileDialog), POINTER(IShellItem), c_int)),
        ("SetDefaultExtension", WINFUNCTYPE(HRESULT, POINTER(IFileDialog), LPCWSTR)),
        ("Close", WINFUNCTYPE(HRESULT, POINTER(IFileDialog), HRESULT)),
        ("SetClientGuid", WINFUNCTYPE(HRESULT, POINTER(IFileDialog), POINTER(GUID))),
        ("ClearClientData", WINFUNCTYPE(HRESULT, POINTER(IFileDialog))),
        ("SetFilter", WINFUNCTYPE(HRESULT, POINTER(IFileDialog), POINTER(IShellItemFilter))),
    ]

IFileDialog._fields_ =  [("lpVtbl", POINTER(IFileDialogVTable))]

class IFileDialogEventsVTable(Structure):
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]] = [
        ("QueryInterface", WINFUNCTYPE(HRESULT, POINTER(IFileDialogEvents), POINTER(GUID), POINTER(POINTER(IUnknown)))),
        ("AddRef", WINFUNCTYPE(c_ulong, POINTER(IFileDialogEvents))),
        ("Release", WINFUNCTYPE(c_ulong, POINTER(IFileDialogEvents))),
        ("OnFileOk", WINFUNCTYPE(HRESULT, POINTER(IFileDialogEvents), POINTER(IFileDialog))),
        ("OnFolderChanging", WINFUNCTYPE(HRESULT, POINTER(IFileDialogEvents), POINTER(IFileDialog), POINTER(IShellItem))),
        ("OnFolderChange", WINFUNCTYPE(HRESULT, POINTER(IFileDialogEvents), POINTER(IFileDialog))),
        ("OnSelectionChange", WINFUNCTYPE(HRESULT, POINTER(IFileDialogEvents), POINTER(IFileDialog))),
        ("OnShareViolation", WINFUNCTYPE(HRESULT, POINTER(IFileDialogEvents), POINTER(IFileDialog), POINTER(IShellItem), POINTER(c_int))),
        ("OnTypeChange", WINFUNCTYPE(HRESULT, POINTER(IFileDialogEvents), POINTER(IFileDialog))),
        ("OnOverwrite", WINFUNCTYPE(HRESULT, POINTER(IFileDialogEvents), POINTER(IFileDialog), POINTER(IShellItem), POINTER(c_int))),
    ]

IFileDialogEvents._fields_ = [("lpVtbl", POINTER(IFileDialogEventsVTable))]


class IEnumShellItems(Structure):
    _iid_ = GUID("{8E8A8A7F-C50F-4F1F-B8F4-3B18BF1C5F40}")

class IEnumShellItemsVTable(Structure):
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]] = [
        ("QueryInterface", WINFUNCTYPE(HRESULT, POINTER(IEnumShellItems), POINTER(GUID), POINTER(POINTER(IUnknown)))),
        ("AddRef", WINFUNCTYPE(c_ulong, POINTER(IEnumShellItems))),
        ("Release", WINFUNCTYPE(c_ulong, POINTER(IEnumShellItems))),
        ("Next", WINFUNCTYPE(HRESULT, POINTER(IEnumShellItems), c_uint, POINTER(POINTER(IShellItem)), POINTER(c_uint))),
        ("Skip", WINFUNCTYPE(HRESULT, POINTER(IEnumShellItems), c_uint)),
        ("Reset", WINFUNCTYPE(HRESULT, POINTER(IEnumShellItems))),
        ("Clone", WINFUNCTYPE(HRESULT, POINTER(IEnumShellItems), POINTER(POINTER(IEnumShellItems)))),
    ]

IEnumShellItems._fields_ = [("lpVtbl", POINTER(IEnumShellItemsVTable))]


class IFileOpenDialog(Structure):
    _iid_ = GUID("{D57C7288-D4AD-4768-BE02-9D969532D960}")

class IFileOpenDialogVTable(Structure):
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]] = [
        ("QueryInterface", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), POINTER(GUID), POINTER(POINTER(IUnknown)))),
        ("AddRef", WINFUNCTYPE(c_ulong, POINTER(IFileOpenDialog))),
        ("Release", WINFUNCTYPE(c_ulong, POINTER(IFileOpenDialog))),
        ("Show", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), HWND)),
        ("SetFileTypes", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), c_uint, POINTER(COMDLG_FILTERSPEC))),
        ("SetFileTypeIndex", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), c_uint)),
        ("GetFileTypeIndex", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), POINTER(c_uint))),
        ("Advise", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), POINTER(IFileDialogEvents), POINTER(c_uint))),
        ("Unadvise", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), c_uint)),
        ("SetOptions", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), c_uint)),
        ("GetOptions", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), POINTER(c_uint))),
        ("SetDefaultFolder", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), POINTER(IShellItem))),
        ("SetFolder", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), POINTER(IShellItem))),
        ("GetFolder", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), POINTER(POINTER(IShellItem)))),
        ("GetCurrentSelection", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), POINTER(POINTER(IShellItem)))),
        ("SetFileName", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), LPCWSTR)),
        ("GetFileName", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), POINTER(LPWSTR))),
        ("SetTitle", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), LPCWSTR)),
        ("SetOkButtonLabel", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), LPCWSTR)),
        ("SetFileNameLabel", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), LPCWSTR)),
        ("GetResult", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), POINTER(POINTER(IShellItem)))),
        ("AddPlace", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), POINTER(IShellItem), c_int)),
        ("SetDefaultExtension", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), LPCWSTR)),
        ("Close", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), HRESULT)),
        ("SetClientGuid", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), POINTER(GUID))),
        ("ClearClientData", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog))),
        ("SetFilter", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), POINTER(IShellItemFilter))),
        ("GetResults", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), POINTER(POINTER(IShellItemArray)))),
        ("GetSelectedItems", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), POINTER(POINTER(IShellItemArray)))),
    ]

IFileOpenDialog._fields_ = [("lpVtbl", POINTER(IFileOpenDialogVTable))]


class IFileSaveDialog(Structure):
    _iid_ = GUID("{84BCCD23-5FDE-4CDB-AEAA-AE8CD7A4D575}")

class IFileSaveDialogVTable(Structure):
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]] = [
        ("QueryInterface", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), POINTER(GUID), POINTER(POINTER(IUnknown)))),
        ("AddRef", WINFUNCTYPE(c_ulong, POINTER(IFileSaveDialog))),
        ("Release", WINFUNCTYPE(c_ulong, POINTER(IFileSaveDialog))),
        ("Show", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), HWND)),
        ("SetFileTypes", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), c_uint, POINTER(COMDLG_FILTERSPEC))),
        ("SetFileTypeIndex", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), c_uint)),
        ("GetFileTypeIndex", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), POINTER(c_uint))),
        ("Advise", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), POINTER(IFileDialogEvents), POINTER(c_uint))),
        ("Unadvise", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), c_uint)),
        ("SetOptions", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), c_uint)),
        ("GetOptions", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), POINTER(c_uint))),
        ("SetDefaultFolder", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), POINTER(IShellItem))),
        ("SetFolder", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), POINTER(IShellItem))),
        ("GetFolder", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), POINTER(POINTER(IShellItem)))),
        ("GetCurrentSelection", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), POINTER(POINTER(IShellItem)))),
        ("SetFileName", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), LPCWSTR)),
        ("GetFileName", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), POINTER(LPWSTR))),
        ("SetTitle", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), LPCWSTR)),
        ("SetOkButtonLabel", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), LPCWSTR)),
        ("SetFileNameLabel", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), LPCWSTR)),
        ("GetResult", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), POINTER(POINTER(IShellItem)))),
        ("AddPlace", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), POINTER(IShellItem), c_int)),
        ("SetDefaultExtension", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), LPCWSTR)),
        ("Close", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), HRESULT)),
        ("SetClientGuid", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), POINTER(GUID))),
        ("ClearClientData", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog))),
        ("SetFilter", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), POINTER(IShellItemFilter))),
        ("SetSaveAsItem", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), POINTER(IShellItem))),
        ("SetProperties", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), POINTER(IUnknown))),
        ("SetCollectedProperties", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), POINTER(IUnknown), BOOL)),
        ("GetProperties", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), POINTER(POINTER(IUnknown)))),
        ("ApplyProperties", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), POINTER(IShellItem), POINTER(IUnknown), HWND, POINTER(IUnknown))),
    ]

IFileSaveDialog._fields_ =  [("lpVtbl", POINTER(IFileSaveDialogVTable))]


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


class FDE_SHAREVIOLATION_RESPONSE(c_int):  # noqa: N801
    FDESVR_DEFAULT = 0x00000000
    FDESVR_ACCEPT = 0x00000001
    FDESVR_REFUSE = 0x00000002


class FDE_OVERWRITE_RESPONSE(c_int):  # noqa: N801
    FDESVR_DEFAULT = 0x00000000
    FDESVR_ACCEPT = 0x00000001
    FDESVR_REFUSE = 0x00000002

# Function prototypes
SHCreateItemFromParsingName = windll.shell32.SHCreateItemFromParsingName
SHCreateItemFromParsingName.argtypes = [LPCWSTR, c_void_p, POINTER(GUID), POINTER(POINTER(IShellItem))]
SHCreateItemFromParsingName.restype = HRESULT

# Define the necessary function pointers
PFN_CoInitialize = WINFUNCTYPE(HRESULT, c_void_p)
PFN_CoUninitialize = WINFUNCTYPE(None)
PFN_CoCreateInstance = WINFUNCTYPE(HRESULT, POINTER(GUID), c_void_p, c_ulong, POINTER(GUID), POINTER(c_void_p))
PFN_CoTaskMemFree = WINFUNCTYPE(None, c_void_p)
PFN_SHCreateItemFromParsingName = WINFUNCTYPE(HRESULT, LPCWSTR, c_void_p, POINTER(GUID), POINTER(c_void_p))

# COM function pointers structure
#class COMFunctionPointers(Structure):
    # *_fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]] = [
    #     ("pCoCreateInstance", PFN_CoCreateInstance),
    #     ("pCoUninitialize", PFN_CoUninitialize),
    #     ("pCoTaskMemFree", PFN_CoTaskMemFree),
    #     ("pCoInitialize", PFN_CoInitialize),
    #     ("pSHCreateItemFromParsingName", PFN_SHCreateItemFromParsingName),
    # ]
# Above is more or less what the c++ would be, but it's easier to just do this:
class COMFunctionPointers:
    def __init__(self):
        self.pCoCreateInstance = windll.ole32.CoCreateInstance
        self.pCoUninitialize = windll.ole32.CoUninitialize
        self.pCoTaskMemFree = windll.ole32.CoTaskMemFree
        self.pCoInitialize = windll.ole32.CoInitialize
        self.pSHCreateItemFromParsingName = windll.shell32.SHCreateItemFromParsingName
