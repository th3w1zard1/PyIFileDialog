from __future__ import annotations

from ctypes import (
    POINTER,
    WINFUNCTYPE,
    Structure,
    c_char_p,
    c_int,
    c_uint,
    c_ulong,
    c_void_p,
    pointer,
    windll,
)
from ctypes.wintypes import BOOL, HWND, LPCWSTR, LPWSTR
from typing import TYPE_CHECKING, Callable, Sequence

from com_types import GUID
from hresult import HRESULT
from iunknown import COMBase, IUnknown, IUnknownVTable

if TYPE_CHECKING:
    from ctypes import _CData, _FuncPointer, _NamedFuncPointer, _Pointer

# COM HRESULT definitions
SFGAO_FILESYSTEM = 0x40000000
SFGAO_FOLDER = 0x20000000
SIGDN_FILESYSPATH = c_ulong(0x80058000)
SIGDN_NORMALDISPLAY = c_ulong(0x00000000)
SICHINT_CANONICAL = 0x10000000

class FDAP(c_int):
    FDAP_BOTTOM = 0x00000000
    FDAP_TOP = 0x00000001

S_OK = HRESULT(0)
S_FALSE = HRESULT(1)
E_NOINTERFACE = HRESULT(0x80004002)
SFGAOF = c_ulong
SFGAO_FILESYSTEM = 0x40000000
SIGDN_FILESYSPATH = c_ulong(0x80058000)
IID_IUnknown = GUID("{00000000-0000-0000-C000-000000000046}")  # GUID(0x00000000, 0x0000, 0x0000, (0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46))
IID_IClassFactory = GUID("{00000001-0000-0000-C000-000000000046}")
IID_IDispatch = GUID("{00020400-0000-0000-C000-000000000046}")
IID_IStream = GUID("{0000000c-0000-0000-C000-000000000046}")
IID_IStorage = GUID("{0000000b-0000-0000-C000-000000000046}")
IID_IBindCtx = GUID("{0000000e-0000-0000-C000-000000000046}")
IID_IFileDialog = GUID("{4292C689-9298-4A37-B8F0-03835B7B6EE6}")
IID_IEnumShellItems = GUID("{70629033-E363-4A28-A567-0DB78006E6D7}")
IID_IContextMenu = GUID("{000214e4-0000-0000-c000-000000000046}")
IID_IContextMenu2 = GUID("{000214f4-0000-0000-c000-000000000046}")
IID_IContextMenu3 = GUID("{bcfce0a0-ec17-11d0-8d10-00a0c90f2719}")
IID_IShellFolder = GUID("{000214e6-0000-0000-c000-000000000046}")
IID_IShellFolder2 = GUID("{93f2f68c-1d1b-11d3-a30e-00c04f79abd1}")
IID_IShellItem = GUID("{43826D1E-E718-42EE-BC55-A1E261C37BFE}")
IID_IShellLibrary = GUID("{11A66EFA-382E-451A-9234-1E0E12EF3085}")
IID_IShellItemArray = GUID("{B63EA76D-1F85-456F-A19C-48159EFA858B}")
IID_IShellItemFilter = GUID("{2659B475-EEB8-48B7-8F07-B378810F48CF}")
IID_IShellView = GUID("{000214e3-0000-0000-c000-000000000046}")
IID_IModalWindow = GUID("{B4DB1657-70D7-485E-8E3E-6FCB5A5C1802}")
IID_IFileSaveDialog = GUID("{84BCCD23-5FDE-4CDB-AEAA-AE8CD7A4D575}")
IID_IFileOpenDialog = GUID("{D57C7288-D4AD-4768-BE02-9D969532D960}")
IID_IFileDialogEvents = GUID("{973510DB-7D7F-452B-8975-74A85828D354}")
IID_IShellLink = GUID("{000214f9-0000-0000-c000-000000000046}")
IID_IShellLinkDataList = GUID("{45E2B4AE-B1C3-11D0-BA91-00C04FD7A083}")
CLSID_FileDialog = GUID("{3D9C8F03-50D4-4E40-BB11-70E74D3F10F3}")
CLSID_FileOpenDialog = GUID("{DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7}")
CLSID_FileSaveDialog = GUID("{C0B4E2F3-BA21-4773-8DBA-335EC946EB8B}")
CLSID_IShellLibrary = GUID("{D9B3211D-E57F-4426-AAEF-30A806ADD397}")
CLSID_ShellItem = GUID("{9AC9FBE1-E0A2-4AD6-B4EE-E212013EA917}")
CLSID_ShellItem2 = GUID("{7E9FB0D3-919F-4307-AB2E-9B1860310C93}")
CLSID_IShellItemArray = GUID("{9C73F5E4-0E6B-4F33-A6A0-1E3A8DC7A742}")

# Or use these constants
SIATTRIBFLAGS_AND = 0x00000001
SIATTRIBFLAGS_OR = 0x00000002
SIATTRIBFLAGS_APPCOMPAT = 0x00000003
SIATTRIBFLAGS_MASK = 0x00000003
SIATTRIBFLAGS_ALLITEMS = 0x00004000

class FDE_SHAREVIOLATION_RESPONSE(c_int):  # noqa: N801
    FDESVR_DEFAULT = 0x00000000
    FDESVR_ACCEPT = 0x00000001
    FDESVR_REFUSE = 0x00000002

FDE_OVERWRITE_RESPONSE = FDE_SHAREVIOLATION_RESPONSE

class COMDLG_FILTERSPEC(Structure):  # noqa: N801
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]] = [("pszName", LPCWSTR), ("pszSpec", LPCWSTR)]


class IModalWindow(Structure):
    _iid_: GUID = IID_IModalWindow
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]]

    def __init__(
        self,
        QueryInterface: Callable[[_Pointer[IModalWindow], _Pointer[GUID], _Pointer[_Pointer[IModalWindow]]], HRESULT] | None = None,  # noqa: N803
        AddRef: Callable[[_Pointer[IModalWindow]], c_ulong] | None = None,  # noqa: N803
        Release: Callable[[_Pointer[IModalWindow]], c_ulong] | None = None,  # noqa: N803
        Show: Callable[[_Pointer[IModalWindow], HWND], HRESULT] | None = None,  # noqa: N803
    ):
        self.vtable: IModalWindowVTable = IModalWindowVTable(
            QueryInterface=IModalWindowVTable._fields_[0][1](self.query_interface if QueryInterface is None else QueryInterface),  # pyright: ignore[reportCallIssue]
            AddRef=IModalWindowVTable._fields_[1][1](self.add_ref if AddRef is None else AddRef),  # pyright: ignore[reportCallIssue]
            Release=IModalWindowVTable._fields_[2][1](self.release if Release is None else Release),  # pyright: ignore[reportCallIssue]
            Show=IModalWindowVTable._fields_[3][1](self.show if Show is None else Show),  # pyright: ignore[reportCallIssue]
        )
        self.lpVtbl: _Pointer[IModalWindowVTable] = pointer(self.vtable)

    def query_interface(self, this: _Pointer[IFileDialogEvents], riid: _Pointer[GUID], ppvObject: _Pointer[_Pointer[IUnknown]]):
        print("Default QueryInterface triggered")
    def add_ref(self, this: _Pointer[IFileDialogEvents]):
        print(f"Default AddRef triggered (this={this})")
    def release(self, this: _Pointer[IFileDialogEvents]):
        print(f"Default Release triggered (this={this})")
    def show(self, this: _Pointer[IModalWindow], parent: HWND):
        print(f"Default Show triggered (this={this})")
class IModalWindowVTable(Structure):
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]] = [
        ("QueryInterface", WINFUNCTYPE(HRESULT, POINTER(IModalWindow), POINTER(GUID), POINTER(POINTER(IUnknown)))),
        ("AddRef", WINFUNCTYPE(c_ulong, POINTER(IModalWindow))),
        ("Release", WINFUNCTYPE(c_ulong, POINTER(IModalWindow))),
        ("Show", WINFUNCTYPE(HRESULT, POINTER(IModalWindow), HWND)),
    ]
    QueryInterface: Callable[[_Pointer[IModalWindow], _Pointer[GUID], _Pointer[_Pointer[IUnknown]]], HRESULT]
    AddRef: Callable[[_Pointer[IModalWindow]], c_ulong]
    Release: Callable[[_Pointer[IModalWindow]], c_ulong]
    Show: Callable[[_Pointer[IModalWindow], HWND], HRESULT]
IModalWindow._fields_ = [("lpVtbl", POINTER(IModalWindowVTable))]


class IShellItem(Structure):
    _iid_: GUID = IID_IShellItem
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]]
    def __init__(
        self,
        QueryInterface: Callable[[_Pointer[IShellItem], _Pointer[GUID], _Pointer[_Pointer[IShellItem]]], HRESULT] | None = None,  # noqa: N803
        AddRef: Callable[[_Pointer[IShellItem]], c_ulong] | None = None,  # noqa: N803
        Release: Callable[[_Pointer[IShellItem]], c_ulong] | None = None,  # noqa: N803
        BindToHandler: Callable[[_Pointer[IShellItem], _Pointer[IShellItem], _Pointer[GUID], _Pointer[GUID], _Pointer[c_void_p]], HRESULT] | None = None,  # noqa: N803
        GetDisplayName: Callable[[_Pointer[IShellItem], c_uint, _Pointer[LPWSTR]], HRESULT] | None = None,  # noqa: N803
        GetAttributes: Callable[[_Pointer[IShellItem], c_ulong, _Pointer[c_ulong]], HRESULT] | None = None,  # noqa: N803
        Compare: Callable[[_Pointer[IShellItem], _Pointer[IShellItem], c_ulong, _Pointer[c_int]], HRESULT] | None = None,  # noqa: N803
        GetParent: Callable[[_Pointer[IShellItem], _Pointer[_Pointer[IShellItem]]], HRESULT] | None = None,  # noqa: N803
    ):
        self.vtable: IShellItemVTable = IShellItemVTable(
            QueryInterface=IShellItemVTable._fields_[0][1](self.query_interface if QueryInterface is None else QueryInterface),  # pyright: ignore[reportCallIssue]
            AddRef=IShellItemVTable._fields_[1][1](self.add_ref if AddRef is None else AddRef),  # pyright: ignore[reportCallIssue]
            Release=IShellItemVTable._fields_[2][1](self.release if Release is None else Release),  # pyright: ignore[reportCallIssue]
            BindToHandler=IShellItemVTable._fields_[3][1](self.bind_to_handler if BindToHandler is None else BindToHandler),  # pyright: ignore[reportCallIssue]
            GetParent=IShellItemVTable._fields_[4][1](self.get_parent if GetParent is None else GetParent),  # pyright: ignore[reportCallIssue]
            GetDisplayName=IShellItemVTable._fields_[5][1](self.get_display_name if GetDisplayName is None else GetDisplayName),  # pyright: ignore[reportCallIssue]
            GetAttributes=IShellItemVTable._fields_[6][1](self.get_attributes if GetAttributes is None else GetAttributes),  # pyright: ignore[reportCallIssue]
            Compare=IShellItemVTable._fields_[7][1](self.compare if Compare is None else Compare),  # pyright: ignore[reportCallIssue]
        )
        self.lpVtbl: _Pointer[IShellItemVTable] = pointer(self.vtable)
    def query_interface(self, this: _Pointer[IFileDialogEvents], riid: _Pointer[GUID], ppvObject: _Pointer[_Pointer[IUnknown]]):
        print("Default QueryInterface triggered")
    def add_ref(self, this: _Pointer[IFileDialogEvents]):
        print(f"Default AddRef triggered (this={this})")
    def release(self, this: _Pointer[IFileDialogEvents]):
        print(f"Default Release triggered (this={this})")
    def bind_to_handler(self, this: _Pointer[IShellItem], pbc: _Pointer[IShellItem], bhid: _Pointer[GUID], riid: _Pointer[GUID], ppv: _Pointer[c_void_p]):
        print(f"Default BindToHandler triggered (this={this})")
    def get_display_name(self, this: _Pointer[IShellItem], sigdnName: c_uint, ppszName: _Pointer[LPWSTR]):
        print(f"Default GetDisplayName triggered (this={this})")
    def get_attributes(self, this: _Pointer[IShellItem], sfgaoMask: c_ulong, psfgaoAttribs: _Pointer[c_ulong]):
        print(f"Default GetAttributes triggered (this={this})")
    def compare(self, this: _Pointer[IShellItem], psi: _Pointer[IShellItem], hint: c_ulong, piOrder: _Pointer[c_int]):
        print(f"Default Compare triggered (this={this})")
    def get_parent(self, this: _Pointer[IShellItem], ppsi: _Pointer[_Pointer[IShellItem]]):
        print(f"Default GetParent triggered (this={this})")
class IShellItemVTable(Structure):
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]] = [
        ("QueryInterface", WINFUNCTYPE(HRESULT, POINTER(IShellItem), POINTER(GUID), POINTER(POINTER(IUnknown)))),
        ("AddRef", WINFUNCTYPE(c_ulong, POINTER(IShellItem))),
        ("Release", WINFUNCTYPE(c_ulong, POINTER(IShellItem))),
        ("BindToHandler", WINFUNCTYPE(HRESULT, POINTER(IShellItem), POINTER(IShellItem), POINTER(GUID), POINTER(GUID), POINTER(c_void_p))),
        ("GetParent", WINFUNCTYPE(HRESULT, POINTER(IShellItem), POINTER(POINTER(IShellItem)))),
        ("GetDisplayName", WINFUNCTYPE(HRESULT, POINTER(IShellItem), c_uint, POINTER(LPWSTR))),
        ("GetAttributes", WINFUNCTYPE(HRESULT, POINTER(IShellItem), c_ulong, POINTER(c_ulong))),
        ("Compare", WINFUNCTYPE(HRESULT, POINTER(IShellItem), POINTER(IShellItem), c_ulong, POINTER(c_int))),
    ]
    QueryInterface: Callable[[_Pointer[IShellItem], _Pointer[GUID], _Pointer[_Pointer[IUnknown]]], HRESULT]
    AddRef: Callable[[_Pointer[IShellItem]], c_ulong]
    Release: Callable[[_Pointer[IShellItem]], c_ulong]
    BindToHandler: Callable[[_Pointer[IShellItem], _Pointer[IShellItem], _Pointer[GUID], _Pointer[GUID], _Pointer[c_void_p]], HRESULT]
    GetParent: Callable[[_Pointer[IShellItem], _Pointer[_Pointer[IShellItem]]], HRESULT]
    GetDisplayName: Callable[[_Pointer[IShellItem], c_uint, _Pointer[LPWSTR]], HRESULT]
    GetAttributes: Callable[[_Pointer[IShellItem], c_ulong, _Pointer[c_ulong]], HRESULT]
    Compare: Callable[[_Pointer[IShellItem], _Pointer[IShellItem], c_ulong, _Pointer[c_int]], HRESULT]
IShellItem._fields_ = [("lpVtbl", POINTER(IShellItemVTable))]


class IShellItemArray(COMBase):
    _iid_: GUID = IID_IShellItemArray
    lpVtbl: _Pointer[IShellItemArrayVTable]
class IShellItemArrayVTable(IUnknownVTable):
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]] = [
        *IUnknownVTable._fields_,
        ("GetCount", WINFUNCTYPE(HRESULT, POINTER(IShellItemArray), POINTER(c_uint))),
        ("GetItemAt", WINFUNCTYPE(HRESULT, POINTER(IShellItemArray), c_uint, POINTER(POINTER(IShellItem)))),
        ("EnumItems", WINFUNCTYPE(HRESULT, POINTER(IShellItemArray), POINTER(POINTER(IUnknown)))),
    ]
IShellItemArray._fields_ = [("lpVtbl", POINTER(IShellItemArrayVTable))]


class IEnumShellItems(COMBase):
    _iid_ = GUID("{70629033-E363-4A28-A567-0DB78006E6D7}")
    lpVtbl: _Pointer[IEnumShellItemsVTable]
class IEnumShellItemsVTable(IUnknownVTable):
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]] = [
        *IUnknownVTable._fields_,
        ("Next", WINFUNCTYPE(HRESULT, POINTER(IEnumShellItems), c_ulong, POINTER(POINTER(IShellItem)), POINTER(c_ulong))),
        ("Skip", WINFUNCTYPE(HRESULT, POINTER(IEnumShellItems), c_ulong)),
        ("Reset", WINFUNCTYPE(HRESULT, POINTER(IEnumShellItems))),
        ("Clone", WINFUNCTYPE(HRESULT, POINTER(IEnumShellItems), POINTER(POINTER(IEnumShellItems)))),
    ]
IEnumShellItems._fields_ = [("lpVtbl", POINTER(IEnumShellItemsVTable))]


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


class IFileDialog(Structure):
    _iid_: GUID = IID_IFileDialog
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]]
    QueryInterface: Callable[[_Pointer[IFileDialog], _Pointer[GUID], _Pointer[_Pointer[IUnknown]]], HRESULT]
    AddRef: Callable[[_Pointer[IFileDialog]], c_ulong]
    Release: Callable[[_Pointer[IFileDialog]], c_ulong]
    Show: Callable[[_Pointer[IFileDialog], HWND], HRESULT]
    SetFileTypes: Callable[[_Pointer[IFileDialog], c_uint, _Pointer[COMDLG_FILTERSPEC]], HRESULT]
    SetFileTypeIndex: Callable[[_Pointer[IFileDialog], c_uint], HRESULT]
    GetFileTypeIndex: Callable[[_Pointer[IFileDialog], _Pointer[c_uint]], HRESULT]
    Advise: Callable[[_Pointer[IFileDialog], _Pointer[IFileDialogEvents], _Pointer[c_ulong]], HRESULT]
    Unadvise: Callable[[_Pointer[IFileDialog], c_ulong], HRESULT]
    SetOptions: Callable[[_Pointer[IFileDialog], c_ulong], HRESULT]
    GetOptions: Callable[[_Pointer[IFileDialog], _Pointer[c_ulong]], HRESULT]
    SetDefaultFolder: Callable[[_Pointer[IFileDialog], _Pointer[IShellItem]], HRESULT]
    SetFolder: Callable[[_Pointer[IFileDialog], _Pointer[IShellItem]], HRESULT]
    GetFolder: Callable[[_Pointer[IFileDialog], _Pointer[_Pointer[IShellItem]]], HRESULT]
    GetCurrentSelection: Callable[[_Pointer[IFileDialog], _Pointer[_Pointer[IShellItem]]], HRESULT]
    SetFileName: Callable[[_Pointer[IFileDialog], LPCWSTR], HRESULT]
    GetFileName: Callable[[_Pointer[IFileDialog], _Pointer[LPWSTR]], HRESULT]
    SetTitle: Callable[[_Pointer[IFileDialog], LPCWSTR], HRESULT]
    SetOkButtonLabel: Callable[[_Pointer[IFileDialog], LPCWSTR], HRESULT]
    SetFileNameLabel: Callable[[_Pointer[IFileDialog], LPCWSTR], HRESULT]
    GetResult: Callable[[_Pointer[IFileDialog], _Pointer[_Pointer[IShellItem]]], HRESULT]
    AddPlace: Callable[[_Pointer[IFileDialog], _Pointer[IShellItem], c_int], HRESULT]
    SetDefaultExtension: Callable[[_Pointer[IFileDialog], LPCWSTR], HRESULT]
    Close: Callable[[_Pointer[IFileDialog], HRESULT], HRESULT]
    SetClientGuid: Callable[[_Pointer[IFileDialog], _Pointer[GUID]], HRESULT]
    ClearClientData: Callable[[_Pointer[IFileDialog]], HRESULT]
    SetFilter: Callable[[_Pointer[IFileDialog], _Pointer[IShellItemFilter]], HRESULT]
    GetResults: Callable[[_Pointer[IFileDialog], _Pointer[_Pointer[IShellItemArray]]], HRESULT]
    GetSelectedItems: Callable[[_Pointer[IFileDialog], _Pointer[_Pointer[IShellItemArray]]], HRESULT]
    def __init__(
        self,
        QueryInterface: Callable[[_Pointer[IFileDialog], _Pointer[GUID], _Pointer[_Pointer[IUnknown]]], HRESULT] | None = None,  # noqa: N803
        AddRef: Callable[[_Pointer[IFileDialog]], c_ulong] | None = None,  # noqa: N803
        Release: Callable[[_Pointer[IFileDialog]], c_ulong] | None = None,  # noqa: N803
        Show: Callable[[_Pointer[IFileDialog], HWND], HRESULT] | None = None,  # noqa: N803
        SetFileTypes: Callable[[_Pointer[IFileDialog], c_uint, _Pointer[COMDLG_FILTERSPEC]], HRESULT] | None = None,  # noqa: N803
        SetFileTypeIndex: Callable[[_Pointer[IFileDialog], c_uint], HRESULT] | None = None,  # noqa: N803
        GetFileTypeIndex: Callable[[_Pointer[IFileDialog], _Pointer[c_uint]], HRESULT] | None = None,  # noqa: N803
        Advise: Callable[[_Pointer[IFileDialog], _Pointer[IFileDialogEvents], _Pointer[c_ulong]], HRESULT] | None = None,  # noqa: N803
        Unadvise: Callable[[_Pointer[IFileDialog], c_ulong], HRESULT] | None = None,  # noqa: N803
        SetOptions: Callable[[_Pointer[IFileDialog], c_ulong], HRESULT] | None = None,  # noqa: N803
        GetOptions: Callable[[_Pointer[IFileDialog], _Pointer[c_ulong]], HRESULT] | None = None,  # noqa: N803
        SetDefaultFolder: Callable[[_Pointer[IFileDialog], _Pointer[IShellItem]], HRESULT] | None = None,  # noqa: N803
        SetFolder: Callable[[_Pointer[IFileDialog], _Pointer[IShellItem]], HRESULT] | None = None,  # noqa: N803
        GetFolder: Callable[[_Pointer[IFileDialog], _Pointer[_Pointer[IShellItem]]], HRESULT] | None = None,  # noqa: N803
        GetCurrentSelection: Callable[[_Pointer[IFileDialog], _Pointer[_Pointer[IShellItem]]], HRESULT] | None = None,  # noqa: N803
        SetFileName: Callable[[_Pointer[IFileDialog], LPCWSTR], HRESULT] | None = None,  # noqa: N803
        GetFileName: Callable[[_Pointer[IFileDialog], _Pointer[LPWSTR]], HRESULT] | None = None,  # noqa: N803
        SetTitle: Callable[[_Pointer[IFileDialog], LPCWSTR], HRESULT] | None = None,  # noqa: N803
        SetOkButtonLabel: Callable[[_Pointer[IFileDialog], LPCWSTR], HRESULT] | None = None,  # noqa: N803
        SetFileNameLabel: Callable[[_Pointer[IFileDialog], LPCWSTR], HRESULT] | None = None,  # noqa: N803
        GetResult: Callable[[_Pointer[IFileDialog], _Pointer[_Pointer[IShellItem]]], HRESULT] | None = None,  # noqa: N803
        AddPlace: Callable[[_Pointer[IFileDialog], _Pointer[IShellItem], c_int], HRESULT] | None = None,  # noqa: N803
        SetDefaultExtension: Callable[[_Pointer[IFileDialog], LPCWSTR], HRESULT] | None = None,  # noqa: N803
        Close: Callable[[_Pointer[IFileDialog], HRESULT], HRESULT] | None = None,  # noqa: N803
        SetClientGuid: Callable[[_Pointer[IFileDialog], _Pointer[GUID]], HRESULT] | None = None,  # noqa: N803
        ClearClientData: Callable[[_Pointer[IFileDialog]], HRESULT] | None = None,  # noqa: N803
        SetFilter: Callable[[_Pointer[IFileDialog], _Pointer[IShellItemFilter]], HRESULT] | None = None,  # noqa: N803
        GetResults: Callable[[_Pointer[IFileDialog], _Pointer[_Pointer[IShellItemArray]]], HRESULT] | None = None,  # noqa: N803
        GetSelectedItems: Callable[[_Pointer[IFileDialog], _Pointer[_Pointer[IShellItemArray]]], HRESULT] | None = None,  # noqa: N803
    ):
        self.vtable: IFileDialogVTable = IFileDialogVTable(
            QueryInterface=IFileDialogVTable._fields_[0][1](self.query_interface if QueryInterface is None else QueryInterface),  # pyright: ignore[reportCallIssue]
            AddRef=IFileDialogVTable._fields_[1][1](self.add_ref if AddRef is None else AddRef),  # pyright: ignore[reportCallIssue]
            Release=IFileDialogVTable._fields_[2][1](self.release if Release is None else Release),  # pyright: ignore[reportCallIssue]
            Show=IFileDialogVTable._fields_[3][1](self.show if Show is None else Show),  # pyright: ignore[reportCallIssue]
            SetFileTypes=IFileDialogVTable._fields_[4][1](self.set_file_types if SetFileTypes is None else SetFileTypes),  # pyright: ignore[reportCallIssue]
            SetFileTypeIndex=IFileDialogVTable._fields_[5][1](self.set_file_type_index if SetFileTypeIndex is None else SetFileTypeIndex),  # pyright: ignore[reportCallIssue]
            GetFileTypeIndex=IFileDialogVTable._fields_[6][1](self.get_file_type_index if GetFileTypeIndex is None else GetFileTypeIndex),  # pyright: ignore[reportCallIssue]
            Advise=IFileDialogVTable._fields_[7][1](self.advise if Advise is None else Advise),  # pyright: ignore[reportCallIssue]
            Unadvise=IFileDialogVTable._fields_[8][1](self.unadvise if Unadvise is None else Unadvise),  # pyright: ignore[reportCallIssue]
            SetOptions=IFileDialogVTable._fields_[9][1](self.set_options if SetOptions is None else SetOptions),  # pyright: ignore[reportCallIssue]
            GetOptions=IFileDialogVTable._fields_[10][1](self.get_options if GetOptions is None else GetOptions),  # pyright: ignore[reportCallIssue]
            SetDefaultFolder=IFileDialogVTable._fields_[11][1](self.set_default_folder if SetDefaultFolder is None else SetDefaultFolder),  # pyright: ignore[reportCallIssue]
            SetFolder=IFileDialogVTable._fields_[12][1](self.set_folder if SetFolder is None else SetFolder),  # pyright: ignore[reportCallIssue]
            GetFolder=IFileDialogVTable._fields_[13][1](self.get_folder if GetFolder is None else GetFolder),  # pyright: ignore[reportCallIssue]
            GetCurrentSelection=IFileDialogVTable._fields_[14][1](self.get_current_selection if GetCurrentSelection is None else GetCurrentSelection),  # pyright: ignore[reportCallIssue]
            SetFileName=IFileDialogVTable._fields_[15][1](self.set_file_name if SetFileName is None else SetFileName),  # pyright: ignore[reportCallIssue]
            GetFileName=IFileDialogVTable._fields_[16][1](self.get_file_name if GetFileName is None else GetFileName),  # pyright: ignore[reportCallIssue]
            SetTitle=IFileDialogVTable._fields_[17][1](self.set_title if SetTitle is None else SetTitle),  # pyright: ignore[reportCallIssue]
            SetOkButtonLabel=IFileDialogVTable._fields_[18][1](self.set_ok_button_label if SetOkButtonLabel is None else SetOkButtonLabel),  # pyright: ignore[reportCallIssue]
            SetFileNameLabel=IFileDialogVTable._fields_[19][1](self.set_file_name_label if SetFileNameLabel is None else SetFileNameLabel),  # pyright: ignore[reportCallIssue]
            GetResult=IFileDialogVTable._fields_[20][1](self.get_result if GetResult is None else GetResult),  # pyright: ignore[reportCallIssue]
            AddPlace=IFileDialogVTable._fields_[21][1](self.add_place if AddPlace is None else AddPlace),  # pyright: ignore[reportCallIssue]
            SetDefaultExtension=IFileDialogVTable._fields_[22][1](self.set_default_extension if SetDefaultExtension is None else SetDefaultExtension),  # pyright: ignore[reportCallIssue]
            Close=IFileDialogVTable._fields_[23][1](self.close if Close is None else Close),  # pyright: ignore[reportCallIssue]
            SetClientGuid=IFileDialogVTable._fields_[24][1](self.set_client_guid if SetClientGuid is None else SetClientGuid),  # pyright: ignore[reportCallIssue]
            ClearClientData=IFileDialogVTable._fields_[25][1](self.clear_client_data if ClearClientData is None else ClearClientData),  # pyright: ignore[reportCallIssue]
            SetFilter=IFileDialogVTable._fields_[26][1](self.set_filter if SetFilter is None else SetFilter),  # pyright: ignore[reportCallIssue]
            GetResults=IFileDialogVTable._fields_[27][1](self.get_results if GetResults is None else GetResults),  # pyright: ignore[reportCallIssue]
            GetSelectedItems=IFileDialogVTable._fields_[28][1](self.get_selected_items if GetSelectedItems is None else GetSelectedItems),  # pyright: ignore[reportCallIssue]
        )
        self.lpVtbl: _Pointer[IFileDialogVTable] = pointer(self.vtable)

    def query_interface(self, this: _Pointer[IFileDialog], riid: _Pointer[GUID], ppvObject: _Pointer[_Pointer[IUnknown]]):
        print("Default QueryInterface triggered")
    def add_ref(self, this: _Pointer[IFileDialog]):
        print(f"Default AddRef triggered (this={this})")
    def release(self, this: _Pointer[IFileDialog]):
        print(f"Default Release triggered (this={this})")
    def show(self, this: _Pointer[IFileDialog], parent: HWND):
        print(f"Default Show triggered (this={this})")
    def set_file_types(self, this: _Pointer[IFileDialog], cFileTypes: c_uint, rgFilterSpec: _Pointer[COMDLG_FILTERSPEC]):
        print(f"Default SetFileTypes triggered (this={this})")
    def set_file_type_index(self, this: _Pointer[IFileDialog], iFileType: c_uint):
        print(f"Default SetFileTypeIndex triggered (this={this})")
    def get_file_type_index(self, this: _Pointer[IFileDialog], piFileType: _Pointer[c_uint]):
        print(f"Default GetFileTypeIndex triggered (this={this})")
    def advise(self, this: _Pointer[IFileDialog], pfde: _Pointer[IFileDialogEvents], pdwCookie: _Pointer[c_ulong]):
        print(f"Default Advise triggered (this={this})")
    def unadvise(self, this: _Pointer[IFileDialog], dwCookie: c_ulong):
        print(f"Default Unadvise triggered (this={this})")
    def set_options(self, this: _Pointer[IFileDialog], fos: c_ulong):
        print(f"Default SetOptions triggered (this={this})")
    def get_options(self, this: _Pointer[IFileDialog], pfos: _Pointer[c_ulong]):
        print(f"Default GetOptions triggered (this={this})")
    def set_default_folder(self, this: _Pointer[IFileDialog], psi: _Pointer[IShellItem]):
        print(f"Default SetDefaultFolder triggered (this={this})")
    def set_folder(self, this: _Pointer[IFileDialog], psi: _Pointer[IShellItem]):
        print(f"Default SetFolder triggered (this={this})")
    def get_folder(self, this: _Pointer[IFileDialog], ppsi: _Pointer[_Pointer[IShellItem]]):
        print(f"Default GetFolder triggered (this={this})")
    def get_current_selection(self, this: _Pointer[IFileDialog], ppsi: _Pointer[_Pointer[IShellItem]]):
        print(f"Default GetCurrentSelection triggered (this={this})")
    def set_file_name(self, this: _Pointer[IFileDialog], pszName: LPCWSTR):
        print(f"Default SetFileName triggered (this={this})")
    def get_file_name(self, this: _Pointer[IFileDialog], ppszName: _Pointer[LPWSTR]):
        print(f"Default GetFileName triggered (this={this})")
    def set_title(self, this: _Pointer[IFileDialog], pszTitle: LPCWSTR):
        print(f"Default SetTitle triggered (this={this})")
    def set_ok_button_label(self, this: _Pointer[IFileDialog], pszText: LPCWSTR):
        print(f"Default SetOkButtonLabel triggered (this={this})")
    def set_file_name_label(self, this: _Pointer[IFileDialog], pszLabel: LPCWSTR):
        print(f"Default SetFileNameLabel triggered (this={this})")
    def get_result(self, this: _Pointer[IFileDialog], ppsi: _Pointer[_Pointer[IShellItem]]):
        print(f"Default GetResult triggered (this={this})")
    def add_place(self, this: _Pointer[IFileDialog], psi: _Pointer[IShellItem], fdap: c_int):
        print(f"Default AddPlace triggered (this={this})")
    def set_default_extension(self, this: _Pointer[IFileDialog], pszDefaultExtension: LPCWSTR):
        print(f"Default SetDefaultExtension triggered (this={this})")
    def close(self, this: _Pointer[IFileDialog], hr: HRESULT):
        print(f"Default Close triggered (this={this})")
    def set_client_guid(self, this: _Pointer[IFileDialog], guid: _Pointer[GUID]):
        print(f"Default SetClientGuid triggered (this={this})")
    def clear_client_data(self, this: _Pointer[IFileDialog]):
        print(f"Default ClearClientData triggered (this={this})")
    def set_filter(self, this: _Pointer[IFileDialog], pFilter: _Pointer[IShellItemFilter]):
        print(f"Default SetFilter triggered (this={this})")
    def get_results(self, this: _Pointer[IFileDialog], ppenum: _Pointer[_Pointer[IShellItemArray]]):
        print(f"Default GetResults triggered (this={this})")
    def get_selected_items(self, this: _Pointer[IFileDialog], ppsai: _Pointer[_Pointer[IShellItemArray]]):
        print(f"Default GetSelectedItems triggered (this={this})")
class IFileDialogEvents(Structure):
    _iid_: GUID = IID_IFileDialogEvents
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]]

    def __init__(
        self,
        QueryInterface: Callable[[_Pointer[IFileDialogEvents], _Pointer[GUID], _Pointer[_Pointer[IFileDialogEvents]]], HRESULT] | None = None,  # noqa: N803
        AddRef: Callable[[_Pointer[IFileDialogEvents]], c_ulong] | None = None,  # noqa: N803
        Release: Callable[[_Pointer[IFileDialogEvents]], c_ulong] | None = None,  # noqa: N803
        OnFileOk: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog]], HRESULT] | None = None,  # noqa: N803
        OnFolderChanging: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog], _Pointer[IShellItem]], HRESULT] | None = None,  # noqa: N803
        OnFolderChange: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog]], HRESULT] | None = None,  # noqa: N803
        OnSelectionChange: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog]], HRESULT] | None = None,  # noqa: N803
        OnShareViolation: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog], _Pointer[IShellItem], _Pointer[c_int]], HRESULT] | None = None,  # noqa: N803
        OnTypeChange: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog]], HRESULT] | None = None,  # noqa: N803
        OnOverwrite: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog], _Pointer[IShellItem], _Pointer[c_int]], HRESULT] | None = None,  # noqa: N803
    ):
        self.vtable: IFileDialogEventsVTable = IFileDialogEventsVTable(
            QueryInterface=IFileDialogEventsVTable._fields_[0][1](self.query_interface if QueryInterface is None else QueryInterface),  # pyright: ignore[reportCallIssue]
            AddRef=IFileDialogEventsVTable._fields_[1][1](self.add_ref if AddRef is None else AddRef),  # pyright: ignore[reportCallIssue]
            Release=IFileDialogEventsVTable._fields_[2][1](self.release if Release is None else Release),  # pyright: ignore[reportCallIssue]
            OnFileOk=IFileDialogEventsVTable._fields_[3][1](self.on_file_ok if OnFileOk is None else OnFileOk),  # pyright: ignore[reportCallIssue]
            OnFolderChanging=IFileDialogEventsVTable._fields_[4][1](self.on_folder_changing if OnFolderChanging is None else OnFolderChanging),  # pyright: ignore[reportCallIssue]
            OnFolderChange=IFileDialogEventsVTable._fields_[5][1](self.on_folder_change if OnFolderChange is None else OnFolderChange),  # pyright: ignore[reportCallIssue]
            OnSelectionChange=IFileDialogEventsVTable._fields_[6][1](self.on_selection_change if OnSelectionChange is None else OnSelectionChange),  # pyright: ignore[reportCallIssue]
            OnShareViolation=IFileDialogEventsVTable._fields_[7][1](self.on_share_violation if OnShareViolation is None else OnShareViolation),  # pyright: ignore[reportCallIssue]
            OnTypeChange=IFileDialogEventsVTable._fields_[8][1](self.on_type_change if OnTypeChange is None else OnTypeChange),  # pyright: ignore[reportCallIssue]
            OnOverwrite=IFileDialogEventsVTable._fields_[9][1](self.on_overwrite if OnOverwrite is None else OnOverwrite),  # pyright: ignore[reportCallIssue]
        )
        self.lpVtbl: _Pointer[IFileDialogEventsVTable] = pointer(self.vtable)

    def query_interface(self, this: _Pointer[IFileDialogEvents], riid: _Pointer[GUID], ppvObject: _Pointer[_Pointer[IUnknown]]):
        print("Default QueryInterface triggered")
    def add_ref(self, this: _Pointer[IFileDialogEvents]):
        print(f"Default AddRef triggered (this={this})")
    def release(self, this: _Pointer[IFileDialogEvents]):
        print(f"Default Release triggered (this={this})")
    def on_file_ok(self, this: _Pointer[IFileDialogEvents], pfd: _Pointer[IFileDialog]):
        print(f"Default OnFileOk triggered (this={this})")
    def on_folder_changing(self, this: _Pointer[IFileDialogEvents], pfd: _Pointer[IFileDialog], psiFolder: _Pointer[IShellItem]):
        print(f"Default OnFolderChanging triggered (this={this})")
    def on_folder_change(self, this: _Pointer[IFileDialogEvents], pfd: _Pointer[IFileDialog]):
        print(f"Default OnFolderChange triggered (this={this})")
    def on_selection_change(self, this: _Pointer[IFileDialogEvents], pfd: _Pointer[IFileDialog]):
        print(f"Default OnSelectionChange triggered (this={this})")
    def on_share_violation(self, this: _Pointer[IFileDialogEvents], pfd: _Pointer[IFileDialog], psi: _Pointer[IShellItem], pResponse: POINTER(c_int)):
        print(f"Default OnShareViolation triggered (this={this})")
    def on_type_change(self, this: _Pointer[IFileDialogEvents], pfd: _Pointer[IFileDialog]):
        print(f"Default OnTypeChange triggered (this={this})")
    def on_overwrite(self, this: _Pointer[IFileDialogEvents], pfd: _Pointer[IFileDialog], psi: _Pointer[IShellItem], pResponse: POINTER(c_int)):
        print(f"Default OnOverwrite triggered (this={this})")
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
    QueryInterface: Callable[[_Pointer[IFileDialogEvents], _Pointer[GUID], _Pointer[_Pointer[IUnknown]]], HRESULT]
    AddRef: Callable[[_Pointer[IFileDialogEvents]], c_ulong]
    Release: Callable[[_Pointer[IFileDialogEvents]], c_ulong]
    OnFileOk: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog]], HRESULT]
    OnFolderChanging: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog], _Pointer[IShellItem]], HRESULT]
    OnFolderChange: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog]], HRESULT]
    OnSelectionChange: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog]], HRESULT]
    OnShareViolation: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog], _Pointer[IShellItem], _Pointer[c_int]], HRESULT]
    OnTypeChange: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog]], HRESULT]
    OnOverwrite: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog], _Pointer[IShellItem], _Pointer[c_int]], HRESULT]

IFileDialog._fields_ = [("lpVtbl", POINTER(IFileDialogVTable))]
IFileDialogEvents._fields_ = [("lpVtbl", POINTER(IFileDialogEventsVTable))]


class IFileOpenDialog(Structure):
    _iid_: GUID = IID_IFileOpenDialog
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]]
    def __init__(
        self,
        QueryInterface: Callable[[_Pointer[IFileOpenDialog], _Pointer[GUID], _Pointer[_Pointer[IUnknown]]], HRESULT] | None = None,  # noqa: N803
        AddRef: Callable[[_Pointer[IFileOpenDialog]], c_ulong] | None = None,  # noqa: N803
        Release: Callable[[_Pointer[IFileOpenDialog]], c_ulong] | None = None,  # noqa: N803
        Show: Callable[[_Pointer[IFileOpenDialog], HWND], HRESULT] | None = None,  # noqa: N803
        SetFileTypes: Callable[[_Pointer[IFileOpenDialog], c_uint, _Pointer[COMDLG_FILTERSPEC]], HRESULT] | None = None,  # noqa: N803
        SetFileTypeIndex: Callable[[_Pointer[IFileOpenDialog], c_uint], HRESULT] | None = None,  # noqa: N803
        GetFileTypeIndex: Callable[[_Pointer[IFileOpenDialog], _Pointer[c_uint]], HRESULT] | None = None,  # noqa: N803
        Advise: Callable[[_Pointer[IFileOpenDialog], _Pointer[IFileDialogEvents], _Pointer[c_ulong]], HRESULT] | None = None,  # noqa: N803
        Unadvise: Callable[[_Pointer[IFileOpenDialog], c_ulong], HRESULT] | None = None,  # noqa: N803
        SetOptions: Callable[[_Pointer[IFileOpenDialog], c_uint], HRESULT] | None = None,  # noqa: N803
        GetOptions: Callable[[_Pointer[IFileOpenDialog], _Pointer[c_uint]], HRESULT] | None = None,  # noqa: N803
        SetDefaultFolder: Callable[[_Pointer[IFileOpenDialog], _Pointer[IShellItem]], HRESULT] | None = None,  # noqa: N803
        SetFolder: Callable[[_Pointer[IFileOpenDialog], _Pointer[IShellItem]], HRESULT] | None = None,  # noqa: N803
        GetFolder: Callable[[_Pointer[IFileOpenDialog], _Pointer[_Pointer[IShellItem]]], HRESULT] | None = None,  # noqa: N803
        GetCurrentSelection: Callable[[_Pointer[IFileOpenDialog], _Pointer[_Pointer[IShellItem]]], HRESULT] | None = None,  # noqa: N803
        SetFileName: Callable[[_Pointer[IFileOpenDialog], LPCWSTR], HRESULT] | None = None,  # noqa: N803
        GetFileName: Callable[[_Pointer[IFileOpenDialog], _Pointer[LPWSTR]], HRESULT] | None = None,  # noqa: N803
        SetTitle: Callable[[_Pointer[IFileOpenDialog], LPCWSTR], HRESULT] | None = None,  # noqa: N803
        SetOkButtonLabel: Callable[[_Pointer[IFileOpenDialog], LPCWSTR], HRESULT] | None = None,  # noqa: N803
        SetFileNameLabel: Callable[[_Pointer[IFileOpenDialog], LPCWSTR], HRESULT] | None = None,  # noqa: N803
        GetResult: Callable[[_Pointer[IFileOpenDialog], _Pointer[_Pointer[IShellItem]]], HRESULT] | None = None,  # noqa: N803
        AddPlace: Callable[[_Pointer[IFileOpenDialog], _Pointer[IShellItem], FDAP], HRESULT] | None = None,  # noqa: N803
        SetDefaultExtension: Callable[[_Pointer[IFileOpenDialog], LPCWSTR], HRESULT] | None = None,  # noqa: N803
        Close: Callable[[_Pointer[IFileOpenDialog], HRESULT], HRESULT] | None = None,  # noqa: N803
        SetClientGuid: Callable[[_Pointer[IFileOpenDialog], _Pointer[GUID]], HRESULT] | None = None,  # noqa: N803
        ClearClientData: Callable[[_Pointer[IFileOpenDialog]], HRESULT] | None = None,  # noqa: N803
        SetFilter: Callable[[_Pointer[IFileOpenDialog], _Pointer[IShellItemFilter]], HRESULT] | None = None,  # noqa: N803
        GetResults: Callable[[_Pointer[IFileOpenDialog], _Pointer[_Pointer[IShellItemArray]]], HRESULT] | None = None,  # noqa: N803
        GetSelectedItems: Callable[[_Pointer[IFileOpenDialog], _Pointer[_Pointer[IShellItemArray]]], HRESULT] | None = None,  # noqa: N803
    ):
        self.vtable: IFileOpenDialogVTable = IFileOpenDialogVTable(
            QueryInterface=IFileOpenDialogVTable._fields_[0][1]((lambda *args: print(*args, "Default QueryInterface") or None) if QueryInterface is None else QueryInterface),  # pyright: ignore[reportCallIssue]
            AddRef=IFileOpenDialogVTable._fields_[1][1]((lambda *args: print(*args, "Default AddRef") or None) if AddRef is None else AddRef),  # pyright: ignore[reportCallIssue]
            Release=IFileOpenDialogVTable._fields_[2][1]((lambda *args: print(*args, "Default Release") or None) if Release is None else Release),  # pyright: ignore[reportCallIssue]
            Show=IFileOpenDialogVTable._fields_[3][1]((lambda *args: print(*args, "Default Show") or None) if Show is None else Show),  # pyright: ignore[reportCallIssue]
            SetFileTypes=IFileOpenDialogVTable._fields_[4][1]((lambda *args: print(*args, "Default SetFileTypes") or None) if SetFileTypes is None else SetFileTypes),  # pyright: ignore[reportCallIssue]
            SetFileTypeIndex=IFileOpenDialogVTable._fields_[5][1]((lambda *args: print(*args, "Default SetFileTypeIndex") or None) if SetFileTypeIndex is None else SetFileTypeIndex),  # pyright: ignore[reportCallIssue]
            GetFileTypeIndex=IFileOpenDialogVTable._fields_[6][1]((lambda *args: print(*args, "Default GetFileTypeIndex") or None)  if GetFileTypeIndex is None else GetFileTypeIndex),  # pyright: ignore[reportCallIssue]
            Advise=IFileOpenDialogVTable._fields_[7][1]((lambda *args: print(*args, "Default Advise") or None)  if Advise is None else Advise),  # pyright: ignore[reportCallIssue]
            Unadvise=IFileOpenDialogVTable._fields_[8][1]((lambda *args: print(*args, "Default Unadvise") or None)  if Unadvise is None else Unadvise),  # pyright: ignore[reportCallIssue]
            SetOptions=IFileOpenDialogVTable._fields_[9][1]((lambda *args: print(*args, "Default SetOptions") or None)  if SetOptions is None else SetOptions),  # pyright: ignore[reportCallIssue]
            GetOptions=IFileOpenDialogVTable._fields_[10][1]((lambda *args: print(*args, "Default GetOptions") or None)  if GetOptions is None else GetOptions),  # pyright: ignore[reportCallIssue]
            SetDefaultFolder=IFileOpenDialogVTable._fields_[11][1]((lambda *args: print(*args, "Default SetDefaultFolder") or None)  if SetDefaultFolder is None else SetDefaultFolder),  # pyright: ignore[reportCallIssue]
            SetFolder=IFileOpenDialogVTable._fields_[12][1]((lambda *args: print(*args, "Default SetFolder") or None)  if SetFolder is None else SetFolder),  # pyright: ignore[reportCallIssue]
            GetFolder=IFileOpenDialogVTable._fields_[13][1]((lambda *args: print(*args, "Default GetFolder") or None)  if GetFolder is None else GetFolder),  # pyright: ignore[reportCallIssue]
            GetCurrentSelection=IFileOpenDialogVTable._fields_[14][1]((lambda *args: print(*args, "Default GetCurrentSelection") or None)  if GetCurrentSelection is None else GetCurrentSelection),  # pyright: ignore[reportCallIssue]
            SetFileName=IFileOpenDialogVTable._fields_[15][1]((lambda *args: print(*args, "Default SetFileName") or None)  if SetFileName is None else SetFileName),  # pyright: ignore[reportCallIssue]
            GetFileName=IFileOpenDialogVTable._fields_[16][1]((lambda *args: print(*args, "Default GetFileName") or None)  if GetFileName is None else GetFileName),  # pyright: ignore[reportCallIssue]
            SetTitle=IFileOpenDialogVTable._fields_[17][1]((lambda *args: print(*args, "Default SetTitle") or None)  if SetTitle is None else SetTitle),  # pyright: ignore[reportCallIssue]
            SetOkButtonLabel=IFileOpenDialogVTable._fields_[18][1]((lambda *args: print(*args, "Default SetOkButtonLabel") or None)  if SetOkButtonLabel is None else SetOkButtonLabel),  # pyright: ignore[reportCallIssue]
            SetFileNameLabel=IFileOpenDialogVTable._fields_[19][1]((lambda *args: print(*args, "Default SetFileNameLabel") or None)  if SetFileNameLabel is None else SetFileNameLabel),  # pyright: ignore[reportCallIssue]
            GetResult=IFileOpenDialogVTable._fields_[20][1]((lambda *args: print(*args, "Default GetResult") or None)  if GetResult is None else GetResult),  # pyright: ignore[reportCallIssue]
            AddPlace=IFileOpenDialogVTable._fields_[21][1]((lambda *args: print(*args, "Default AddPlace") or None)  if AddPlace is None else AddPlace),  # pyright: ignore[reportCallIssue]
            SetDefaultExtension=IFileOpenDialogVTable._fields_[22][1]((lambda *args: print(*args, "Default SetDefaultExtension") or None)  if SetDefaultExtension is None else SetDefaultExtension),  # pyright: ignore[reportCallIssue]
            Close=IFileOpenDialogVTable._fields_[23][1]((lambda *args: print(*args, "Default Close") or None) if Close is None else Close),  # pyright: ignore[reportCallIssue]
            SetClientGuid=IFileOpenDialogVTable._fields_[24][1]((lambda *args: print(*args, "Default SetClientGuid") or None)  if SetClientGuid is None else SetClientGuid),  # pyright: ignore[reportCallIssue]
            ClearClientData=IFileOpenDialogVTable._fields_[25][1]((lambda *args: print(*args, "Default ClearClientData") or None)  if ClearClientData is None else ClearClientData),  # pyright: ignore[reportCallIssue]
            SetFilter=IFileOpenDialogVTable._fields_[26][1]((lambda *args: print(*args, "Default SetFilter") or None)  if SetFilter is None else SetFilter),  # pyright: ignore[reportCallIssue]
            GetResults=IFileOpenDialogVTable._fields_[27][1](self.get_results if GetResults is None else GetResults),  # pyright: ignore[reportCallIssue]
            GetSelectedItems=IFileOpenDialogVTable._fields_[28][1](self.get_selected_items if GetSelectedItems is None else GetSelectedItems),  # pyright: ignore[reportCallIssue]
        )
        self.lpVtbl: _Pointer[IFileOpenDialogVTable] = pointer(self.vtable)

    def get_results(self, this: _Pointer[IFileOpenDialog], ppenum: _Pointer[_Pointer[IShellItemArray]]) -> HRESULT:
        print(f"Default GetResults triggered (this={this}, ppenum={ppenum})")
        return S_OK
    def get_selected_items(self, this: _Pointer[IFileOpenDialog], ppsai: _Pointer[_Pointer[IShellItemArray]]) -> HRESULT:
        print(f"Default GetSelectedItems triggered (this={this}, ppsai={ppsai})")
        return S_OK
class IFileOpenDialogEvents(Structure):
    _iid_: GUID = IID_IFileDialogEvents
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]]

    def __init__(
        self,
        QueryInterface: Callable[[_Pointer[IFileDialogEvents], _Pointer[GUID], _Pointer[_Pointer[IFileDialogEvents]]], HRESULT] | None = None,  # noqa: N803
        AddRef: Callable[[_Pointer[IFileDialogEvents]], c_ulong] | None = None,  # noqa: N803
        Release: Callable[[_Pointer[IFileDialogEvents]], c_ulong] | None = None,  # noqa: N803
        OnFileOk: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog]], HRESULT] | None = None,  # noqa: N803
        OnFolderChanging: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog], _Pointer[IShellItem]], HRESULT] | None = None,  # noqa: N803
        OnFolderChange: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog]], HRESULT] | None = None,  # noqa: N803
        OnSelectionChange: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog]], HRESULT] | None = None,  # noqa: N803
        OnShareViolation: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog], _Pointer[IShellItem], _Pointer[c_int]], HRESULT] | None = None,  # noqa: N803
        OnTypeChange: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog]], HRESULT] | None = None,  # noqa: N803
        OnOverwrite: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog], _Pointer[IShellItem], _Pointer[c_int]], HRESULT] | None = None,  # noqa: N803
    ):
        self.vtable: IFileDialogEventsVTable = IFileDialogEventsVTable(
            QueryInterface=IFileDialogEventsVTable._fields_[0][1](self.query_interface if QueryInterface is None else QueryInterface),  # pyright: ignore[reportCallIssue]
            AddRef=IFileDialogEventsVTable._fields_[1][1](self.add_ref if AddRef is None else AddRef),  # pyright: ignore[reportCallIssue]
            Release=IFileDialogEventsVTable._fields_[2][1](self.release if Release is None else Release),  # pyright: ignore[reportCallIssue]
            OnFileOk=IFileDialogEventsVTable._fields_[3][1](self.on_file_ok if OnFileOk is None else OnFileOk),  # pyright: ignore[reportCallIssue]
            OnFolderChanging=IFileDialogEventsVTable._fields_[4][1](self.on_folder_changing if OnFolderChanging is None else OnFolderChanging),  # pyright: ignore[reportCallIssue]
            OnFolderChange=IFileDialogEventsVTable._fields_[5][1](self.on_folder_change if OnFolderChange is None else OnFolderChange),  # pyright: ignore[reportCallIssue]
            OnSelectionChange=IFileDialogEventsVTable._fields_[6][1](self.on_selection_change if OnSelectionChange is None else OnSelectionChange),  # pyright: ignore[reportCallIssue]
            OnShareViolation=IFileDialogEventsVTable._fields_[7][1](self.on_share_violation if OnShareViolation is None else OnShareViolation),  # pyright: ignore[reportCallIssue]
            OnTypeChange=IFileDialogEventsVTable._fields_[8][1](self.on_type_change if OnTypeChange is None else OnTypeChange),  # pyright: ignore[reportCallIssue]
            OnOverwrite=IFileDialogEventsVTable._fields_[9][1](self.on_overwrite if OnOverwrite is None else OnOverwrite),  # pyright: ignore[reportCallIssue]
        )
        self.lpVtbl: _Pointer[IFileDialogEventsVTable] = pointer(self.vtable)

    def query_interface(self, this: _Pointer[IFileDialogEvents], riid: _Pointer[GUID], ppvObject: _Pointer[_Pointer[IUnknown]]):
        print("Default QueryInterface triggered")
    def add_ref(self, this: _Pointer[IFileDialogEvents]):
        print(f"Default AddRef triggered (this={this})")
    def release(self, this: _Pointer[IFileDialogEvents]):
        print(f"Default Release triggered (this={this})")
    def on_file_ok(self, this: _Pointer[IFileDialogEvents], pfd: _Pointer[IFileDialog]):
        print(f"Default OnFileOk triggered (this={this})")
    def on_folder_changing(self, this: _Pointer[IFileDialogEvents], pfd: _Pointer[IFileDialog], psiFolder: _Pointer[IShellItem]):
        print(f"Default OnFolderChanging triggered (this={this})")
    def on_folder_change(self, this: _Pointer[IFileDialogEvents], pfd: _Pointer[IFileDialog]):
        print(f"Default OnFolderChange triggered (this={this})")
    def on_selection_change(self, this: _Pointer[IFileDialogEvents], pfd: _Pointer[IFileDialog]):
        print(f"Default OnSelectionChange triggered (this={this})")
    def on_share_violation(self, this: _Pointer[IFileDialogEvents], pfd: _Pointer[IFileDialog], psi: _Pointer[IShellItem], pResponse: _Pointer[c_int]):
        print(f"Default OnShareViolation triggered (this={this})")
    def on_type_change(self, this: _Pointer[IFileDialogEvents], pfd: _Pointer[IFileDialog]):
        print(f"Default OnTypeChange triggered (this={this})")
    def on_overwrite(self, this: _Pointer[IFileDialogEvents], pfd: _Pointer[IFileDialog], psi: _Pointer[IShellItem], pResponse: _Pointer[c_int]):
        print(f"Default OnOverwrite triggered (this={this})")
class IFileOpenDialogEventsVTable(Structure):
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
    QueryInterface: Callable[[_Pointer[IFileDialogEvents], _Pointer[GUID], _Pointer[_Pointer[IUnknown]]], HRESULT]
    AddRef: Callable[[_Pointer[IFileDialogEvents]], c_ulong]
    Release: Callable[[_Pointer[IFileDialogEvents]], c_ulong]
    OnFileOk: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog]], HRESULT]
    OnFolderChanging: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog], _Pointer[IShellItem]], HRESULT]
    OnFolderChange: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog]], HRESULT]
    OnSelectionChange: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog]], HRESULT]
    OnShareViolation: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog], _Pointer[IShellItem], _Pointer[c_int]], HRESULT]
    OnTypeChange: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog]], HRESULT]
    OnOverwrite: Callable[[_Pointer[IFileDialogEvents], _Pointer[IFileDialog], _Pointer[IShellItem], _Pointer[c_int]], HRESULT]
class IFileOpenDialogVTable(Structure):
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]] = [
        ("QueryInterface", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), POINTER(GUID), POINTER(POINTER(IUnknown)))),
        ("AddRef", WINFUNCTYPE(c_ulong, POINTER(IFileOpenDialog))),
        ("Release", WINFUNCTYPE(c_ulong, POINTER(IFileOpenDialog))),
        ("Show", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), HWND)),
        ("SetFileTypes", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), c_uint, POINTER(COMDLG_FILTERSPEC))),
        ("SetFileTypeIndex", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), c_uint)),
        ("GetFileTypeIndex", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), POINTER(c_uint))),
        ("Advise", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), POINTER(IFileOpenDialogEvents), POINTER(c_uint))),
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
        ("AddPlace", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), POINTER(IShellItem), FDAP)),
        ("SetDefaultExtension", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), LPCWSTR)),
        ("Close", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), HRESULT)),
        ("SetClientGuid", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), POINTER(GUID))),  # REFGUID
        ("ClearClientData", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog))),
        ("SetFilter", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), POINTER(IShellItemFilter))),
        ("GetResults", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), POINTER(POINTER(IShellItemArray)))),
        ("GetSelectedItems", WINFUNCTYPE(HRESULT, POINTER(IFileOpenDialog), POINTER(POINTER(IShellItemArray)))),
    ]
    GetResults: Callable[[_Pointer[IFileOpenDialog], _Pointer[_Pointer[IShellItemArray]]], HRESULT]
    GetSelectedItems: Callable[[_Pointer[IFileOpenDialog], _Pointer[_Pointer[IShellItemArray]]], HRESULT]
IFileOpenDialog._fields_ = [("lpVtbl", POINTER(IFileOpenDialogVTable))]


class IFileSaveDialog(COMBase):
    _iid_: GUID = IID_IFileSaveDialog
    lpVtbl: _Pointer[IFileSaveDialogVTable]
class IFileSaveDialogVTable(IFileDialogVTable):
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]] = [
        *IFileDialogVTable._fields_,
        ("SetSaveAsItem", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), POINTER(IShellItem))),
        ("SetProperties", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), POINTER(IUnknown))),
        ("SetCollectedProperties", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), POINTER(IUnknown), BOOL)),
        ("GetProperties", WINFUNCTYPE(HRESULT, POINTER(IFileSaveDialog), POINTER(POINTER(IUnknown)))),
        ("ApplyProperties", WINFUNCTYPE( HRESULT, POINTER(IFileSaveDialog), POINTER(IShellItem), POINTER(IUnknown), HWND, POINTER(IUnknown))),
    ]
IFileSaveDialog._fields_ = [("lpVtbl", POINTER(IFileSaveDialogVTable))]


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


class SIATTRIBFLAGS(c_int):
    SIATTRIBFLAGS_AND = 0x1
    SIATTRIBFLAGS_OR = 0x2
    SIATTRIBFLAGS_APPCOMPAT = 0x3
    SIATTRIBFLAGS_MASK = 0x3


# Function prototypes
SHCreateItemFromParsingName: _NamedFuncPointer = windll.shell32.SHCreateItemFromParsingName
SHCreateItemFromParsingName.argtypes = [LPCWSTR, c_void_p, POINTER(GUID), POINTER(POINTER(IShellItem))]
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
        return windll.kernel32.LoadLibraryW(dll_name)

    @staticmethod
    def resolve_function(handle: _FuncPointer, func: bytes, func_type: type[_FuncPointer]) -> _FuncPointer:
        windll.kernel32.GetProcAddress.argtypes = [c_void_p, c_char_p]
        windll.kernel32.GetProcAddress.restype = c_void_p
        return func_type(windll.kernel32.GetProcAddress(handle, func))
