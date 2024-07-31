from __future__ import annotations

from _ctypes import _Pointer
from ctypes import (
    POINTER,
    WINFUNCTYPE,
    Structure,
    byref,
    c_long,
    c_uint,
    c_ulong,
    c_void_p,
    c_wchar_p,
    cast as cast_with_ctypes,
    pointer,
    windll,
)
from ctypes.wintypes import DWORD, LPCWSTR, LPWSTR
from typing import TYPE_CHECKING, Sequence

from com_helpers import COMInitializeContext
from com_types import GUID
from hresult import HRESULT
from interfaces import (
    COMDLG_FILTERSPEC,
    SFGAO_FILESYSTEM,
    SFGAO_FOLDER,
    SIGDN_FILESYSPATH,
    SIGDN_NORMALDISPLAY,
    CLSID_FileOpenDialog,
    COMFunctionPointers,
    IEnumShellItems,
    IFileDialogEvents,
    IFileDialogEventsVTable,
    IFileOpenDialog,
    IID_IFileDialogEvents,
    IShellItem,
    IShellItemArray,
)
from iunknown import E_NOINTERFACE, HMODULE, S_OK, IID_IUnknown, IUnknown

if TYPE_CHECKING:
    from ctypes import _CData, _FuncPointer, _Pointer
    from ctypes.wintypes import HWND

    from interfaces import (
        IFileDialog,
        IFileSaveDialog,
    )
    from iunknown import LPVOID


# Helper to convert std::wstring to LPCWSTR
def string_to_LPCWSTR(s: str) -> LPCWSTR:
    return c_wchar_p(s)


# Helper to convert LPCWSTR to std::wstring
def LPCWSTR_to_string(s: LPCWSTR) -> str | None:
    return s.value


# Load COM function pointers
def LoadCOMFunctionPointers(dialog_type: type[IFileDialog | IFileOpenDialog | IFileSaveDialog]) -> COMFunctionPointers:
    comFuncPtrs = COMFunctionPointers()
    comFuncPtrs.hOle32 = comFuncPtrs.load_library("ole32.dll")
    comFuncPtrs.hShell32 = comFuncPtrs.load_library("shell32.dll")


    # Get function pointers
    if comFuncPtrs.hOle32:
        PFN_CoInitialize: type[_FuncPointer] = WINFUNCTYPE(HRESULT, POINTER(dialog_type))
        PFN_CoUninitialize: type[_FuncPointer] = WINFUNCTYPE(None)
        PFN_CoCreateInstance: type[_FuncPointer] = WINFUNCTYPE(HRESULT, POINTER(GUID), c_void_p, c_ulong, POINTER(GUID), POINTER(POINTER(dialog_type)))
        PFN_CoTaskMemFree: type[_FuncPointer] = WINFUNCTYPE(None, c_void_p)
        comFuncPtrs.pCoInitialize = comFuncPtrs.resolve_function(comFuncPtrs.hOle32, b"CoInitialize", PFN_CoInitialize)
        comFuncPtrs.pCoUninitialize = comFuncPtrs.resolve_function(comFuncPtrs.hOle32, b"CoUninitialize", PFN_CoUninitialize)
        comFuncPtrs.pCoCreateInstance = comFuncPtrs.resolve_function(comFuncPtrs.hOle32, b"CoCreateInstance", PFN_CoCreateInstance)
        comFuncPtrs.pCoTaskMemFree = comFuncPtrs.resolve_function(comFuncPtrs.hOle32, b"CoTaskMemFree", PFN_CoTaskMemFree)

    if comFuncPtrs.hShell32:
        PFN_SHCreateItemFromParsingName: type[_FuncPointer] = WINFUNCTYPE(HRESULT, LPCWSTR, c_void_p, POINTER(GUID), POINTER(POINTER(IShellItem)))
        comFuncPtrs.pSHCreateItemFromParsingName = comFuncPtrs.resolve_function(comFuncPtrs.hShell32, b"SHCreateItemFromParsingName", PFN_SHCreateItemFromParsingName)
    return comFuncPtrs


def FreeCOMFunctionPointers(comFuncPtrs: COMFunctionPointers):  # noqa: N803
    # PFN_FreeLibrary: type[_FuncPointer] = WINFUNCTYPE(c_ulong, c_void_p)
    if comFuncPtrs.hOle32:
        windll.kernel32.FreeLibrary(cast_with_ctypes(comFuncPtrs.hOle32, HMODULE))
    if comFuncPtrs.hShell32:
        windll.kernel32.FreeLibrary(cast_with_ctypes(comFuncPtrs.hShell32, HMODULE))


# FileDialogEventHandler class
class FileDialogEventHandler(Structure):
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]] = [
        ("lpVtbl", POINTER(IFileDialogEvents))
    ]

    def __init__(self):
        self.vtable = IFileDialogEventsVTable(
            QueryInterface=self.QueryInterface,
            AddRef=self.AddRef,
            Release=self.Release,
            OnFileOk=self.OnFileOk,
            OnFolderChanging=self.OnFolderChanging,
            OnFolderChange=self.OnFolderChange,
            OnSelectionChange=self.OnSelectionChange,
            OnShareViolation=self.OnShareViolation,
            OnTypeChange=self.OnTypeChange,
            OnOverwrite=self.OnOverwrite,
        )
        self.lpVtbl: _Pointer[IFileDialogEventsVTable] = pointer(self.vtable)
        self.refCount = c_long(1)

    def QueryInterface(self, riid: _Pointer[GUID], ppv: _Pointer[LPVOID]):
        if riid.contents in (IID_IUnknown, IID_IFileDialogEvents):
            ppv[0] = byref(self)
            self.AddRef()
            return S_OK
        ppv[0] = None
        return E_NOINTERFACE

    def AddRef(self):
        self.refCount.value += 1
        return self.refCount.value

    def Release(self):
        self.refCount.value -= 1
        if self.refCount.value == 0:
            del self
        return self.refCount.value

    def OnFileOk(self, pfd):
        return S_OK

    def OnFolderChanging(self, pfd, psiFolder):
        return S_OK

    def OnFolderChange(self, pfd):
        return S_OK

    def OnSelectionChange(self, pfd):
        return S_OK

    def OnShareViolation(self, pfd, psi, pResponse):
        return S_OK

    def OnTypeChange(self, pfd):
        return S_OK

    def OnOverwrite(self, pfd, psi, pResponse):
        return S_OK


def createFileDialog(
    comFuncs: COMFunctionPointers,  # noqa: N803
    fileDialog: _Pointer[c_void_p],  # noqa: N803
    clsid: GUID,
    iid: GUID,
):  # noqa: N803
    hr = comFuncs.pCoCreateInstance(byref(clsid), None, 1, byref(iid), byref(fileDialog))
    if hr != S_OK:
        raise HRESULT(hr).exception("CoCreateInstance failed! Cannot create file dialog")


def showDialog(
    pFileDialog: _Pointer[IFileOpenDialog | IFileSaveDialog | IFileDialog],  # noqa: N803
    hwndOwner: HWND,  # noqa: N803
):
    hr = pFileDialog.contents.lpVtbl.contents.Show(pFileDialog, hwndOwner)
    if hr != S_OK:
        raise HRESULT(hr).exception("Failed to show the file dialog!")


def createShellItem(comFuncs: COMFunctionPointers, path: str) -> _Pointer[IShellItem]:  # noqa: N803, ARG001
    if not comFuncs.pSHCreateItemFromParsingName:
        raise OSError("comFuncs.pSHCreateItemFromParsingName not found")
    shell_item = POINTER(IShellItem)()
    hr = comFuncs.pSHCreateItemFromParsingName(path, None, IShellItem._iid_, byref(shell_item))
    if hr != S_OK:
        raise HRESULT(hr).exception(f"Failed to create shell item from path: {path}")
    return shell_item


def setDialogAttributes(
    fileDialog: IFileDialog | IFileOpenDialog | IFileSaveDialog,  # noqa: N803
    title: str,
    okButtonLabel: str | None = None,  # noqa: N803
    fileNameLabel: str | None = None,  # noqa: N803
) -> None:
    #cast_fileDialog: IFileDialog = cast_with_ctypes(pointer(fileDialog), POINTER(IFileDialog)).contents
    #assert addressof(fileDialog) == addressof(cast_fileDialog)
    if title:
        #cast_fileDialog.lpVtbl.contents.SetTitle(cast_fileDialog, string_to_LPCWSTR(title))
        fileDialog.lpVtbl.contents.SetTitle(byref(fileDialog), string_to_LPCWSTR(title))

    if okButtonLabel:
        fileDialog.lpVtbl.contents.SetOkButtonLabel(byref(fileDialog), string_to_LPCWSTR(okButtonLabel))

    if fileNameLabel:
        fileDialog.lpVtbl.contents.SetFileName(byref(fileDialog), string_to_LPCWSTR(fileNameLabel))


# Function to configure file dialog
def configureFileDialog(  # noqa: PLR0913
    comFuncs: COMFunctionPointers,  # noqa: N803
    fileDialog: IFileOpenDialog | IFileSaveDialog | IFileDialog,  # noqa: N803
    filters: list[COMDLG_FILTERSPEC],  # noqa: N803
    defaultFolder: str,  # noqa: N803
    options: c_ulong,
    forceFileSystem: bool = False,  # noqa: N803, FBT001, FBT002
    allowMultiselect: bool = False,  # noqa: N803, FBT001, FBT002
):
    if defaultFolder:
        pFolder: _Pointer[IShellItem] = createShellItem(comFuncs, defaultFolder)
        hr = fileDialog.lpVtbl.contents.SetFolder(fileDialog, pFolder)
        if hr != S_OK:
            raise HRESULT(hr).exception("Failed to set default folder")
        pFolder.contents.lpVtbl.contents.Release(pFolder)

    if forceFileSystem:
        options = c_ulong(options.value | 0x00000040)  # FOS_FORCEFILESYSTEM

    if allowMultiselect:
        options = c_ulong(options.value | 0x00000200)  # FOS_ALLOWMULTISELECT

    if filters:
        filter_array = (COMDLG_FILTERSPEC * len(filters))()
        for i, dialogFilter in enumerate(filters):
            filter_array[i].pszName = dialogFilter.pszName
            filter_array[i].pszSpec = dialogFilter.pszSpec

        hr = fileDialog.lpVtbl.contents.SetFileTypes(byref(fileDialog), len(filters), byref(filter_array[0]))
        if hr != S_OK:
            raise HRESULT(hr).exception("Failed to set file types")

    hr = fileDialog.lpVtbl.contents.SetOptions(fileDialog, options)
    if hr != S_OK:
        raise HRESULT(hr).exception("Failed to set dialog options")


def getFileDialogResults(  # noqa: C901, PLR0912, PLR0915
    comFuncs: COMFunctionPointers,  # noqa: N803
    pFileOpenDialog: _Pointer[IFileOpenDialog | IFileSaveDialog | IFileDialog],  # noqa: N803
) -> list[str]:
    results: list[str] = []
    pResultsArray = POINTER(IShellItemArray)()
    hr = pFileOpenDialog.contents.lpVtbl.contents.GetResults(pFileOpenDialog, byref(pResultsArray))
    if hr != S_OK:
        raise HRESULT(hr).exception(f"Failed to get dialog results. HRESULT: {hr}")

    itemCount = DWORD()
    hr = pResultsArray.contents.lpVtbl.contents.GetCount(pResultsArray, byref(itemCount))
    if hr != S_OK:
        raise HRESULT(hr).exception(f"Failed to get item count. HRESULT: {hr}")

    print(f"Number of items selected: {itemCount.value}")

    pEnumShellItems = POINTER(IEnumShellItems)()
    hr = pResultsArray.contents.lpVtbl.contents.EnumItems(pResultsArray, byref(pEnumShellItems))
    if hr == S_OK and pEnumShellItems:
        pEnumItem = POINTER(IShellItem)()
        fetched = DWORD()
        while (
            pEnumShellItems.contents.lpVtbl.contents.Next(pEnumShellItems, 1, byref(pEnumItem), byref(fetched)) == S_OK
            and fetched.value > 0
        ):
            print("Enumerating item:")
            pszName = LPWSTR()
            hr = pEnumItem.contents.lpVtbl.contents.GetDisplayName(pEnumItem, SIGDN_NORMALDISPLAY, byref(pszName))
            if hr != S_OK:
                raise HRESULT(hr).exception("Failed to get pszName from IEnumShellItems")
            print(f" - Display name: {pszName.value}")
            comFuncs.pCoTaskMemFree(pszName)

            pszFilePath = LPWSTR()
            hr = pEnumItem.contents.lpVtbl.contents.GetDisplayName(pEnumItem, SIGDN_FILESYSPATH, byref(pszFilePath))
            if hr != S_OK:
                raise HRESULT(hr).exception("Failed to get pszFilePath from IEnumShellItems")
            print(f" - File path: {pszFilePath.value}")
            comFuncs.pCoTaskMemFree(pszFilePath)

            attributes = c_ulong()
            hr = pEnumItem.contents.lpVtbl.contents.GetAttributes(pEnumItem, SFGAO_FILESYSTEM | SFGAO_FOLDER, byref(attributes))
            if hr != S_OK:
                raise HRESULT(hr).exception("Failed to get attributes from IEnumShellItems")
            print(f" - Attributes: {attributes.value}")

            pParentItem = POINTER(IShellItem)()
            hr = pEnumItem.contents.lpVtbl.contents.GetParent(pEnumItem, byref(pParentItem))
            if hr != S_OK:
                raise HRESULT(hr).exception("Failed to GetParent from IEnumShellItems")
            if pParentItem:
                pszParentName = LPWSTR()
                hr = pParentItem.contents.lpVtbl.contents.GetDisplayName(pParentItem, SIGDN_NORMALDISPLAY, byref(pszParentName))
                if hr != S_OK:
                    raise HRESULT(hr).exception("Failed to GetDisplayName from GetParent from IEnumShellItems")
                print(f" - Parent: {pszParentName.value}")
                comFuncs.pCoTaskMemFree(pszParentName)
                pParentItem.contents.lpVtbl.contents.Release(pParentItem)

            pEnumItem.contents.lpVtbl.contents.Release(pEnumItem)

        pEnumShellItems.contents.lpVtbl.contents.Release(pEnumShellItems)

    for i in range(itemCount.value):
        pItem = POINTER(IShellItem)()
        hr = pResultsArray.contents.lpVtbl.contents.GetItemAt(pResultsArray, i, byref(pItem))
        if hr != S_OK:
            raise HRESULT(hr).exception(f"Failed to get item at index {i}")

        pszFilePath = LPWSTR()
        hr = pItem.contents.lpVtbl.contents.GetDisplayName(pItem, SIGDN_FILESYSPATH, byref(pszFilePath))
        if hr == S_OK and pszFilePath.value is not None:
            results.append(pszFilePath.value)
            print(f"Item {i} file path: {pszFilePath.value}")
            comFuncs.pCoTaskMemFree(pszFilePath)
        else:
            raise HRESULT(hr).exception(f"Failed to get file path for item {i}")

        attributes = c_ulong()
        hr = pItem.contents.lpVtbl.contents.GetAttributes(pItem, SFGAO_FILESYSTEM | SFGAO_FOLDER, byref(attributes))
        if hr != S_OK:
            raise HRESULT(hr).exception(f"Failed to get item attributes for item {i}")
        print(f"Item {i} attributes: {attributes.value}")

        pParentItem = POINTER(IShellItem)()
        hr = pItem.contents.lpVtbl.contents.GetParent(pItem, byref(pParentItem))
        if hr != S_OK:
            raise HRESULT(hr).exception(f"Failed to get item attributes for item {i}")
        if pParentItem:
            pszParentName = LPWSTR()
            hr = pParentItem.contents.lpVtbl.contents.GetDisplayName(pParentItem, SIGDN_NORMALDISPLAY, byref(pszParentName))
            if hr != S_OK:
                raise HRESULT(hr).exception(f"Failed to GetDisplayName for GetParent on item {i}")
            print(f"Item {i} parent: {pszParentName.value}")
            comFuncs.pCoTaskMemFree(pszParentName)
            pParentItem.contents.lpVtbl.contents.Release(pParentItem)

        pItem.contents.lpVtbl.contents.Release(pItem)

    pResultsArray.contents.lpVtbl.contents.Release(pResultsArray)
    return results
