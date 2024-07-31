from __future__ import annotations

from _ctypes import _Pointer
from ctypes import POINTER, Structure, byref, c_long, c_ulong, c_wchar_p, pointer, windll
from ctypes.wintypes import DWORD, LPWSTR
from typing import TYPE_CHECKING, Sequence

from interfaces import (
    COMDLG_FILTERSPEC,
    E_NOINTERFACE,
    SFGAO_FILESYSTEM,
    SFGAO_FOLDER,
    SIGDN_FILESYSPATH,
    SIGDN_NORMALDISPLAY,
    CLSID_FileDialog,
    CLSID_FileOpenDialog,
    CLSID_FileSaveDialog,
    COMFunctionPointers,
    IEnumShellItems,
    IFileDialog,
    IFileDialogEvents,
    IFileDialogEventsVTable,
    IFileOpenDialog,
    IFileSaveDialog,
    IID_IFileDialog,
    IID_IFileOpenDialog,
    IID_IFileSaveDialog,
    IShellItem,
    IShellItemArray,
    IUnknown,
    SHCreateItemFromParsingName,
)
from hresult import S_OK, decode_hresult

if TYPE_CHECKING:
    from ctypes import _CData, _Pointer
    from ctypes.wintypes import HWND, LPCWSTR



# Helper to convert std::wstring to LPCWSTR
def string_to_LPCWSTR(s: str) -> LPCWSTR:
    return c_wchar_p(s)


# Helper to convert LPCWSTR to std::wstring
def LPCWSTR_to_string(s: LPCWSTR) -> str | None:
    return s.value


# Load COM function pointers
def LoadCOMFunctionPointers() -> COMFunctionPointers:
    comFuncPtrs = COMFunctionPointers()

    return comFuncPtrs


# Free COM libraries
def FreeCOMFunctionPointers(comFuncPtrs: COMFunctionPointers):  # noqa: N803
    windll.kernel32.FreeLibrary()



# FileDialogEventHandler class
class FileDialogEventHandler(Structure):
    _fields_: Sequence[tuple[str, type[_CData]] | tuple[str, type[_CData], int]] = [("lpVtbl", POINTER(IFileDialogEvents))]

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
        self.lpVtbl = pointer(self.vtable)
        self.refCount = c_long(1)

    def QueryInterface(self, riid, ppv):
        if riid.contents in (IUnknown._iid_, IFileDialogEvents._iid_):
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
    dialogType: int,  # noqa: N803
) -> _Pointer[IFileOpenDialog | IFileSaveDialog | IFileDialog]:  # noqa: N803
    if dialogType in (1, 4):
        fileDialog = IFileOpenDialog()
        iid = IID_IFileOpenDialog
        clsid = CLSID_FileOpenDialog
    elif dialogType == 2:
        fileDialog = IFileSaveDialog()
        iid = IID_IFileSaveDialog
        clsid = CLSID_FileSaveDialog
    elif dialogType == 3:
        fileDialog = IFileDialog()
        iid = IID_IFileDialog
        clsid = CLSID_FileDialog
    else:
        raise ValueError(f"Wrong dialogType: {dialogType}")
    pFileDialog = pointer(fileDialog)
    #arg = cast_with_ctypes(ppFileDialog, POINTER(c_void_p))  # ???
    hr = comFuncs.pCoCreateInstance(byref(clsid), None, 1, byref(iid), pFileDialog)
    if hr != S_OK:
        print(decode_hresult(hr))
        raise RuntimeError("Failed to create file dialog")
    return pFileDialog


def showDialog(
    pFileOpenDialog: _Pointer[ IFileOpenDialog | IFileSaveDialog | IFileDialog ],  # noqa: N803
    hwndOwner: HWND,  # noqa: N803
):
    hr = pFileOpenDialog.contents.lpVtbl.contents.Show(pFileOpenDialog, hwndOwner)
    if hr != S_OK:
        print("Failed to show the file open dialog!")
        print(decode_hresult(hr))

def createShellItem(comFuncs: COMFunctionPointers, path: str) -> _Pointer[IShellItem]:  # noqa: N803, ARG001
    shell_item = POINTER(IShellItem)()
    hr = SHCreateItemFromParsingName(path, None, IShellItem._iid_, byref(shell_item))
    if hr != S_OK:
        print(decode_hresult(hr))
        raise ValueError(f"Failed to create shell item from path: {path}")
    return shell_item


# Function to configure file dialog
def configureFileDialog(  # noqa: PLR0913
    comFuncs: COMFunctionPointers,  # noqa: N803
    pFileDialog: _Pointer[IFileOpenDialog | IFileSaveDialog | IFileDialog],  # noqa: N803
    filters: list[COMDLG_FILTERSPEC],  # noqa: N803
    defaultFolder: str,  # noqa: N803
    options: c_ulong,
    forceFileSystem: bool = False,  # noqa: N803, FBT001, FBT002
    allowMultiselect: bool = False,  # noqa: N803, FBT001, FBT002
):
    if filters:
        filter_array = (COMDLG_FILTERSPEC * len(filters))()
        for i, dialogFilter in enumerate(filters):
            filter_array[i].pszName = dialogFilter.pszName
            filter_array[i].pszSpec = dialogFilter.pszSpec

        hr = pFileDialog.contents.lpVtbl.contents.SetFileTypes(pFileDialog, len(filters), filter_array)
        if hr != S_OK:
            print("Failed to set file types")
            print(decode_hresult(hr))
            return

    if defaultFolder:
        pFolder: _Pointer[IShellItem] = createShellItem(comFuncs, defaultFolder)
        if pFolder:
            hr = pFileDialog.contents.lpVtbl.contents.SetFolder(pFileDialog, pFolder)
            if hr != S_OK:
                print("Failed to set default folder")
                print(decode_hresult(hr))
                pFolder.contents.lpVtbl.contents.Release(pFolder)
                return
            pFolder.contents.lpVtbl.contents.Release(pFolder)

    if forceFileSystem:
        options = c_ulong(options.value | 0x00000040)  # FOS_FORCEFILESYSTEM

    if allowMultiselect:
        options = c_ulong(options.value | 0x00000200)  # FOS_ALLOWMULTISELECT

    hr = pFileDialog.contents.lpVtbl.contents.SetOptions(pFileDialog, options)
    if hr != S_OK:
        print("Failed to set dialog options")
        print(decode_hresult(hr))


def getFileDialogResults(  # noqa: C901, PLR0912, PLR0915
    comFuncs: COMFunctionPointers,  # noqa: N803
    pFileOpenDialog: _Pointer[IFileOpenDialog | IFileSaveDialog | IFileDialog],  # noqa: N803
) -> list[str]:
    results = []
    pResultsArray = POINTER(IShellItemArray)()
    hr = pFileOpenDialog.contents.lpVtbl.contents.GetResults(pFileOpenDialog, byref(pResultsArray))
    if hr != S_OK:
        print(f"Failed to get dialog results. HRESULT: {hr}")
        comFuncs.pCoUninitialize()
        return results

    itemCount = DWORD()
    hr = pResultsArray.contents.lpVtbl.contents.GetCount(pResultsArray, byref(itemCount))
    if hr != S_OK:
        print(f"Failed to get item count. HRESULT: {hr}")
        pResultsArray.contents.lpVtbl.contents.Release(pResultsArray)
        comFuncs.pCoUninitialize()
        return results

    print(f"Number of items selected: {itemCount.value}")

    pEnumShellItems = POINTER(IEnumShellItems)()
    hr = pResultsArray.contents.lpVtbl.contents.EnumItems(pResultsArray, byref(pEnumShellItems))
    if hr == S_OK and pEnumShellItems:
        pEnumItem = POINTER(IShellItem)()
        fetched = DWORD()
        while pEnumShellItems.contents.lpVtbl.contents.Next(pEnumShellItems, 1, byref(pEnumItem), byref(fetched)) == S_OK and fetched.value > 0:
            print("Enumerating item:")
            pszName = LPWSTR()
            hr = pEnumItem.contents.lpVtbl.contents.GetDisplayName(pEnumItem, SIGDN_NORMALDISPLAY, byref(pszName))
            if hr == S_OK:
                print(f" - Display name: {pszName.value}")
                comFuncs.pCoTaskMemFree(pszName)

            pszFilePath = LPWSTR()
            hr = pEnumItem.contents.lpVtbl.contents.GetDisplayName(pEnumItem, SIGDN_FILESYSPATH, byref(pszFilePath))
            if hr == S_OK:
                print(f" - File path: {pszFilePath.value}")
                comFuncs.pCoTaskMemFree(pszFilePath)

            attributes = c_ulong()
            hr = pEnumItem.contents.lpVtbl.contents.GetAttributes(pEnumItem, SFGAO_FILESYSTEM | SFGAO_FOLDER, byref(attributes))
            if hr == S_OK:
                print(f" - Attributes: {attributes.value}")

            pParentItem = POINTER(IShellItem)()
            hr = pEnumItem.contents.lpVtbl.contents.GetParent(pEnumItem, byref(pParentItem))
            if hr == S_OK and pParentItem:
                pszParentName = LPWSTR()
                hr = pParentItem.contents.lpVtbl.contents.GetDisplayName(pParentItem, SIGDN_NORMALDISPLAY, byref(pszParentName))
                if hr == S_OK:
                    print(f" - Parent: {pszParentName.value}")
                    comFuncs.pCoTaskMemFree(pszParentName)
                pParentItem.contents.lpVtbl.contents.Release(pParentItem)

            pEnumItem.contents.lpVtbl.contents.Release(pEnumItem)

        pEnumShellItems.contents.lpVtbl.contents.Release(pEnumShellItems)

    for i in range(itemCount.value):
        pItem = POINTER(IShellItem)()
        hr = pResultsArray.contents.lpVtbl.contents.GetItemAt(pResultsArray, i, byref(pItem))
        if hr != S_OK:
            print(f"Failed to get item at index {i}. HRESULT: {hr}")
            continue

        pszFilePath = LPWSTR()
        hr = pItem.contents.lpVtbl.contents.GetDisplayName(pItem, SIGDN_FILESYSPATH, byref(pszFilePath))
        if hr == S_OK:
            results.append(pszFilePath.value)
            print(f"Item {i} file path: {pszFilePath.value}")
            comFuncs.pCoTaskMemFree(pszFilePath)
        else:
            print(f"Failed to get file path for item {i}. HRESULT: {hr}")

        attributes = c_ulong()
        hr = pItem.contents.lpVtbl.contents.GetAttributes(pItem, SFGAO_FILESYSTEM | SFGAO_FOLDER, byref(attributes))
        if hr == S_OK:
            print(f"Item {i} attributes: {attributes.value}")

        pParentItem = POINTER(IShellItem)()
        hr = pItem.contents.lpVtbl.contents.GetParent(pItem, byref(pParentItem))
        if hr == S_OK and pParentItem:
            pszParentName = LPWSTR()
            hr = pParentItem.contents.lpVtbl.contents.GetDisplayName(pParentItem, SIGDN_NORMALDISPLAY, byref(pszParentName))
            if hr == S_OK:
                print(f"Item {i} parent: {pszParentName.value}")
                comFuncs.pCoTaskMemFree(pszParentName)
            pParentItem.contents.lpVtbl.contents.Release(pParentItem)

        pItem.contents.lpVtbl.contents.Release(pItem)

    pResultsArray.contents.lpVtbl.contents.Release(pResultsArray)
    return results
