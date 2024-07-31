from __future__ import annotations

from _ctypes import _Pointer
import random
import sys

from ctypes import POINTER, WINFUNCTYPE, addressof, byref, c_uint, c_ulong, c_void_p, pointer, cast as cast_with_ctypes
from ctypes.wintypes import HWND
from typing import TYPE_CHECKING

from com_helpers import COMInitializeContext
from com_types import GUID
from hresult import HRESULT, HRESULTException, decode_hresult
from interfaces import COMDLG_FILTERSPEC, CLSID_FileDialog, CLSID_FileOpenDialog, CLSID_FileSaveDialog, IFileDialog, IFileOpenDialog, IFileOpenDialogVTable, IFileSaveDialog, IID_IFileDialog, IID_IFileOpenDialog, IID_IFileSaveDialog
from iunknown import LPVOID, S_OK, IID_IUnknown, IUnknown
from windialogs import (
    FreeCOMFunctionPointers,
    LoadCOMFunctionPointers,
    configureFileDialog,
    getFileDialogResults,
    showDialog,
)

if TYPE_CHECKING:
    from _ctypes import _Pointer
    from ctypes import Array

    from interfaces import COMFunctionPointers

# Enum for option states
class OptionState:
    Default = 0
    Enabled = 1
    Disabled = -1


def GetRandomNumber(low: int, high: int) -> int:
    if low > high:
        raise ValueError("Lower bound must be less than or equal to upper bound.")
    return random.randint(low, high)


def getUserInputStr(prompt: str, defaultValue: str) -> str:
    input_str = input(prompt).strip()
    return input_str or defaultValue


def getUserInputInt(prompt: str, validChoices: list[int], defaultValue: int) -> int:
    while True:
        try:
            input_str = input(prompt).strip()
            if not input_str:
                return defaultValue
            value = int(input_str)
            if value in validChoices:
                return value
            print(f"Invalid choice. Please enter one of the following choices: {validChoices}")
        except ValueError:
            print("Invalid input. Please enter a valid integer.")


def manageFilters(filters: list[COMDLG_FILTERSPEC], randomize: bool) -> None:
    if randomize:
        numFilters = GetRandomNumber(1, 25)
        filters.append(COMDLG_FILTERSPEC("All Files", "*.*"))
        for i in range(numFilters):
            filterName = f"Random Filter {i + 1}"
            filterSpec = f"*.{i + 1}"
            filters.append(COMDLG_FILTERSPEC(filterName, filterSpec))
        return

    while True:
        print("\nFile Type Filters Menu:")
        print("1. Add New Filter")
        for i, filter_spec in enumerate(filters):
            print(f"{i + 2}. Edit Filter: {filter_spec.pszName} ({filter_spec.pszSpec})")
        print(f"{len(filters) + 2}. Done")

        choice = getUserInputInt("Choose an option: ", list(range(1, len(filters) + 3)), len(filters) + 2)

        if choice == 1:
            filterName = getUserInputStr("Enter filter name: ", "Default Filter")
            filterSpec = getUserInputStr("Enter filter spec: ", "*.*")
            filters.append(COMDLG_FILTERSPEC(filterName, filterSpec))
        elif choice == len(filters) + 2:
            break
        else:
            filter_index = choice - 2
            print(f"\nEditing Filter: {filters[filter_index].pszName} ({filters[filter_index].pszSpec})")
            filters[filter_index].pszName = getUserInputStr("Enter new filter name: ", filters[filter_index].pszName)
            filters[filter_index].pszSpec = getUserInputStr("Enter new filter spec: ", filters[filter_index].pszSpec)


def configureDialogOptions(optionStates: list[int], randomize: bool) -> None:
    options = [
        ("Allow multiple selection", OptionState.Default),
        ("Do not add to recent", OptionState.Default),
        ("Show hidden files", OptionState.Default),
        ("No change dir", OptionState.Default),
        ("Confirm overwrite", OptionState.Default),
        ("Hide MRU places", OptionState.Default),
        ("Hide pinned places", OptionState.Default),
        ("Share aware", OptionState.Default)
    ]

    optionStates.extend([OptionState.Default] * len(options))  # Ensure optionStates is properly initialized

    if randomize:
        for i in range(len(options)):
            optionStates[i] = GetRandomNumber(-1, 1)
        return

    while True:
        print("\nDialog Options Menu:")
        for i, (name, _) in enumerate(options):
            state_str = {
                OptionState.Enabled: "Enabled",
                OptionState.Disabled: "Disabled",
                OptionState.Default: "Default"
            }[optionStates[i]]
            print(f"{i + 1}. {name} (Current: {state_str})")
        print(f"{len(options) + 1}. Done")

        choice = getUserInputInt("Choose an option to change or done to continue: ", list(range(1, len(options) + 2)), len(options) + 1)

        if choice == len(options) + 1:
            break
        option_index = choice - 1
        new_value = getUserInputInt(
            f"Select {options[option_index][0]} option:\n1. Enabled\n2. Disabled\n3. Default\nChoose an option: ",
            [1, 2, 3], 3
        )
        optionStates[option_index] = {1: OptionState.Enabled, 2: OptionState.Disabled, 3: OptionState.Default}[new_value]


def getFileDialogOptions(isSaveDialog: bool, randomize: bool, filters: list[COMDLG_FILTERSPEC]) -> c_ulong:
    options = 0x00000002 | 0x00000008 | 0x00000020  # FOS_FILEMUSTEXIST | FOS_PATHMUSTEXIST | FOS_FORCEFILESYSTEM

    optionStates: list[int] = []
    configureDialogOptions(optionStates, randomize)

    optionFlags: list[tuple[int, bool]] = [
        (0x00000200, False),   # FOS_ALLOWMULTISELECT
        (0x02000000, False),   # FOS_DONTADDTORECENT
        (0x10000000, False),   # FOS_FORCESHOWHIDDEN
        (0x00010000, False),   # FOS_NOCHANGEDIR
        (0x00000002, True),    # FOS_OVERWRITEPROMPT
        (0x00400000, False),   # FOS_HIDEMRUPLACES
        (0x00200000, False),   # FOS_HIDEPINNEDPLACES
        (0x00004000, False)    # FOS_SHAREAWARE
    ]

    for state, (flag, saveDialogOnly) in zip(optionStates, optionFlags):
        if state == OptionState.Enabled:
            options |= flag
        elif state == OptionState.Disabled:
            options &= ~flag

    manageFilters(filters, randomize)

    return c_ulong(options)




def main() -> int:
    while True:
        dialogType = getUserInputInt(
            "Select Dialog Type:\n1. Open File Dialog\n2. Save File Dialog\n3. Base File Dialog (parent of open/save dialogs)\n4. Randomize all options\nChoose an option: ",
            [1, 2, 3, 4], 4
        )

        randomize = (dialogType == 4)
        isSaveDialog = (dialogType == 2)

        title = getUserInputStr("Dialog title (default: My Python IFileOpenDialog): ", "My Python IFileOpenDialog") if not randomize else "My Python IFileOpenDialog"
        defaultFolder = getUserInputStr("Default folder path (default: C:): ", "C:") if not randomize else "C:"

        filters: list[COMDLG_FILTERSPEC] = []
        options: c_ulong = getFileDialogOptions(isSaveDialog, randomize, filters)

        if dialogType in (1, 4):
            clsid = CLSID_FileOpenDialog
            iid = IID_IFileOpenDialog
            _type_ = IFileOpenDialog
        elif dialogType == 2:
            clsid = CLSID_FileSaveDialog
            iid = IID_IFileSaveDialog
            _type_ = IFileSaveDialog
        elif dialogType == 3:
            clsid = CLSID_FileDialog
            iid = IID_IFileDialog
            _type_ = IFileDialog
        else:
            raise ValueError(f"Unexpected dialogType: {dialogType} (should be 1-4)")

        comFuncPtrs: COMFunctionPointers = LoadCOMFunctionPointers(_type_)
        pFileDialog: _Pointer[IFileOpenDialog | IFileSaveDialog | IFileDialog] | None = None
        try:
            if not all([comFuncPtrs.pCoInitialize, comFuncPtrs.pCoCreateInstance, comFuncPtrs.pCoTaskMemFree, comFuncPtrs.pCoUninitialize, comFuncPtrs.pSHCreateItemFromParsingName]):
                raise RuntimeError("Failed to load one or more COM functions.")  # noqa: TRY301

            hr = comFuncPtrs.pCoInitialize(None)
            if hr != S_OK:
                raise OSError(hr, decode_hresult(hr), "Failed to call CoInitialize in order to create IFileDialog")

            #pFileDialog = cast_with_ctypes(arrFileDialog, POINTER(_type_))
            #pFileDialog = POINTER(_type_)()
            fileDialog: IFileDialog | IFileOpenDialog | IFileSaveDialog = _type_()
            arrFileDialog: Array[IFileDialog] | Array[IFileOpenDialog] | Array[IFileSaveDialog] = (_type_ * 1)()
            pFileDialog = pointer(fileDialog)
            hr = comFuncPtrs.resolve_function(
                comFuncPtrs.hOle32,
                b"CoCreateInstance",
                WINFUNCTYPE(
                    HRESULT,
                    POINTER(GUID),
                    c_void_p,
                    c_ulong,
                    POINTER(GUID),
                    POINTER(_type_* 1),
                ),
            )(byref(clsid), None, 1, byref(iid), byref(arrFileDialog))
            if hr != S_OK:
                raise HRESULT(hr).exception("CoCreateInstance failed! Cannot create file dialog")
            #createFileDialog(comFuncs, cast_with_ctypes(pFileDialog, POINTER(_type_)), clsid, iid)

            addr1 = addressof(fileDialog)
            print("address of arrFileDialog[0]:", addr1)
            addr1check1 = addressof(pFileDialog.contents)
            assert addr1 == addr1check1, f"addressof(arrFileDialog[0]) != addressof(pFileDialog.contents) ({addr1} != {addr1check1})"

            #setDialogAttributes(fileDialog, title)
            configureFileDialog(comFuncPtrs, fileDialog, filters, defaultFolder, options, forceFileSystem=True, allowMultiselect=True)
            showDialog(pFileDialog, HWND(0))

            if not isSaveDialog:
                results: list[str] = getFileDialogResults(comFuncPtrs, pFileDialog)
                for filePath in results:
                    print(f"Selected file: {filePath}")

            if getUserInputStr("Do you want to configure another dialog? (yes/no): ", "yes").lower() != "yes":
                break
        except OSError as e:
            if e.winerror and not isinstance(e, HRESULTException):
                hr = HRESULT(e.winerror)
                raise hr.exception(str(e)) from e
            raise
        finally:
            hr = S_OK
            #if pFileDialog is not None:
            #    hr = pFileDialog.contents.lpVtbl.contents.Release(cast_with_ctypes(pFileDialog, POINTER(IUnknown)))
            if comFuncPtrs.pCoUninitialize and hr == S_OK:
                comFuncPtrs.pCoUninitialize()
            FreeCOMFunctionPointers(comFuncPtrs)

    return 0

if __name__ == "__main__":
    import sys
    sys.exit(main())
