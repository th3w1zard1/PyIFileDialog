from __future__ import annotations

from _ctypes import _Pointer
import random
import sys

from ctypes import POINTER, c_ulong, cast
from ctypes.wintypes import HWND, LPCWSTR

from interfaces import COMDLG_FILTERSPEC, IFileDialog, IFileOpenDialog, IFileSaveDialog
from windialogs import (
    FreeCOMFunctionPointers,
    LoadCOMFunctionPointers,
    configureFileDialog,
    createFileDialog,
    getFileDialogResults,
    showDialog,
)
from hresult import S_OK


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
        filters.append(COMDLG_FILTERSPEC(LPCWSTR("All Files"), LPCWSTR("*.*")))
        for i in range(numFilters):
            filterName = f"Random Filter {i + 1}"
            filterSpec = f"*.{i + 1}"
            filters.append(COMDLG_FILTERSPEC(LPCWSTR(filterName), LPCWSTR(filterSpec)))
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
            filters.append(COMDLG_FILTERSPEC(LPCWSTR(filterName), LPCWSTR(filterSpec)))
        elif choice == len(filters) + 2:
            break
        else:
            filter_index = choice - 2
            print(f"\nEditing Filter: {filters[filter_index].pszName} ({filters[filter_index].pszSpec})")
            filters[filter_index].pszName = LPCWSTR(getUserInputStr("Enter new filter name: ", filters[filter_index].pszName))
            filters[filter_index].pszSpec = LPCWSTR(getUserInputStr("Enter new filter spec: ", filters[filter_index].pszSpec))


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

        filters = []
        options = getFileDialogOptions(isSaveDialog, randomize, filters)

        comFuncs = LoadCOMFunctionPointers()
        try:
            if not all([comFuncs.pCoInitialize, comFuncs.pCoCreateInstance, comFuncs.pCoTaskMemFree, comFuncs.pCoUninitialize, comFuncs.pSHCreateItemFromParsingName]):
                raise RuntimeError("Failed to load one or more COM functions.")  # noqa: TRY301

            hr = comFuncs.pCoInitialize(None)
            if hr != S_OK:
                raise RuntimeError(f"CoInitialize failed: {hr}")  # noqa: TRY301
            pFileDialog: _Pointer[IFileOpenDialog | IFileSaveDialog | IFileDialog] = createFileDialog(comFuncs, dialogType)

            configureFileDialog(comFuncs, pFileDialog, filters, defaultFolder, options, forceFileSystem=True, allowMultiselect=True)
            showDialog(pFileDialog, HWND(0))

            if not isSaveDialog:
                results: list[str] = getFileDialogResults(comFuncs, pFileDialog)
                for filePath in results:
                    print(f"Selected file: {filePath}")

            pFileDialog.contents.lpVtbl.contents.Release(pFileDialog)
            comFuncs.pCoUninitialize()
            FreeCOMFunctionPointers(comFuncs)

            if getUserInputStr("Do you want to configure another dialog? (yes/no): ", "yes").lower() != "yes":
                break

        except Exception as ex:  # noqa: BLE001
            print(f"Error: {ex}")
            if comFuncs.pCoUninitialize:
                comFuncs.pCoUninitialize()
            FreeCOMFunctionPointers(comFuncs)
            return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
