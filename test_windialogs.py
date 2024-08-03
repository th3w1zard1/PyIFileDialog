from __future__ import annotations

import secrets
import sys
import typing

from ctypes import c_ulong
from ctypes.wintypes import HWND
from typing import TYPE_CHECKING

import comtypes  # pyright: ignore[reportMissingTypeStubs]
import comtypes.client  # pyright: ignore[reportMissingTypeStubs]

from hresult import HRESULT, S_FALSE, S_OK, HRESULTError, decode_hresult
from interfaces import (
    COMDLG_FILTERSPEC,
    FOS_ALLNONSTORAGEITEMS,
    FOS_ALLOWMULTISELECT,
    FOS_CREATEPROMPT,
    FOS_DEFAULTNOMINIMODE,
    FOS_DONTADDTORECENT,
    FOS_FILEMUSTEXIST,
    FOS_FORCEFILESYSTEM,
    FOS_FORCEPREVIEWPANEON,
    FOS_FORCESHOWHIDDEN,
    FOS_HIDEMRUPLACES,
    FOS_HIDEPINNEDPLACES,
    FOS_NOCHANGEDIR,
    FOS_NODEREFERENCELINKS,
    FOS_NOREADONLYRETURN,
    FOS_NOTESTFILECREATE,
    FOS_NOVALIDATE,
    FOS_OVERWRITEPROMPT,
    FOS_PATHMUSTEXIST,
    FOS_SHAREAWARE,
    FOS_STRICTFILETYPES,
    CLSID_FileOpenDialog,
    CLSID_FileSaveDialog,
    IFileOpenDialog,
    IFileSaveDialog,
)
from windialogs import (
    FreeCOMFunctionPointers,
    LoadCOMFunctionPointers,
    configureFileDialog,
    getFileOpenDialogResults,
    setDialogAttributes,
    setupFileDialogEvents,
    showDialog,
)

if TYPE_CHECKING:
    from interfaces import (
        COMFunctionPointers,
        IFileDialog,
    )
    from typing_extensions import Literal  # pyright: ignore[reportMissingModuleSource]

# Enum for option states
class OptionState:
    Default = 0
    Enabled = 1
    Disabled = -1


def GetRandomNumber(low: int, high: int) -> int:
    if low > high:
        raise ValueError("Lower bound must be less than or equal to upper bound.")
    return secrets.randbelow(high - low + 1) + low  # noqa: S311


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
        except ValueError:  # noqa: PERF203
            print("Invalid input. Please enter a valid integer.")


def printDialogOptions(options: c_ulong) -> None:
    # Map the constants to their names
    OPTIONS_FLAGS = {
        FOS_OVERWRITEPROMPT: "FOS_OVERWRITEPROMPT",
        FOS_STRICTFILETYPES: "FOS_STRICTFILETYPES",
        FOS_NOCHANGEDIR: "FOS_NOCHANGEDIR",
        FOS_FORCEFILESYSTEM: "FOS_FORCEFILESYSTEM",
        FOS_ALLNONSTORAGEITEMS: "FOS_ALLNONSTORAGEITEMS",
        FOS_NOVALIDATE: "FOS_NOVALIDATE",
        FOS_ALLOWMULTISELECT: "FOS_ALLOWMULTISELECT",
        FOS_PATHMUSTEXIST: "FOS_PATHMUSTEXIST",
        FOS_FILEMUSTEXIST: "FOS_FILEMUSTEXIST",
        FOS_CREATEPROMPT: "FOS_CREATEPROMPT",
        FOS_SHAREAWARE: "FOS_SHAREAWARE",
        FOS_NOREADONLYRETURN: "FOS_NOREADONLYRETURN",
        FOS_NOTESTFILECREATE: "FOS_NOTESTFILECREATE",
        FOS_HIDEMRUPLACES: "FOS_HIDEMRUPLACES",
        FOS_HIDEPINNEDPLACES: "FOS_HIDEPINNEDPLACES",
        FOS_NODEREFERENCELINKS: "FOS_NODEREFERENCELINKS",
        FOS_DONTADDTORECENT: "FOS_DONTADDTORECENT",
        FOS_FORCESHOWHIDDEN: "FOS_FORCESHOWHIDDEN",
        FOS_DEFAULTNOMINIMODE: "FOS_DEFAULTNOMINIMODE",
        FOS_FORCEPREVIEWPANEON: "FOS_FORCEPREVIEWPANEON",
    }
    print("Dialog options set:")
    for flag, name in OPTIONS_FLAGS.items():
        if options.value & flag:
            print(f"  - {name}")


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
        "Allow multiple selection",
        "Do not add to recent",
        "Show hidden files",
        "No change dir",
        "Confirm overwrite",
        "Hide MRU places",
        "Hide pinned places",
        "Share aware",
        "Strict file types",
        "All non-storage items",
        "No validate",
        "Create prompt",
        "No readonly return",
        "No test file create",
        "No dereference links",
        "Force preview pane on",
        "Default no mini mode"
    ]

    # Ensure optionStates is properly initialized
    optionStates.extend([OptionState.Disabled] * len(options))

    if randomize:
        for i in range(len(options)):
            optionStates[i] = GetRandomNumber(OptionState.Disabled, OptionState.Enabled)
        return

    while True:
        print("\nDialog Options Menu:")
        for i, name in enumerate(options):
            state_str = {
                OptionState.Enabled: "Enabled",
                OptionState.Disabled: "Disabled"
            }[optionStates[i]]
            print(f"{i + 1}. {name} (Current: {state_str})")
        print(f"{len(options) + 1}. Done")

        choice = getUserInputInt("Choose an option to change or done to continue: ", list(range(1, len(options) + 2)), len(options) + 1)

        if choice == len(options) + 1:
            break

        option_index = choice - 1
        new_value = getUserInputInt(
            f"Select {options[option_index]} option:\n1. Enabled\n2. Disabled\nChoose an option: ",
            [1, 2], 2
        )
        optionStates[option_index] = {1: OptionState.Enabled, 2: OptionState.Disabled}[new_value]


def getFileDialogOptions(dialogType: int, randomize: bool, filters: list[COMDLG_FILTERSPEC], defaultOptions: int) -> c_ulong:
    options = defaultOptions
    optionStates: list[int] = [
        OptionState.Enabled
        if defaultOptions & flag
        else OptionState.Disabled
        for flag in (
            FOS_OVERWRITEPROMPT, FOS_STRICTFILETYPES, FOS_NOCHANGEDIR, FOS_FORCEFILESYSTEM, FOS_ALLNONSTORAGEITEMS,
            FOS_NOVALIDATE, FOS_ALLOWMULTISELECT, FOS_PATHMUSTEXIST, FOS_FILEMUSTEXIST, FOS_CREATEPROMPT,
            FOS_SHAREAWARE, FOS_NOREADONLYRETURN, FOS_NOTESTFILECREATE, FOS_HIDEMRUPLACES, FOS_HIDEPINNEDPLACES,
            FOS_NODEREFERENCELINKS, FOS_DONTADDTORECENT, FOS_FORCESHOWHIDDEN, FOS_DEFAULTNOMINIMODE, FOS_FORCEPREVIEWPANEON,
        )
    ]
    configureDialogOptions(optionStates, randomize)

    optionFlags: list[tuple[int, bool]] = [
        (FOS_OVERWRITEPROMPT, True),
        (FOS_STRICTFILETYPES, False),
        (FOS_NOCHANGEDIR, False),
        (FOS_FORCEFILESYSTEM, False),
        (FOS_ALLNONSTORAGEITEMS, False),
        (FOS_NOVALIDATE, False),
        (FOS_ALLOWMULTISELECT, False),
        (FOS_PATHMUSTEXIST, False),
        (FOS_FILEMUSTEXIST, False),
        (FOS_CREATEPROMPT, False),
        (FOS_SHAREAWARE, False),
        (FOS_NOREADONLYRETURN, False),
        (FOS_NOTESTFILECREATE, False),
        (FOS_HIDEMRUPLACES, False),
        (FOS_HIDEPINNEDPLACES, False),
        (FOS_NODEREFERENCELINKS, False),
        (FOS_DONTADDTORECENT, False),
        (FOS_FORCESHOWHIDDEN, False),
        (FOS_DEFAULTNOMINIMODE, False),
        (FOS_FORCEPREVIEWPANEON, False),
    ]

    isSaveDialog = dialogType == 2  # noqa: PLR2004
    for state, (flag, saveDialogOnly) in zip(optionStates, optionFlags):
        if state == OptionState.Enabled:
            if not saveDialogOnly or (saveDialogOnly and isSaveDialog):  # Check for save dialog
                options |= flag
        elif state == OptionState.Disabled and (not saveDialogOnly or (saveDialogOnly and isSaveDialog)):  # Check for save dialog
            options &= ~flag

    manageFilters(filters, randomize)

    return c_ulong(options)


def main() -> int:
    while True:
        firstChoice: Literal[1, 2, 3, 4] = getUserInputInt(  # pyright: ignore[reportAssignmentType]
            "Select Dialog Type:\n1. Open File Dialog\n2. Save File Dialog\n3. Base File Dialog (parent of open/save dialogs)\n4. Randomize all options\nChoose an option: ",
            [1, 2, 3, 4], 4
        )

        randomize = (firstChoice == 4)

        title = getUserInputStr("Dialog title (default: My Python IFileOpenDialog): ", "My Python IFileOpenDialog") if not randomize else "My Python IFileOpenDialog"
        defaultFolder = getUserInputStr("Default folder path (default: C:): ", "C:") if not randomize else ""

        dialogType = firstChoice  # pyright: ignore[reportAssignmentType]
        if randomize:
            dialogType: Literal[1, 2, 3] = GetRandomNumber(1, 3)  # pyright: ignore[reportAssignmentType]

        filters: list[COMDLG_FILTERSPEC] = []

        if dialogType == 1:
            clsid = CLSID_FileOpenDialog
            _type_ = IFileOpenDialog
            print("Using IFileOpenDialog")
        elif dialogType == 2:
            clsid = CLSID_FileSaveDialog
            _type_ = IFileSaveDialog
            print("Using IFileSaveDialog")
        else:
            raise ValueError(f"Unexpected dialogType: {dialogType} (should be 1-4)")

        comFuncPtrs: COMFunctionPointers = LoadCOMFunctionPointers(_type_)
        fileDialog: IFileOpenDialog | IFileSaveDialog | IFileDialog | None = None
        try:
            if not all([comFuncPtrs.pCoInitialize, comFuncPtrs.pCoCreateInstance, comFuncPtrs.pCoTaskMemFree, comFuncPtrs.pCoUninitialize, comFuncPtrs.pSHCreateItemFromParsingName]):
                raise RuntimeError("Failed to load one or more COM functions.")  # noqa: TRY301

            hr = comFuncPtrs.pCoInitialize(None)
            if hr not in (S_OK, S_FALSE):
                raise OSError(hr, decode_hresult(hr), "Failed to call CoInitialize in order to create IFileDialog")  # noqa: TRY301

            fileDialog = comtypes.client.CreateObject(clsid, interface=_type_)
            HRESULT.raise_for_status(hr, "CoCreateInstance failed! Cannot create file dialog", ignore_s_false=True)

            # Retrieve and print default options
            default_options = fileDialog.GetOptions()
            print("Default dialog options:")
            printDialogOptions(c_ulong(default_options))
            options: c_ulong = getFileDialogOptions(dialogType, randomize, filters, default_options)
            configureFileDialog(comFuncPtrs, fileDialog, filters, defaultFolder, options)
            cookie: int = setupFileDialogEvents(fileDialog)
            print(f"Dialog events setup, cookie is {cookie}")
            setDialogAttributes(fileDialog, title, "Some Button", "Yoooooooooo")
            showDialog(fileDialog, HWND(0))

            if dialogType == 1:
                results: list[str] = getFileOpenDialogResults(comFuncPtrs, typing.cast(IFileOpenDialog, fileDialog))
                for filePath in results:
                    print(f"Selected file: {filePath}")
            fileDialog.Unadvise(cookie)
            if getUserInputStr("Do you want to configure another dialog? (yes/no): ", "yes").lower() != "yes":
                break
        except OSError as e:
            if e.winerror and not isinstance(e, HRESULTError):
                hr = HRESULT(e.winerror)
                raise hr.exception(str(e)) from e
            raise
        finally:
            hr = S_OK
            if fileDialog is not None:
                hr = fileDialog.Release()
            if comFuncPtrs.pCoUninitialize and hr == S_OK:
                comFuncPtrs.pCoUninitialize()
            FreeCOMFunctionPointers(comFuncPtrs)

    return 0

if __name__ == "__main__":
    import sys
    sys.exit(main())
