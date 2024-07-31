from __future__ import annotations

import threading
import weakref

from contextlib import suppress
from ctypes import (
    POINTER,
    Structure,
    byref,
    c_int,
    c_ubyte,
    c_uint,
    c_ulong,
    c_ushort,
    c_wchar_p,
    oledll,
    windll,
)
from typing import TYPE_CHECKING, Sequence

if TYPE_CHECKING:
    from ctypes import _CData, _Pointer as PointerType

    with suppress(ImportError, ModuleNotFoundError):
        from comtypes import CoClass  # pyright: ignore[reportMissingTypeStubs, reportAttributeAccessIssue, reportMissingImports]
    from comtypes.GUID import GUID as COMTYPE_GUID  # pyright: ignore[reportMissingTypeStubs, reportMissingImports]
    from typing_extensions import Self  # pyright: ignore[reportMissingModuleSource]

if not TYPE_CHECKING:
    PointerType = POINTER(c_uint).__class__


class FDE_SHAREVIOLATION_RESPONSE(c_int):  # noqa: N801
    FDESVR_DEFAULT = 0x00000000
    FDESVR_ACCEPT = 0x00000001
    FDESVR_REFUSE = 0x00000002


FDE_OVERWRITE_RESPONSE = FDE_SHAREVIOLATION_RESPONSE


class GUID(Structure):
    _instances = weakref.WeakValueDictionary()  # Class-level dictionary to hold GUID instances
    _fields_: Sequence[ tuple[str, type[_CData]] | tuple[str, type[_CData], int] ] = [
        ("Data1", c_ulong),
        ("Data2", c_ushort),
        ("Data3", c_ushort),
        ("Data4", c_ubyte * 8),
    ]
    _lock: threading.Lock = threading.Lock()

    NULL: GUID

    def __new__(
        cls,
        d1: int
        | str
        | tuple[int, int, int, int, int, int, int, int, int, int, int, int, int, int, int, int]
        | tuple[int, int, int, int, int, int, int, int, int, int, int]
        | None = None,
        d2: int | None = None,
        d3: int | None = None,
        d4: tuple[int, int, int, int, int, int, int, int] | int | bytes | None = None,
        *args,
    ) -> Self:
        # Our class level singleton pattern fixes the following occasional error using ole32.StringFromCLSID:
        # Fatal Python error: bad ID: Allocated using API 'n', verified using API 'o'
        d1, d2, d3, d4 = cls._parse_args(d1, d2, d3, d4, *args)
        identifier = (d1, d2, d3, d4)

        with cls._lock:
            if identifier in cls._instances:
                return cls._instances[identifier]

            instance = super().__new__(cls)
            cls._instances[identifier] = instance
            super(cls, instance).__init__()

        instance.Data1 = c_ulong(d1)
        instance.Data2 = c_ushort(d2)
        instance.Data3 = c_ushort(d3)
        instance.Data4 = (c_ubyte * 8)(*d4)

        return instance

    def __init__(
        self,
        d1: int
        | str
        | tuple[int, int, int, int, int, int, int, int, int, int, int, int, int, int, int, int]
        | tuple[int, int, int, int, int, int, int, int, int, int, int]
        | None = None,
        d2: int | None = None,
        d3: int | None = None,
        d4: tuple[int, int, int, int, int, int, int, int] | int | bytes | None = None,
        *args,
    ):
        """Do nothing here to handle __init__ in our __new__ method."""

    def __repr__(self):
        return f'GUID("{self!s}")'

    def __str__(self):  # sourcery skip: remove-unreachable-code
        # return f"{{{self.Data1:08X}-{self.Data2:04X}-{self.Data3:04X}-{self.Data4[0]:02X}{self.Data4[1]:02X}-{self.Data4[2]:02X}{self.Data4[3]:02X}{self.Data4[4]:02X}{self.Data4[5]:02X}{self.Data4[6]:02X}{self.Data4[7]:02X}}}"
        # Comment out above line to use ole32.StringFromCLSID directly.
        p = c_wchar_p()
        windll.ole32.StringFromCLSID(byref(self), byref(p))
        result = p.value
        windll.ole32.CoTaskMemFree(p)
        if result is None:
            d4_hex = "".join(f"{byte:02X}" for byte in self.Data4)
            result = f"{{{self.Data1:08X}-{self.Data2:04X}-{self.Data3:04X}-{d4_hex[:4]}-{d4_hex[4:]}}}"
        return result and result.strip() or str(self.NULL)

    @classmethod
    def guid_ducktypes(cls) -> tuple[type[COMTYPE_GUID], type[Self]] | tuple[type[COMTYPE_GUID], type[GUID], type[Self]]:
        """Returns (cls, GUID, comtypes.GUID). If comtypes cannot be imported, return (cls, GUID).

        This allows duck typing and subclasses to work in isinstance checks.
        """
        COMTYPE_GUID = None
        with suppress(ImportError, ModuleNotFoundError):
            from comtypes.GUID import GUID as COMTYPE_GUID  # pyright: ignore[reportMissingTypeStubs, reportMissingImports]
        return (GUID, cls) if COMTYPE_GUID is None else (cls, GUID, COMTYPE_GUID)  # pyright: ignore[reportReturnType]

    def __bool__(self):
        return self != self.NULL

    def __eq__(self, other):
        return (
            self.Data1 == getattr(other, "Data1", None)
            and self.Data2 == getattr(other, "Data2", None)
            and self.Data3 == getattr(other, "Data3", None)
            and self.Data4 == getattr(other, "Data4", None)
        )

    def __bytes__(self) -> bytes:
        return (
            self.Data1.value.to_bytes(4, byteorder="little")
            + self.Data2.value.to_bytes(2, byteorder="little")
            + self.Data3.value.to_bytes(2, byteorder="little")
            + self.Data4
        )

    def __hash__(self):
        # We make GUID instances hashable, although ctypes.Structure instances are technically supposed to be mutable.
        return hash(bytes(self))

    def copy(self) -> Self:
        return self.__class__(str(self))

    @classmethod
    def from_progid(cls, progid_or_guid: str | CoClass | GUID) -> Self | GUID:
        """Get guid from progid. Also accepts a guid argument which will return a guid instance of that."""
        progid_or_guid = getattr(progid_or_guid, "_reg_clsid_", progid_or_guid)
        if isinstance(progid_or_guid, cls.guid_ducktypes()):
            return progid_or_guid  # pyright: ignore[reportReturnType]
        if not isinstance(progid_or_guid, str):
            raise TypeError(f"Cannot construct GUID from {progid_or_guid!r}")
        if progid_or_guid.startswith("{"):
            return cls(progid_or_guid)
        inst = cls()
        oledll.ole32.CLSIDFromProgID(str(progid_or_guid), byref(inst))
        return inst

    def as_progid(self) -> str | None:
        """Convert a GUID into a progid(human readable alternative to GUIDs)."""
        progid = c_wchar_p()
        oledll.ole32.ProgIDFromCLSID(byref(self), byref(progid))
        result = progid.value
        oledll.ole32.CoTaskMemFree(progid)
        return result

    @classmethod
    def create_new(cls) -> Self:
        """Create a brand new guid."""
        guid = cls()
        oledll.ole32.CoCreateGuid(byref(guid))
        return guid

    @classmethod
    def _parse_args(
        cls,
        d1: int
        | str
        | tuple[int, int, int, int, int, int, int, int, int, int, int, int, int, int, int, int]
        | tuple[int, int, int, int, int, int, int, int, int, int, int]
        | None = None,
        d2: int | None = None,
        d3: int | None = None,
        d4: tuple[int, int, int, int, int, int, int, int] | int | bytes | None = None,
        *args,
    ) -> tuple[int, int, int, bytes]:  # sourcery skip: low-code-quality
        # Null GUID
        if not d1 and not d2 and not d3 and not d4 and not args:
            return cls.NULL.Data1.value, cls.NULL.Data2.value, cls.NULL.Data3.value, cls.NULL.Data4.value

        if d2 is None and d3 is None and d4 is None and not args:
            if isinstance(d1, str):
                return cls.from_string(d1)
            if isinstance(d1, tuple):
                all_byte_ints = len(d1) == 16  # noqa: PLR2004
                altern_format = len(d1) == 11  # noqa: PLR2004
                if all_byte_ints or altern_format:
                    return cls._parse_args(*d1)

        if isinstance(d1, int) and isinstance(d2, int) and isinstance(d3, int):
            if isinstance(d4, bytes):
                d4 = d4[:8]
                guid_str = f"{{{d1:08X}-{d2:04X}-{d3:04X}-{d4[0]:02X}{d4[1]:02X}-{d4[2]:02X}{d4[3]:02X}{d4[4]:02X}{d4[5]:02X}{d4[6]:02X}{d4[7]:02X}}}"
                return cls.from_string(guid_str)
            if isinstance(d4, tuple) and len(d4) == 8:
                guid_str = f"{{{d1:08X}-{d2:04X}-{d3:04X}-{d4[0]:02X}{d4[1]:02X}-{d4[2]:02X}{d4[3]:02X}-{d4[4]:02X}{d4[5]:02X}{d4[6]:02X}{d4[7]:02X}}}"
                return cls.from_string(guid_str)

        if args or isinstance(d4, int):
            all_byte_ints = len(args) == 12  # noqa: PLR2004
            altern_format = len(args) == 7  # noqa: PLR2004
            if not all_byte_ints and not altern_format:
                raise ValueError(f"Incorrect arguments passed to GUID({d1}, {d2}, {d3}, {d4}, *{args})")
            if all_byte_ints:
                gd = (d1, d2, d3, d4, *args)
                guid_str = f"{{{gd[0]:08X}-{gd[1]:04X}-{gd[2]:04X}-{gd[3]:02X}{gd[4]:02X}-{gd[5]:02X}{gd[6]:02X}{gd[7]:02X}-{gd[8]:02X}{gd[9]:02X}{gd[10]:02X}{gd[11]:02X}}}"
            else:
                guid_str = f"{{{d1:08X}-{d2:04X}-{d3:04X}-{d4:02X}{args[0]:02X}-{args[1]:02X}{args[2]:02X}{args[3]:02X}{args[4]:02X}{args[5]:02X}{args[6]:02X}}}"
            return cls.from_string(guid_str)
        raise ValueError(f"Incorrect arguments passed to GUID({d1}, {d2}, {d3}, {d4}, *{args})")

    @staticmethod
    def from_string(guid_string: str) -> tuple[int, int, int, bytes]:
        hex_values = guid_string.strip("{}").split("-")
        data1 = int(hex_values[0], 16)
        data2 = int(hex_values[1], 16)
        data3 = int(hex_values[2], 16)
        data4 = bytes.fromhex(hex_values[3] + hex_values[4])
        return data1, data2, data3, data4
GUID.NULL = GUID("{00000000-0000-0000-0000-000000000000}")
