"""Contains enumerations for SolidWorks Errors."""

from enum import Enum


class FileSaveError(Enum):
    FILE_LOCK_ERROR = 16
    FILE_NAME_CONTAINS_AT_SIGN = 8
    FILE_NAME_EMPTY = 4
    FILE_SAVE_AS_BAD_EDRAWINGS_VERSION = 1024
    FILE_SAVE_AS_DO_NOT_OVERWRITE = 128
    FILE_SAVE_AS_INVALID_FILE_EXTENSION = 256
    FILE_SAVE_AS_NO_SELECTION = 512
    FILE_SAVE_AS_NOT_SUPPORTED = 4096
    FILE_SAVE_FORMAT_NOT_AVAILABLE = 32
    FILE_SAVE_REQUIRES_SAVING_REFERENCES = 8192
    GENERIC_SAVE_ERROR = 1
    READ_ONLY_SAVE_ERROR = 2


class FileLoadError(Enum):
    ADDIN_INTERUPT_ERROR = 1048576
    FILE_NOT_FOUND_ERROR = 2
    FILE_WITH_SAME_TITLE_ALREADY_OPEN = 65536
    FUTURE_VERSION = 8192
    GENERIC_ERROR = 1
    INVALID_FILE_TYPE_ERROR = 1024
    LIQUID_MACHINE_DOC = 131072
    LOW_RESOURCES_ERROR = 262144
    NO_DISPLAY_DATA = 524288