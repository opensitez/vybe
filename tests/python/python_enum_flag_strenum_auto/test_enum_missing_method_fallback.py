# vybe-test: python/python_enum_flag_strenum_auto/test_enum_missing_method_fallback
# origin: languages/python/tests/python/test_python_enum_flag_strenum_auto.rs

from enum import Enum

class CaseInsensitiveEnum(Enum):
    FOO = 1
    BAR = 2

    @classmethod
    def _missing_(cls, value):
        if isinstance(value, str):
            for member in cls:
                if member.name == value.upper():
                    return member
        return None

print(CaseInsensitiveEnum("foo") == CaseInsensitiveEnum.FOO)
