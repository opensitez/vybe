# vybe-test: python/dataclass_enum_runtime/enum_unique_decorator
# origin: languages/python/tests/python/test_dataclass_enum_runtime.rs
# vybe-test-mode: compile

from enum import Enum, unique
@unique
class E(Enum):
 A = 1
