# vybe-test: python/stdlib_compile_extended/enum_auto
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

from enum import Enum, auto
class E(Enum):
 A = auto()
