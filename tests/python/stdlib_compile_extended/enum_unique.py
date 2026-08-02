# vybe-test: python/stdlib_compile_extended/enum_unique
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

from enum import Enum, unique
@unique
class E(Enum):
 A = 1
