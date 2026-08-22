# vybe-test: python/stdlib_compile_extended/enum_flag
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs

from enum import Flag, auto
class F(Flag):
 A = auto()
