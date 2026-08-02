# vybe-test: python/stdlib_compile_extended/typing_newtype
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

from typing import NewType
UserId = NewType('UserId', int)
