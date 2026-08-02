# vybe-test: python/stdlib_compile_extended/typing_literal
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

from typing import Literal
x: Literal['a', 'b'] = 'a'
