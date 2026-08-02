# vybe-test: python/stdlib_compile_extended/typing_cast
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

from typing import cast
x = cast(int, '1')
