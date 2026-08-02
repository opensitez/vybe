# vybe-test: python/stdlib_compile_extended/typing_no_return
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

from typing import NoReturn
def f() -> NoReturn:
 raise SystemExit()
