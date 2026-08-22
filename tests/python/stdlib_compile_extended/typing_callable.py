# vybe-test: python/stdlib_compile_extended/typing_callable
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs

from typing import Callable
def f(x: Callable[[int], int]) -> int:
 return x(1)
