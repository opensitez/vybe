# vybe-test: python/stdlib_compile_extended/typing_classvar
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs

from typing import ClassVar
class C:
 x: ClassVar[int] = 1
