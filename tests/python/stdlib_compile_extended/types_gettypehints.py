# vybe-test: python/stdlib_compile_extended/types_gettypehints
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

import typing
from typing import get_type_hints
def f(x: int) -> str: pass
get_type_hints(f)
