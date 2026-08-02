# vybe-test: python/stdlib_compile_extended/typing_protocol
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

from typing import Protocol
class P(Protocol):
 def m(self) -> int: ...
