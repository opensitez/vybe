# vybe-test: python/stdlib_compile_extended/dataclasses_frozen
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

from dataclasses import dataclass
@dataclass(frozen=True)
class P:
 x: int
