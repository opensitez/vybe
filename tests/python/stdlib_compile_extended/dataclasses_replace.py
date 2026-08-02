# vybe-test: python/stdlib_compile_extended/dataclasses_replace
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

from dataclasses import dataclass, replace
@dataclass
class P:
 x: int
