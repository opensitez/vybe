# vybe-test: python/stdlib_compile_extended/dataclasses_field
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

from dataclasses import dataclass, field
@dataclass
class P:
 xs: list = field(default_factory=list)
