# vybe-test: python/stdlib_compile_extended/dataclasses_asdict
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

from dataclasses import dataclass, asdict
@dataclass
class P:
 x: int
