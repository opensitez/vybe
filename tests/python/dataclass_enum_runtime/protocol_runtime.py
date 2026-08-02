# vybe-test: python/dataclass_enum_runtime/protocol_runtime
# origin: languages/python/tests/python/test_dataclass_enum_runtime.rs
# vybe-test-mode: compile

from typing import Protocol
class P(Protocol):
 def m(self) -> int: ...
