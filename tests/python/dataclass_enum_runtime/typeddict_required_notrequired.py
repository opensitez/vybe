# vybe-test: python/dataclass_enum_runtime/typeddict_required_notrequired
# origin: languages/python/tests/python/test_dataclass_enum_runtime.rs

from typing import TypedDict, Required, NotRequired
class D(TypedDict):
 x: Required[int]
