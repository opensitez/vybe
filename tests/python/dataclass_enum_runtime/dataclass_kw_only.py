# vybe-test: python/dataclass_enum_runtime/dataclass_kw_only
# origin: languages/python/tests/python/test_dataclass_enum_runtime.rs

from dataclasses import dataclass, KW_ONLY
@dataclass(kw_only=True)
class P:
 x: int
