# vybe-test: python/descriptor_metaclass_extended/dataclass_with_slots
# origin: languages/python/tests/python/test_descriptor_metaclass_extended.rs

from dataclasses import dataclass
@dataclass(slots=True)
class P:
 x: int
