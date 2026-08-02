# vybe-test: python/python_dataclass_kwonly_slots/test_dataclass_fields_name_type_inspection
# origin: languages/python/tests/python/test_python_dataclass_kwonly_slots.rs

from dataclasses import dataclass, fields

@dataclass
class Record:
    id: int
    data: str

f_types = {f.name: f.type.__name__ for f in fields(Record)}
print(f_types["id"])
print(f_types["data"])
