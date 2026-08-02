# vybe-test: python/py_dataclasses/test_py_dataclass_fields_introspection
# origin: languages/python/tests/python/test_py_dataclasses.rs

from dataclasses import dataclass, fields

@dataclass
class Model:
    name: str
    score: float = 0.0

for f in fields(Model):
    print(f.name, f.type, f.default)
