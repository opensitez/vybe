# vybe-test: python/py_dataclass_advanced_features/test_py_dataclass_fields_inspection_list
# origin: languages/python/tests/python/test_py_dataclass_advanced_features.rs

from dataclasses import dataclass, fields

@dataclass
class Product:
    id: int
    name: str
    price: float = 0.0

f_names = [f.name for f in fields(Product)]
print(f_names)
