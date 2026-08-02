# vybe-test: python/py_meta_class_creation_hooks/test_py_metaclass_prepare_ordered_dict
# origin: languages/python/tests/python/test_py_meta_class_creation_hooks.rs

class OrderedMeta(type):
    @classmethod
    def __prepare__(mcs, name, bases):
        return {"_field_order": []}

    def __new__(mcs, name, bases, attrs):
        for k in attrs:
            if not k.startswith("__"):
                attrs["_field_order"].append(k)
        return super().__new__(mcs, name, bases, attrs)

class Model(metaclass=OrderedMeta):
    id = 1
    name = "test"
    active = True

print(Model._field_order)
