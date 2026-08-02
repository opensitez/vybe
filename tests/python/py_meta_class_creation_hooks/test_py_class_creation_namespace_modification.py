# vybe-test: python/py_meta_class_creation_hooks/test_py_class_creation_namespace_modification
# origin: languages/python/tests/python/test_py_meta_class_creation_hooks.rs

class AutoPropertyMeta(type):
    def __new__(mcs, name, bases, attrs):
        for k, v in list(attrs.items()):
            if k.startswith("get_"):
                prop_name = k[4:]
                attrs[prop_name] = property(v)
        return super().__new__(mcs, name, bases, attrs)

class User(metaclass=AutoPropertyMeta):
    def __init__(self, name):
        self._name = name
    def get_name(self):
        return self._name

u = User("Alice")
print(u.name)
