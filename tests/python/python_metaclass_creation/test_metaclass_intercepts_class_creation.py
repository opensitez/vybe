# vybe-test: python/python_metaclass_creation/test_metaclass_intercepts_class_creation
# origin: languages/python/tests/python/test_python_metaclass_creation.rs

class UpperMeta(type):
    def __new__(mcs, name, bases, namespace):
        upper_ns = {k.upper(): v for k, v in namespace.items() if not k.startswith("__")}
        upper_ns.update({k: v for k, v in namespace.items() if k.startswith("__")})
        return super().__new__(mcs, name, bases, upper_ns)

class MyClass(metaclass=UpperMeta):
    value = 42

print(MyClass.VALUE)
