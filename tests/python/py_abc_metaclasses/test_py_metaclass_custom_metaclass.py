# vybe-test: python/py_abc_metaclasses/test_py_metaclass_custom_metaclass
# origin: languages/python/tests/python/test_py_abc_metaclasses.rs

class UpperAttrMeta(type):
    def __new__(mcs, name, bases, namespace):
        upper_attrs = {
            k.upper() if not k.startswith('_') else k: v
            for k, v in namespace.items()
        }
        return super().__new__(mcs, name, bases, upper_attrs)

class MyClass(metaclass=UpperAttrMeta):
    greeting = "hello"
    count = 42

print(hasattr(MyClass, 'GREETING'))
print(MyClass.GREETING)
print(MyClass.COUNT)
