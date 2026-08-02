# vybe-test: python/py_inspect/test_py_inspect_getmembers
# origin: languages/python/tests/python/test_py_inspect.rs

import inspect

class MyClass:
    class_var = 42

    def method(self):
        pass

    @classmethod
    def class_method(cls):
        pass

    @staticmethod
    def static_method():
        pass

methods = [name for name, m in inspect.getmembers(MyClass, inspect.isfunction)]
print("method" in methods)
print("class_var" not in methods)
