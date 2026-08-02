# vybe-test: python/py_inspect_reflection_introspection/test_py_inspect_getmembers_predicates
# origin: languages/python/tests/python/test_py_inspect_reflection_introspection.rs

import inspect

class MyClass:
    class_var = 42
    def method(self): pass

funcs = [name for name, _ in inspect.getmembers(MyClass, inspect.isfunction)]
print(funcs)
