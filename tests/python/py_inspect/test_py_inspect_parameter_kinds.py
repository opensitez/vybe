# vybe-test: python/py_inspect/test_py_inspect_parameter_kinds
# origin: languages/python/tests/python/test_py_inspect.rs

import inspect

def complex_func(pos_only, /, regular, *args, kw_only, **kwargs):
    pass

sig = inspect.signature(complex_func)
for name, param in sig.parameters.items():
    print(f"{name}: {param.kind.name}")
