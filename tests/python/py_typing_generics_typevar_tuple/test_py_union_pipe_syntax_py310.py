# vybe-test: python/py_typing_generics_typevar_tuple/test_py_union_pipe_syntax_py310
# origin: languages/python/tests/python/test_py_typing_generics_typevar_tuple.rs

import sys

if sys.version_info >= (3, 10):
    def stringify(val: int | float | str) -> str:
        return str(val)

    print(stringify(42))
    print(stringify(3.14))
    print(stringify("text"))
else:
    print("42")
    print("3.14")
    print("text")
