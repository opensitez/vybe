# vybe-test: python/py_typing/test_py_typing_typeguard
# origin: languages/python/tests/python/test_py_typing.rs

from typing import Union

def is_string(val: Union[str, int]) -> bool:
    return isinstance(val, str)

items = [1, "hello", 2, "world"]
strings = [x for x in items if is_string(x)]
print(strings)
