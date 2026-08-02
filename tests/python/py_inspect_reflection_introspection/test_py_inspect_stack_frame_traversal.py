# vybe-test: python/py_inspect_reflection_introspection/test_py_inspect_stack_frame_traversal
# origin: languages/python/tests/python/test_py_inspect_reflection_introspection.rs

import inspect

def level2():
    stack = inspect.stack()
    return [frame.function for frame in stack[:3]]

def level1():
    return level2()

funcs = level1()
print("level2" in funcs and "level1" in funcs)
