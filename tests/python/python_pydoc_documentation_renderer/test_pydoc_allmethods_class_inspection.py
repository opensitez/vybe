# vybe-test: python/python_pydoc_documentation_renderer/test_pydoc_allmethods_class_inspection
# origin: languages/python/tests/python/test_python_pydoc_documentation_renderer.rs

import pydoc

class A:
    def method_a(self): pass

class B(A):
    def method_b(self): pass

methods = pydoc.allmethods(B)
names = [m.__name__ for m in methods]
print("method_a" in names and "method_b" in names)
