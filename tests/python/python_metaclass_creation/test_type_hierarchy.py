# vybe-test: python/python_metaclass_creation/test_type_hierarchy
# origin: languages/python/tests/python/test_python_metaclass_creation.rs

class A:
    pass

class B(A):
    pass

class C(B):
    pass

print([t.__name__ for t in C.__mro__])
