# vybe-test: python/py_class_inheritance_mro/test_py_diamond_inheritance_mro
# origin: languages/python/tests/python/test_py_class_inheritance_mro.rs

class A:
    def who(self): return "A"

class B(A):
    def who(self): return "B->" + super().who()

class C(A):
    def who(self): return "C->" + super().who()

class D(B, C):
    def who(self): return "D->" + super().who()

d = D()
print(d.who())
mro_names = [cls.__name__ for cls in D.__mro__]
print(mro_names)
