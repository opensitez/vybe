# vybe-test: python/python_super_mro/test_mro_diamond
# origin: languages/python/tests/python/test_python_super_mro.rs

class A:
    def who(self):
        return "A"

class B(A):
    def who(self):
        return "B->" + super().who()

class C(A):
    def who(self):
        return "C->" + super().who()

class D(B, C):
    def who(self):
        return "D->" + super().who()

d = D()
print(d.who())
print([cls.__name__ for cls in D.__mro__])
