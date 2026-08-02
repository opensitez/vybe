# vybe-test: python/inheritance_core/inheritance_mro_order
# origin: languages/python/tests/python/test_inheritance_core.rs

class B:
 pass
class D(B):
 pass
print([c.__name__ for c in D.__mro__])
