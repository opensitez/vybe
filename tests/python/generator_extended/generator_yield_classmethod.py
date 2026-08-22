# vybe-test: python/generator_extended/generator_yield_classmethod
# origin: languages/python/tests/python/test_generator_extended.rs

class C:
 @classmethod
 def f(cls):
  yield cls
list(C.f())
