# vybe-test: python/generator_extended/generator_yield_classmethod
# origin: languages/python/tests/python/test_generator_extended.rs
# vybe-test-mode: compile

class C:
 @classmethod
 def f(cls):
  yield cls
list(C.f())
