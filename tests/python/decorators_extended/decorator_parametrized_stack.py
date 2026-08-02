# vybe-test: python/decorators_extended/decorator_parametrized_stack
# origin: languages/python/tests/python/test_decorators_extended.rs
# vybe-test-mode: compile

def a(x):
 def deco(f):
  return f
 return deco
@a(1)
@a(2)
def f():
 pass
