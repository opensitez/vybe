# vybe-test: python/function_signatures_spec/decorator_factory_compile
# origin: languages/python/tests/python/test_function_signatures_spec.rs
# vybe-test-mode: compile

def deco(name):
    def wrap(fn):
        return fn
    return wrap
@deco('x')
def f():
    pass
