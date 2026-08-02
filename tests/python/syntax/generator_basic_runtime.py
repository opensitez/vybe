# vybe-test: python/syntax/generator_basic_runtime
# origin: languages/python/tests/python/test_syntax.rs

def gen():
    yield 1
    yield 2
    yield 3
for v in gen():
    print(v)
