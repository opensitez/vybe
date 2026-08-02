# vybe-test: python/generators/default_yield_preserves_eager_list_semantics
# origin: languages/python/tests/python/test_generators.rs

def gen():
    yield 1
    yield 2
    yield 3

for v in gen():
    print(v)
