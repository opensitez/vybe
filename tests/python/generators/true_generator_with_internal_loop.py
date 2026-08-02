# vybe-test: python/generators/true_generator_with_internal_loop
# origin: languages/python/tests/python/test_generators.rs

@generator
def gen():
    i = 0
    while i < 3:
        yield i
        i = i + 1

for v in gen():
    print(v)
