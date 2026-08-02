# vybe-test: python/generators/true_generator_with_bound_args_yields_in_order
# origin: languages/python/tests/python/test_generators.rs

@generator
def ramp(n):
    i = 0
    while i < n:
        yield i
        i = i + 1

for v in ramp(4):
    print(v)
