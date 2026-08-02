# vybe-test: python/generators/true_generator_with_bound_args_yields_once
# origin: languages/python/tests/python/test_generators.rs

@generator
def ramp(n):
    yield n

for v in ramp(4):
    print(v)
