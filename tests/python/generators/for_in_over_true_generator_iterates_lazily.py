# vybe-test: python/generators/for_in_over_true_generator_iterates_lazily
# origin: languages/python/tests/python/test_generators.rs

@generator
def count_to_three():
    yield 1
    yield 2
    yield 3

for v in count_to_three():
    print(v)
