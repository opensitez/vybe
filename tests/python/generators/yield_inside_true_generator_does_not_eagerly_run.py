# vybe-test: python/generators/yield_inside_true_generator_does_not_eagerly_run
# origin: languages/python/tests/python/test_generators.rs

@generator
def gen():
    print("bad: generator body ran without resume")
    yield 1

_ = gen()
print("ok")
