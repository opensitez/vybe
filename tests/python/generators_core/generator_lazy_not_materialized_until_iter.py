# vybe-test: python/generators_core/generator_lazy_not_materialized_until_iter
# origin: languages/python/tests/python/test_generators_core.rs

def g():
 yield 1
gen = (x for x in g())
print(next(gen))
