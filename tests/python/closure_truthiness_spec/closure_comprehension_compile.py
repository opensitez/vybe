# vybe-test: python/closure_truthiness_spec/closure_comprehension_compile
# origin: languages/python/tests/python/test_closure_truthiness_spec.rs

funcs = [lambda y, x=x: x + y for x in range(3)]
