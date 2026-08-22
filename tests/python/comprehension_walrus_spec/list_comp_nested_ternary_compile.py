# vybe-test: python/comprehension_walrus_spec/list_comp_nested_ternary_compile
# origin: languages/python/tests/python/test_comprehension_walrus_spec.rs

xs = ['big' if x > 5 else 'small' for x in range(10)]
