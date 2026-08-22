# vybe-test: python/comprehension_walrus_spec/comp_with_walrus_filter_compile
# origin: languages/python/tests/python/test_comprehension_walrus_spec.rs

xs = [y for x in range(5) if (y := x * 2) > 2]
