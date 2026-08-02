# vybe-test: python/comprehension_walrus_spec/comp_with_condition_expr_compile
# origin: languages/python/tests/python/test_comprehension_walrus_spec.rs
# vybe-test-mode: compile

xs = [x if x > 1 else 0 for x in range(4)]
