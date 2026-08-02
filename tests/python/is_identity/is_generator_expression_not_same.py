# vybe-test: python/is_identity/is_generator_expression_not_same
# origin: languages/python/tests/python/test_is_identity.rs

print((x for x in range(1)) is (x for x in range(1)))
