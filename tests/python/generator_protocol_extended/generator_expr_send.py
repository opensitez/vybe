# vybe-test: python/generator_protocol_extended/generator_expr_send
# origin: languages/python/tests/python/test_generator_protocol_extended.rs

g = (x for x in range(2))
print(next(g))
