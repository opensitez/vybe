# vybe-test: python/generator_protocol_extended/generator_throw_uncaught
# origin: languages/python/tests/python/test_generator_protocol_extended.rs

def g():
 yield 1
try:
 g().throw(ValueError)
 print('ok')
except ValueError:
 print('err')
