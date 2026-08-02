# vybe-test: python/generator_protocol_extended/generator_send_not_started
# origin: languages/python/tests/python/test_generator_protocol_extended.rs

def g():
 x = yield 1
 yield x
try:
 g().send(1)
 print('ok')
except TypeError:
 print('err')
