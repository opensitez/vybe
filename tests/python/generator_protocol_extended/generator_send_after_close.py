# vybe-test: python/generator_protocol_extended/generator_send_after_close
# origin: languages/python/tests/python/test_generator_protocol_extended.rs

def g():
 yield 1
it = g()
next(it)
it.close()
try:
 it.send(1)
 print('ok')
except StopIteration:
 print('stop')
