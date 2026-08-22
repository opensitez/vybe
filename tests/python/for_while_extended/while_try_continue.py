# vybe-test: python/for_while_extended/while_try_continue
# origin: languages/python/tests/python/test_for_while_extended.rs

i = 0
while i < 2:
 i += 1
 try:
  continue
 finally:
  pass
