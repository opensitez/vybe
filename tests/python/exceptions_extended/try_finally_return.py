# vybe-test: python/exceptions_extended/try_finally_return
# origin: languages/python/tests/python/test_exceptions_extended.rs
# vybe-test-mode: compile

def f():
 try:
  return 1
 finally:
  return 2
