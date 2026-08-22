# vybe-test: python/closure_extended/closure_nonlocal_error
# origin: languages/python/tests/python/test_closure_extended.rs
# This fixture's SUBJECT is that Python REJECTS the construct, so the file
# cannot itself be valid Python. `compile()` lets it assert the rejection
# while remaining a runnable test.
_SRC = """
def outer():
 def inner():
  nonlocal x
"""
try:
    compile(_SRC, '<fixture>', 'exec')
except SyntaxError:
    pass
