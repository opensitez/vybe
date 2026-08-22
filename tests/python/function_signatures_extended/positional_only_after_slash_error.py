# vybe-test: python/function_signatures_extended/positional_only_after_slash_error
# origin: languages/python/tests/python/test_function_signatures_extended.rs
# This fixture's SUBJECT is that Python REJECTS the construct, so the file
# cannot itself be valid Python. `compile()` lets it assert the rejection
# while remaining a runnable test.
_SRC = """
def f(a, /, /, b): pass
"""
try:
    compile(_SRC, '<fixture>', 'exec')
except SyntaxError:
    pass
