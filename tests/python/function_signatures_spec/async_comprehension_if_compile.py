# vybe-test: python/function_signatures_spec/async_comprehension_if_compile
# origin: languages/python/tests/python/test_function_signatures_spec.rs
# This fixture's SUBJECT is that Python REJECTS the construct, so the file
# cannot itself be valid Python. `compile()` lets it assert the rejection
# while remaining a runnable test.
_SRC = """
result = [x async for x in aiter if x > 0]
"""
try:
    compile(_SRC, '<fixture>', 'exec')
except SyntaxError:
    pass
