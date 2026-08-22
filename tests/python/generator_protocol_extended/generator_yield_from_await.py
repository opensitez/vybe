# vybe-test: python/generator_protocol_extended/generator_yield_from_await
# origin: languages/python/tests/python/test_generator_protocol_extended.rs
# This fixture's SUBJECT is that Python REJECTS the construct, so the file
# cannot itself be valid Python. `compile()` lets it assert the rejection
# while remaining a runnable test.
_SRC = """
async def ag():
 yield from async_iter()
"""
try:
    compile(_SRC, '<fixture>', 'exec')
except SyntaxError:
    pass
