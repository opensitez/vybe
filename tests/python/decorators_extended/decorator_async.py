# vybe-test: python/decorators_extended/decorator_async
# origin: languages/python/tests/python/test_decorators_extended.rs
# vybe-test-mode: compile

def deco(f):
 return f
@deco
async def ag():
 return 1
