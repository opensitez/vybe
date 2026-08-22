# vybe-test: python/decorators_extended/decorator_async
# origin: languages/python/tests/python/test_decorators_extended.rs

def deco(f):
 return f
@deco
async def ag():
 return 1
