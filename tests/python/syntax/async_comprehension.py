# vybe-test: python/syntax/async_comprehension
# origin: languages/python/tests/python/test_syntax.rs
# An `async for` comprehension is only legal INSIDE an `async def` — at
# module level it is a SyntaxError. Wrapped in a coroutine, with a real
# async iterator to consume.
import asyncio

async def _agen():
    yield 1
    yield 2

async def main():
    result = [x async for x in _agen()]
    return result

asyncio.run(main())
