# vybe-test: python/syntax/async_dict_comprehension
# origin: languages/python/tests/python/test_syntax.rs
# vybe-test-mode: compile

result = {k: v async for k, v in aitems}
