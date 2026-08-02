# vybe-test: python/walrus_core/walrus_in_while_condition
# origin: languages/python/tests/python/test_walrus_core.rs

it = iter([1, 2])
while (x := next(it, None)) is not None:
 print(x)
