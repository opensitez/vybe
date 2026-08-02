# vybe-test: python/divmod_floordiv/divmod_in_loop_accumulator
# origin: languages/python/tests/python/test_divmod_floordiv.rs

total = 0
for n in [10, 11, 12]:
 q, r = divmod(n, 3)
 total += r
print(total)
