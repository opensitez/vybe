# vybe-test: python/walrus_core/walrus_sum_accumulator_pattern
# origin: languages/python/tests/python/test_walrus_core.rs

total = 0
nums = [1, 2, 3]
while nums and (total := total + nums.pop()):
 pass
print(total)
