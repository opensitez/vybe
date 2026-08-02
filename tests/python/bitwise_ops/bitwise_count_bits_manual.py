# vybe-test: python/bitwise_ops/bitwise_count_bits_manual
# origin: languages/python/tests/python/test_bitwise_ops.rs

n = 7
count = 0
while n:
 count += n & 1
 n >>= 1
print(count)
