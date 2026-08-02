# vybe-test: python/while_loops/while_counts_bits_in_integer
# origin: languages/python/tests/python/test_while_loops.rs

n = 13
count = 0
while n:
 count += n & 1
 n >>= 1
print(count)
