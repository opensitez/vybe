# vybe-test: python/while_loops/while_reads_input_sentinel_pattern
# origin: languages/python/tests/python/test_while_loops.rs

data = [1, 2, -1, 3]
i = 0
total = 0
while True:
 v = data[i]
 i += 1
 if v < 0:
  break
 total += v
print(total)
