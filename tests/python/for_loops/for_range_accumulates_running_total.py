# vybe-test: python/for_loops/for_range_accumulates_running_total
# origin: languages/python/tests/python/test_for_loops.rs

total = 0
for i in range(1, 5):
    total += i
print(total)
