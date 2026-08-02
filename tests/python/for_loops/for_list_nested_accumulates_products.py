# vybe-test: python/for_loops/for_list_nested_accumulates_products
# origin: languages/python/tests/python/test_for_loops.rs

total = 0
for row in [[1, 2], [3, 4]]:
    for x in row:
        total += x
print(total)
