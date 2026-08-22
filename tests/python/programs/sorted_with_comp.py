# vybe-test: python/programs/sorted_with_comp
# origin: languages/python/tests/python/test_programs.rs

data = [5, 2, 8, 1, 9, 3]
ascending = sorted(data)
print(ascending)
total = sum(data)
avg = total / len(data)
print(f"sum={total}, avg={avg}")
