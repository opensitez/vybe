# vybe-test: python/programs/list_processing
# origin: languages/python/tests/python/test_programs.rs
# vybe-test-mode: compile

numbers = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
evens = [x for x in numbers if x % 2 == 0]
squares = [x ** 2 for x in evens]
total = 0
for s in squares:
    total += s
print(total)
