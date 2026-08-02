# vybe-test: python/programs/nested_comprehension_real
# origin: languages/python/tests/python/test_programs.rs
# vybe-test-mode: compile

matrix = [[1,2,3],[4,5,6],[7,8,9]]
flat = [x for row in matrix for x in row]
evens = [x for x in flat if x % 2 == 0]
print(sorted(evens))
print(sum(evens))
