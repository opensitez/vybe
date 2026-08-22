# vybe-test: python/programs/enumerate_pattern
# origin: languages/python/tests/python/test_programs.rs

items = ["apple", "banana", "cherry"]
for i, item in enumerate(items):
    print(f"{i}: {item}")
