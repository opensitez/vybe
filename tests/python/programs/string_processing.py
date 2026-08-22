# vybe-test: python/programs/string_processing
# origin: languages/python/tests/python/test_programs.rs

words = "hello world foo bar"
parts = words.split()
upper_parts = [w.upper() for w in parts]
result = " ".join(upper_parts)
print(result)
