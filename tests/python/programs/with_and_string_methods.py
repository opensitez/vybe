# vybe-test: python/programs/with_and_string_methods
# origin: languages/python/tests/python/test_programs.rs

text = "  Hello, World!  "
cleaned = text.strip().lower()
words = cleaned.split()
print(len(words))
print("hello" in cleaned)
