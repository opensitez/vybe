# vybe-test: python/syntax/walrus_in_while
# origin: languages/python/tests/python/test_syntax.rs
# vybe-test-mode: compile

while chunk := input():
    process(chunk)
