# vybe-test: python/syntax/walrus_in_if
# origin: languages/python/tests/python/test_syntax.rs
# vybe-test-mode: compile

if (n := 10) > 5:
    print(n)
