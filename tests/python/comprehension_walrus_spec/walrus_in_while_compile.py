# vybe-test: python/comprehension_walrus_spec/walrus_in_while_compile
# origin: languages/python/tests/python/test_comprehension_walrus_spec.rs
# vybe-test-mode: compile

while (line := reader()):
    print(line)
