# vybe-test: python/comprehension_walrus_spec/walrus_in_while_compile
# origin: languages/python/tests/python/test_comprehension_walrus_spec.rs
def reader(*_a, **_k):
    return None

while (line := reader()):
    print(line)
