# vybe-test: python/python_yield_from/test_yield_from_chained_generators
# origin: languages/python/tests/python/test_python_yield_from.rs

def counter(start, stop):
    while start < stop:
        yield start
        start += 1

def merged(*ranges):
    for r in ranges:
        yield from r

result = list(merged(counter(0, 3), counter(10, 13)))
print(result)
