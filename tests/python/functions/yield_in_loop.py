# vybe-test: python/functions/yield_in_loop
# origin: languages/python/tests/python/test_functions.rs

def count_up(n):
    i = 0
    while i < n:
        yield i
        i += 1
