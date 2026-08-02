# vybe-test: python/programs/fibonacci_runtime
# origin: languages/python/tests/python/test_programs.rs

def fib(n):
    if n <= 1:
        return n
    return fib(n - 1) + fib(n - 2)

for i in range(10):
    print(fib(i))
