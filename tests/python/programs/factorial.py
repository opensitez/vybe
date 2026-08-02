# vybe-test: python/programs/factorial
# origin: languages/python/tests/python/test_programs.rs
# vybe-test-mode: compile

def factorial(n):
    result = 1
    for i in range(1, n + 1):
        result *= i
    return result

print(factorial(10))
