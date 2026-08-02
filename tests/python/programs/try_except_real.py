# vybe-test: python/programs/try_except_real
# origin: languages/python/tests/python/test_programs.rs
# vybe-test-mode: compile

def safe_divide(a, b):
    try:
        return a / b
    except:
        print("division error")
        return 0

print(safe_divide(10, 2))
print(safe_divide(10, 0))
