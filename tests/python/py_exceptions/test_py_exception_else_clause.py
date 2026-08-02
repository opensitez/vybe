# vybe-test: python/py_exceptions/test_py_exception_else_clause
# origin: languages/python/tests/python/test_py_exceptions.rs

results = []

for x in [1, 0, 2]:
    try:
        val = 10 / x
    except ZeroDivisionError:
        results.append("error")
    else:
        results.append(f"ok:{val}")

print(results)
