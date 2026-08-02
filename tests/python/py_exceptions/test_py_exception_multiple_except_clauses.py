# vybe-test: python/py_exceptions/test_py_exception_multiple_except_clauses
# origin: languages/python/tests/python/test_py_exceptions.rs

def risky(x):
    if x == 0:
        raise ZeroDivisionError("zero!")
    if x < 0:
        raise ValueError("negative!")
    return x

for val in [1, 0, -1]:
    try:
        print(risky(val))
    except ZeroDivisionError:
        print("ZDE")
    except ValueError:
        print("VE")
