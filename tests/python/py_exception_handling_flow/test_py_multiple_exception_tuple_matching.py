# vybe-test: python/py_exception_handling_flow/test_py_multiple_exception_tuple_matching
# origin: languages/python/tests/python/test_py_exception_handling_flow.rs

def parse(val):
    try:
        if val == "zero":
            1 / 0
        elif val == "int":
            int("abc")
        elif val == "key":
            {}[val]
    except (ZeroDivisionError, ValueError, KeyError) as e:
        print(f"Caught {type(e).__name__}")

parse("zero")
parse("int")
parse("key")
