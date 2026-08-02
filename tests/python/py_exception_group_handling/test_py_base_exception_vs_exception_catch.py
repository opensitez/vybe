# vybe-test: python/py_exception_group_handling/test_py_base_exception_vs_exception_catch
# origin: languages/python/tests/python/test_py_exception_group_handling.rs

# KeyboardInterrupt inherits from BaseException, NOT Exception
try:
    raise KeyboardInterrupt("stop")
except Exception:
    print("Caught by Exception")
except BaseException as e:
    print(f"Caught by BaseException: {type(e).__name__}")
