# vybe-test: python/exceptions/typed_except_with_names
# origin: languages/python/tests/python/test_exceptions.rs

try:
    result = dangerous_operation()
except ValueError as ve:
    print("ValueError:", ve)
except TypeError as te:
    print("TypeError:", te)
except Exception as e:
    print("Other:", e)
