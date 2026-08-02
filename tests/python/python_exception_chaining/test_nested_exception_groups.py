# vybe-test: python/python_exception_chaining/test_nested_exception_groups
# origin: languages/python/tests/python/test_python_exception_chaining.rs

errors = []
for i in range(3):
    try:
        if i % 2 == 0:
            raise ValueError(f"even {i}")
    except ValueError as e:
        errors.append(str(e))

print(errors)
