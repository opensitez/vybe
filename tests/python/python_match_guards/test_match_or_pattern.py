# vybe-test: python/python_match_guards/test_match_or_pattern
# origin: languages/python/tests/python/test_python_match_guards.rs

for status in [200, 201, 404, 500]:
    match status:
        case 200 | 201:
            print("ok")
        case 404:
            print("not found")
        case _:
            print("error")
