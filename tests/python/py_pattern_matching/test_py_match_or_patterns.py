# vybe-test: python/py_pattern_matching/test_py_match_or_patterns
# origin: languages/python/tests/python/test_py_pattern_matching.rs

def describe_status(code):
    match code:
        case 200 | 201 | 202:
            return "success"
        case 301 | 302:
            return "redirect"
        case 400:
            return "bad request"
        case 404:
            return "not found"
        case 500 | 503:
            return "server error"
        case _:
            return "unknown"

for code in [200, 302, 404, 500, 999]:
    print(describe_status(code))
