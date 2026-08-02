# vybe-test: python/py_pattern_matching/test_py_match_guard_clauses
# origin: languages/python/tests/python/test_py_pattern_matching.rs

def classify_number(n):
    match n:
        case x if x < 0:
            return "negative"
        case 0:
            return "zero"
        case x if x % 2 == 0:
            return "positive even"
        case _:
            return "positive odd"

for n in [-5, 0, 4, 7]:
    print(classify_number(n))
