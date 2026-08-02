# vybe-test: python/py_heapq_bisect_array/test_py_bisect_grade_lookup
# origin: languages/python/tests/python/test_py_heapq_bisect_array.rs

import bisect

grades = [("F", 60), ("D", 65), ("C", 70), ("B", 80), ("A", 90)]
breakpoints = [bp for _, bp in grades]
letters = [letter for letter, _ in grades]

def grade(score):
    idx = bisect.bisect_left(breakpoints, score)
    if idx >= len(letters):
        return "A+"
    return letters[idx]

print(grade(55))
print(grade(65))
print(grade(80))
print(grade(95))
