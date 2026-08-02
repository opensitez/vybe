# vybe-test: python/python_warnings_categories/test_warnings_filterwarnings_action_always_accumulates
# origin: languages/python/tests/python/test_python_warnings_categories.rs

import warnings
with warnings.catch_warnings(record=True) as w:
    warnings.filterwarnings("always", message=".*repeat.*", category=UserWarning)
    for _ in range(3):
        warnings.warn("repeat me", UserWarning)
print(len(w))
