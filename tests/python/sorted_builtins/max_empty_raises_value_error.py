# vybe-test: python/sorted_builtins/max_empty_raises_value_error
# origin: languages/python/tests/python/test_sorted_builtins.rs

try:
 print(max([]))
except Exception as e:
 print(type(e).__name__)
