# vybe-test: python/sorted_builtins/min_empty_raises_or_handles
# origin: languages/python/tests/python/test_sorted_builtins.rs

try:
 print(min([]))
except Exception as e:
 print(type(e).__name__)
