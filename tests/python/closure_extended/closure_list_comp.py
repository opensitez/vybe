# vybe-test: python/closure_extended/closure_list_comp
# origin: languages/python/tests/python/test_closure_extended.rs

n = 2
print([lambda: n for _ in range(1)][0]())
