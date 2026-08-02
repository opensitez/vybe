# vybe-test: python/list_methods_extended/list_sort_reverse_key_compile
# origin: languages/python/tests/python/test_list_methods_extended.rs
# vybe-test-mode: compile

a = [3, 1, 2]
a.sort(reverse=True, key=abs)
