# vybe-test: python/list_methods_spec/list_sort_key_compile
# origin: languages/python/tests/python/test_list_methods_spec.rs
# vybe-test-mode: compile

x = ['bbb', 'a', 'cc']
x.sort(key=len)
