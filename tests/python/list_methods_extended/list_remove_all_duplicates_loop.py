# vybe-test: python/list_methods_extended/list_remove_all_duplicates_loop
# origin: languages/python/tests/python/test_list_methods_extended.rs
# vybe-test-mode: compile

a = [1, 2, 2, 3]
while 2 in a:
    a.remove(2)
