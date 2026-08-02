# vybe-test: python/del_statement/del_comprehension_temp_not_applicable
# origin: languages/python/tests/python/test_del_statement.rs

a = [x for x in range(3)]
del a[0]
print(a)
