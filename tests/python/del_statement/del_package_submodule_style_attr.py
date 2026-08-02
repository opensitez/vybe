# vybe-test: python/del_statement/del_package_submodule_style_attr
# origin: languages/python/tests/python/test_del_statement.rs

class M:
 value = 1
del M.value
try:
 print(M.value)
except AttributeError:
 print('ok')
