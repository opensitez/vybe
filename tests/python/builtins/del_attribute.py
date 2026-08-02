# vybe-test: python/builtins/del_attribute
# origin: languages/python/tests/python/test_builtins.rs
# vybe-test-mode: compile

class C:
    pass
c = C()
c.x = 10
del c.x
