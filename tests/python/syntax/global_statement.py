# vybe-test: python/syntax/global_statement
# origin: languages/python/tests/python/test_syntax.rs

x = 10
def change():
    global x
    x = 20
