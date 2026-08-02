# vybe-test: python/python_scope_closures/test_unbound_local_error
# origin: languages/python/tests/python/test_python_scope_closures.rs

x = 10

def bad():
    try:
        print(x)
        x = 20
    except UnboundLocalError:
        print("UnboundLocalError")

bad()
