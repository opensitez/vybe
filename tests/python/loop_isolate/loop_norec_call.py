# vybe-test: python/loop_isolate/loop_norec_call
# origin: languages/python/tests/python/test_loop_isolate.rs

def g(n):
    return n + 1
for i in range(3):
    print(g(i))
