# vybe-test: python/loop_isolate/loop_call_with_if_early_return
# origin: languages/python/tests/python/test_loop_isolate.rs

def h(n):
    if n <= 0:
        return n
    return n + 100
for i in range(3):
    print(h(i))
