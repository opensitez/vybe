# vybe-test: python/loop_isolate/loop_rec_call
# origin: languages/python/tests/python/test_loop_isolate.rs

def f(n):
    if n <= 0:
        return n
    return f(n - 1)
for i in range(3):
    print(f(i))
