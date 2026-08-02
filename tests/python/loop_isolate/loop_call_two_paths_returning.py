# vybe-test: python/loop_isolate/loop_call_two_paths_returning
# origin: languages/python/tests/python/test_loop_isolate.rs

def k(n):
    if n <= 0:
        return n
    return n
for i in range(3):
    print(k(i))
