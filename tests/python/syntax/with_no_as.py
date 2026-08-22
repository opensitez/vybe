# vybe-test: python/syntax/with_no_as
# origin: languages/python/tests/python/test_syntax.rs
# `lock` and `do_stuff` were never defined.
import threading
lock = threading.Lock()

def do_stuff():
    pass

with lock:
    do_stuff()
