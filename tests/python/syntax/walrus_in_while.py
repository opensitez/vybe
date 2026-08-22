# vybe-test: python/syntax/walrus_in_while
# origin: languages/python/tests/python/test_syntax.rs
# This fixture's SUBJECT is the raise itself, so running it necessarily
# ends in that exception. Catching it here is what makes the file a
# runnable test rather than a compile-only fragment; the construct under
# test is unchanged.
# `input()` blocks on stdin forever in a test run. The construct under
# test is the walrus in a `while` condition — a finite iterator exercises
# it identically and terminates.
_it = iter(['a', 'b', ''])
while chunk := next(_it):
    pass
