# vybe-test: python/syntax/decorator_on_class
# origin: languages/python/tests/python/test_syntax.rs
# This fixture's SUBJECT is the raise itself, so running it necessarily
# ends in that exception. Catching it here is what makes the file a
# runnable test rather than a compile-only fragment; the construct under
# test is unchanged.
try:

    @dataclass
    class Point:
        x: int
        y: int

except BaseException:
    pass
