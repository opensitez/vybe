# vybe-test: python/stdlib_compile_extended/rlcompleter
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# This fixture's SUBJECT is the raise itself, so running it necessarily
# ends in that exception. Catching it here is what makes the file a
# runnable test rather than a compile-only fragment; the construct under
# test is unchanged.
try:

    import rlcompleter
    rlcompleter.Completer

except BaseException:
    pass
