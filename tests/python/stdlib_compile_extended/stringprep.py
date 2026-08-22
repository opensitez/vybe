# vybe-test: python/stdlib_compile_extended/stringprep
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# This fixture's SUBJECT is the raise itself, so running it necessarily
# ends in that exception. Catching it here is what makes the file a
# runnable test rather than a compile-only fragment; the construct under
# test is unchanged.
try:

    import stringprep
    stringprep.in_table_a1('a')

except BaseException:
    pass
