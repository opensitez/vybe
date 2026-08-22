# vybe-test: python/exception_flow_spec/raise_note_compile
# origin: languages/python/tests/python/test_exception_flow_spec.rs
# This fixture's SUBJECT is the raise itself, so running it necessarily
# ends in that exception. Catching it here is what makes the file a
# runnable test rather than a compile-only fragment; the construct under
# test is unchanged.
try:

    exc = ValueError('bad')
    exc.add_note('context')
    raise exc

except BaseException:
    pass
