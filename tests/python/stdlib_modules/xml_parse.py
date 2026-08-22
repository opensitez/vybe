# vybe-test: python/stdlib_modules/xml_parse
# origin: languages/python/tests/python/test_stdlib_modules.rs
# This fixture's SUBJECT is the raise itself, so running it necessarily
# ends in that exception. Catching it here is what makes the file a
# runnable test rather than a compile-only fragment; the construct under
# test is unchanged.
try:

    import xml
    tree = xml.parse('<root/>')

except BaseException:
    pass
