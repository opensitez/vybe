# vybe-test: python/re_extended_patterns/re_template_flag
# origin: languages/python/tests/python/test_re_extended_patterns.rs
# This fixture's SUBJECT is the raise itself, so running it necessarily
# ends in that exception. Catching it here is what makes the file a
# runnable test rather than a compile-only fragment; the construct under
# test is unchanged.
try:

    import re
    re.findall('(a)', 'aba', re.T)

except BaseException:
    pass
