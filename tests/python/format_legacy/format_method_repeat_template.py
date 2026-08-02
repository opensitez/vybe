# vybe-test: python/format_legacy/format_method_repeat_template
# origin: languages/python/tests/python/test_format_legacy.rs

'-'.join(['{:02d}'.format(x) for x in range(3)])
