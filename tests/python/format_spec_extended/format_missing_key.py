# vybe-test: python/format_spec_extended/format_missing_key
# origin: languages/python/tests/python/test_format_spec_extended.rs
# vybe-test-mode: compile

try:
 '{x}'.format()
except KeyError:
 pass
