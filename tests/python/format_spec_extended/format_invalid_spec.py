# vybe-test: python/format_spec_extended/format_invalid_spec
# origin: languages/python/tests/python/test_format_spec_extended.rs
# vybe-test-mode: compile

try:
 '{:z}'.format(1)
except ValueError:
 pass
