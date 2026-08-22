# vybe-test: python/string_methods_extended_spec/str_translate_compile
# origin: languages/python/tests/python/test_string_methods_extended_spec.rs

tbl = str.maketrans({'a': 'x'})
s = 'aba'.translate(tbl)
