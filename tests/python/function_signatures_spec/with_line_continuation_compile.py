# vybe-test: python/function_signatures_spec/with_line_continuation_compile
# origin: languages/python/tests/python/test_function_signatures_spec.rs
# vybe-test-mode: compile

with open('a') as f, \
+     open('b') as g:
    pass
