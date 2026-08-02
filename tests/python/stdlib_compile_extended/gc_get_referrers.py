# vybe-test: python/stdlib_compile_extended/gc_get_referrers
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

import gc
x = []
gc.get_referrers(x)
