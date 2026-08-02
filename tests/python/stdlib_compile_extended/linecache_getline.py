# vybe-test: python/stdlib_compile_extended/linecache_getline
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

import linecache
linecache.getline(__file__, 1)
