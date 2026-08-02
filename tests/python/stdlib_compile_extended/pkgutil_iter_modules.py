# vybe-test: python/stdlib_compile_extended/pkgutil_iter_modules
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

import pkgutil
list(pkgutil.iter_modules())
