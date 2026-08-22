# vybe-test: python/stdlib_compile_extended/traceback_format_exc
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs

import traceback
try:
 1/0
except:
 traceback.format_exc()
