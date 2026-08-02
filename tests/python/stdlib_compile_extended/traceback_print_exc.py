# vybe-test: python/stdlib_compile_extended/traceback_print_exc
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

import traceback
try:
 1/0
except:
 traceback.print_exc()
