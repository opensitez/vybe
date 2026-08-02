# vybe-test: python/stdlib_compile_extended/msvcrt_locking
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

try:
 import msvcrt
except ImportError:
 pass
