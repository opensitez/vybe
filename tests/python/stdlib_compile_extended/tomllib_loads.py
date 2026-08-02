# vybe-test: python/stdlib_compile_extended/tomllib_loads
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

try:
 import tomllib
 tomllib.loads('a=1')
except ImportError:
 pass
