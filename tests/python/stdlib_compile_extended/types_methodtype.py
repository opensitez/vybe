# vybe-test: python/stdlib_compile_extended/types_methodtype
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

import types
class C:
 def m(self): pass
types.MethodType
