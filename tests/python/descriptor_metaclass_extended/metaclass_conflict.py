# vybe-test: python/descriptor_metaclass_extended/metaclass_conflict
# origin: languages/python/tests/python/test_descriptor_metaclass_extended.rs
# This fixture's SUBJECT is that Python REJECTS the construct, so the file
# cannot itself be valid Python. `compile()` lets it assert the rejection
# while remaining a runnable test.
_SRC = """
class M1(type): pass
class M2(type): pass
try:
 class C(metaclass=M1, metaclass=M2): pass
except TypeError: pass
"""
try:
    compile(_SRC, '<fixture>', 'exec')
except SyntaxError:
    pass
