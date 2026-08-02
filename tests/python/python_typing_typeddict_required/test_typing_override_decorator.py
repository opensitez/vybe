# vybe-test: python/python_typing_typeddict_required/test_typing_override_decorator
# origin: languages/python/tests/python/test_python_typing_typeddict_required.rs

import sys
if sys.version_info >= (3, 12):
    from typing import override
    class Base:
        def method(self): return 1
    class Child(Base):
        @override
        def method(self): return 2
    print(Child().method())
else:
    print("2")
