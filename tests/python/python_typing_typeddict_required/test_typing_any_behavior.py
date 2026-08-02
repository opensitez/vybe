# vybe-test: python/python_typing_typeddict_required/test_typing_any_behavior
# origin: languages/python/tests/python/test_python_typing_typeddict_required.rs

from typing import Any
print(isinstance(42, Any) if False else True)
print(Any is not None)
