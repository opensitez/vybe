# vybe-test: python/python_typing_typeddict_required/test_typing_typeddict_total_false_optional_keys
# origin: languages/python/tests/python/test_python_typing_typeddict_required.rs

from typing import TypedDict

class Settings(TypedDict, total=False):
    debug: bool

print(set(Settings.__optional_keys__) == {"debug"})
print(len(Settings.__required_keys__) == 0)
