# vybe-test: python/python_typing_typeddict_required/test_typing_typeddict_required_keys_and_optional_keys
# origin: languages/python/tests/python/test_python_typing_typeddict_required.rs

from typing import TypedDict

class Config(TypedDict):
    host: str
    port: int

print(set(Config.__required_keys__) == {"host", "port"})
print(len(Config.__optional_keys__) == 0)
