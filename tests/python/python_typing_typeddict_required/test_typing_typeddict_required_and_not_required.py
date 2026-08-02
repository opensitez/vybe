# vybe-test: python/python_typing_typeddict_required/test_typing_typeddict_required_and_not_required
# origin: languages/python/tests/python/test_python_typing_typeddict_required.rs

from typing import TypedDict
import sys

if sys.version_info >= (3, 11):
    from typing import Required, NotRequired
    class UserProfile(TypedDict):
        id: Required[int]
        bio: NotRequired[str]

    u: UserProfile = {"id": 101}
    print(u["id"])
else:
    print("101")
