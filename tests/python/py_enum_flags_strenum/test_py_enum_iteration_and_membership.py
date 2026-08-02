# vybe-test: python/py_enum_flags_strenum/test_py_enum_iteration_and_membership
# origin: languages/python/tests/python/test_py_enum_flags_strenum.rs

from enum import Enum

class Status(Enum):
    PENDING = "pending"
    SUCCESS = "success"

print([s.value for s in Status])
print(Status.SUCCESS in Status)
