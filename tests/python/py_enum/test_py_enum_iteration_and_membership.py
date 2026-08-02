# vybe-test: python/py_enum/test_py_enum_iteration_and_membership
# origin: languages/python/tests/python/test_py_enum.rs

from enum import Enum

class Status(Enum):
    PENDING = "pending"
    ACTIVE = "active"
    DONE = "done"

print([s.value for s in Status])
print(Status.ACTIVE in Status)
print("active" in [s.value for s in Status])
