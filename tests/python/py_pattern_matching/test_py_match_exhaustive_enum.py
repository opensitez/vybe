# vybe-test: python/py_pattern_matching/test_py_match_exhaustive_enum
# origin: languages/python/tests/python/test_py_pattern_matching.rs

from enum import Enum, auto

class Status(Enum):
    PENDING = auto()
    ACTIVE = auto()
    DONE = auto()
    FAILED = auto()

def describe(s: Status) -> str:
    match s:
        case Status.PENDING:
            return "waiting"
        case Status.ACTIVE:
            return "running"
        case Status.DONE:
            return "completed"
        case Status.FAILED:
            return "failed"

for s in Status:
    print(describe(s))
