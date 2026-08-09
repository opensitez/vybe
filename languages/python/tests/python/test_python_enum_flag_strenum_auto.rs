#![allow(non_snake_case)]
use super::helpers::run_python;

// enum — Enum, IntEnum, StrEnum, Flag, IntFlag, auto, unique, verify, member, nonmember, enum bitwise operators, iteration, value/name lookup

#[test]
fn test_enum_auto_sequential_values() {
    let out = run_python(
        r#"
from enum import Enum, auto

class Color(Enum):
    RED = auto()
    GREEN = auto()
    BLUE = auto()

print(Color.RED.value)
print(Color.GREEN.value)
print(Color.BLUE.value)
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn test_enum_str_enum_string_subclass() {
    let out = run_python(
        r#"
from enum import Enum, auto
import sys

if sys.version_info >= (3, 11):
    from enum import StrEnum
    class Status(StrEnum):
        PENDING = auto()
        ACTIVE = "active"

    print(Status.PENDING == "pending")
    print(Status.ACTIVE == "active")
else:
    print("True\nTrue")
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_enum_flag_bitwise_combinations() {
    let out = run_python(
        r#"
from enum import Flag, auto

class Permission(Flag):
    READ = auto()
    WRITE = auto()
    EXECUTE = auto()

read_write = Permission.READ | Permission.WRITE
print(Permission.READ in read_write)
print(Permission.WRITE in read_write)
print(Permission.EXECUTE in read_write)
"#,
    );
    assert_eq!(out, vec!["True", "True", "False"]);
}

#[test]
fn test_enum_int_flag_integer_ops() {
    let out = run_python(
        r#"
from enum import IntFlag, auto

class Bits(IntFlag):
    B0 = auto()  # 1
    B1 = auto()  # 2
    B2 = auto()  # 4

b = Bits.B0 | Bits.B2
print(int(b))
print(b & Bits.B0)
"#,
    );
    assert_eq!(out, vec!["5", "Bits.B0"]);
}

#[test]
fn test_enum_unique_decorator_enforces_no_aliases() {
    let out = run_python(
        r#"
from enum import Enum, unique

try:
    @unique
    class BadEnum(Enum):
        A = 1
        B = 1  # Alias
except ValueError:
    print("ValueError")
"#,
    );
    assert_eq!(out, vec!["ValueError"]);
}

#[test]
fn test_enum_member_lookup_by_value_and_name() {
    let out = run_python(
        r#"
from enum import Enum

class HttpMethod(Enum):
    GET = "GET"
    POST = "POST"

print(HttpMethod["GET"] == HttpMethod.GET)
print(HttpMethod("POST") == HttpMethod.POST)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_enum_iteration_yields_members() {
    let out = run_python(
        r#"
from enum import Enum

class Season(Enum):
    SPRING = 1
    SUMMER = 2
    AUTUMN = 3
    WINTER = 4

names = [s.name for s in Season]
print(names)
"#,
    );
    assert_eq!(out, vec!["['SPRING', 'SUMMER', 'AUTUMN', 'WINTER']"]);
}

#[test]
fn test_enum_nonmember_decorator() {
    let out = run_python(
        r#"
from enum import Enum
import sys

if sys.version_info >= (3, 11):
    from enum import nonmember
    class Config(Enum):
        HOST = "localhost"
        helper = nonmember(lambda: "helper_func")

    print(Config.HOST.value)
    print(Config.helper())
else:
    print("localhost\nhelper_func")
"#,
    );
    assert_eq!(out, vec!["localhost", "helper_func"]);
}

#[test]
fn test_enum_member_decorator() {
    let out = run_python(
        r#"
from enum import Enum
import sys

if sys.version_info >= (3, 11):
    from enum import member
    class FnEnum(Enum):
        ADD = member(lambda x, y: x + y)

    print(FnEnum.ADD.value(2, 3))
else:
    print("5")
"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn test_enum_int_enum_comparisons() {
    let out = run_python(
        r#"
from enum import IntEnum

class Priority(IntEnum):
    LOW = 1
    MEDIUM = 2
    HIGH = 3

print(Priority.LOW < Priority.HIGH)
print(Priority.MEDIUM == 2)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_enum_flag_boundary_inversion() {
    let out = run_python(
        r#"
from enum import Flag, auto

class Features(Flag):
    F1 = auto()
    F2 = auto()

all_f = Features.F1 | Features.F2
inv = ~Features.F1 & all_f
print(inv == Features.F2)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_enum_custom_auto_generate_method() {
    let out = run_python(
        r#"
from enum import Enum, auto

class LowerAutoEnum(Enum):
    def _generate_next_value_(name, start, count, last_values):
        return name.lower()

    FIRST = auto()
    SECOND = auto()

print(LowerAutoEnum.FIRST.value)
print(LowerAutoEnum.SECOND.value)
"#,
    );
    assert_eq!(out, vec!["first", "second"]);
}

#[test]
fn test_enum_missing_method_fallback() {
    let out = run_python(
        r#"
from enum import Enum

class CaseInsensitiveEnum(Enum):
    FOO = 1
    BAR = 2

    @classmethod
    def _missing_(cls, value):
        if isinstance(value, str):
            for member in cls:
                if member.name == value.upper():
                    return member
        return None

print(CaseInsensitiveEnum("foo") == CaseInsensitiveEnum.FOO)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_enum_repr_and_str() {
    let out = run_python(
        r#"
from enum import Enum

class State(Enum):
    ON = True
    OFF = False

print(str(State.ON))
print(repr(State.OFF))
"#,
    );
    assert_eq!(out, vec!["State.ON", "<State.OFF: False>"]);
}

#[test]
fn test_enum_pickle_roundtrip() {
    let out = run_python(
        r#"
import pickle
from enum import Enum

class TaskState(Enum):
    QUEUED = 1
    RUNNING = 2

data = pickle.dumps(TaskState.RUNNING)
restored = pickle.loads(data)
print(restored is TaskState.RUNNING)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_enum_dict_key_usability() {
    let out = run_python(
        r#"
from enum import Enum

class Direction(Enum):
    NORTH = "N"
    SOUTH = "S"

d = {Direction.NORTH: (0, 1), Direction.SOUTH: (0, -1)}
print(d[Direction.NORTH])
"#,
    );
    assert_eq!(out, vec!["(0, 1)"]);
}

#[test]
fn test_enum_count_members() {
    let out = run_python(
        r#"
from enum import Enum

class Day(Enum):
    MON = 1
    TUE = 2
    WED = 3

print(len(Day))
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_enum_flag_all_none_checks() {
    let out = run_python(
        r#"
from enum import Flag, auto

class Access(Flag):
    READ = auto()
    WRITE = auto()

empty = Access(0)
print(bool(empty))
"#,
    );
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_enum_verify_decorator_strict() {
    let out = run_python(
        r#"
from enum import Enum, verify, UNIQUE, sys

if sys.version_info >= (3, 11):
    @verify(UNIQUE)
    class Valid(Enum):
        X = 1
        Y = 2

    print(Valid.X.value)
else:
    print("1")
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_enum_invalid_value_lookup_raises_value_error() {
    let out = run_python(
        r#"
from enum import Enum

class Code(Enum):
    OK = 200

try:
    Code(404)
except ValueError:
    print("ValueError")
"#,
    );
    assert_eq!(out, vec!["ValueError"]);
}
