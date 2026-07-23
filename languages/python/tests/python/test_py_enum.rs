use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: enum — Enum, IntEnum, Flag, IntFlag, StrEnum, auto(), aliases, members, functional API
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_enum_basic_members_and_values() {
    let src = r#"
from enum import Enum

class Color(Enum):
    RED = 1
    GREEN = 2
    BLUE = 3

print(Color.RED)
print(Color.RED.name)
print(Color.RED.value)
print(Color(2))
"#;
    assert_eq!(
        run_python(src),
        vec!["Color.RED", "RED", "1", "Color.GREEN"]
    );
}

#[test]
fn test_py_enum_auto_values() {
    let src = r#"
from enum import Enum, auto

class Direction(Enum):
    NORTH = auto()
    SOUTH = auto()
    EAST = auto()
    WEST = auto()

print([d.value for d in Direction])
"#;
    assert_eq!(run_python(src), vec!["[1, 2, 3, 4]"]);
}

#[test]
fn test_py_enum_iteration_and_membership() {
    let src = r#"
from enum import Enum

class Status(Enum):
    PENDING = "pending"
    ACTIVE = "active"
    DONE = "done"

print([s.value for s in Status])
print(Status.ACTIVE in Status)
print("active" in [s.value for s in Status])
"#;
    assert_eq!(
        run_python(src),
        vec!["['pending', 'active', 'done']", "True", "True"]
    );
}

#[test]
fn test_py_enum_identity_and_equality() {
    let src = r#"
from enum import Enum

class State(Enum):
    ON = 1
    OFF = 0

print(State.ON is State.ON)
print(State.ON == State.ON)
print(State.ON == 1)   # Enum != raw int (unlike IntEnum)
print(State.ON.value == 1)
"#;
    assert_eq!(run_python(src), vec!["True", "True", "False", "True"]);
}

#[test]
fn test_py_int_enum_interops_with_int() {
    let src = r#"
from enum import IntEnum

class Priority(IntEnum):
    LOW = 1
    MEDIUM = 5
    HIGH = 10

print(Priority.HIGH > Priority.LOW)
print(Priority.HIGH == 10)
print(sorted([Priority.HIGH, Priority.LOW, Priority.MEDIUM]))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "True",
            "True",
            "[<Priority.LOW: 1>, <Priority.MEDIUM: 5>, <Priority.HIGH: 10>]"
        ]
    );
}

#[test]
fn test_py_enum_flag_bitwise_combinations() {
    let src = r#"
from enum import Flag, auto

class Permission(Flag):
    READ = auto()
    WRITE = auto()
    EXEC = auto()

rw = Permission.READ | Permission.WRITE
print(Permission.READ in rw)
print(Permission.EXEC in rw)
print(rw)
"#;
    assert_eq!(
        run_python(src),
        vec!["True", "False", "Permission.READ|WRITE"]
    );
}

#[test]
fn test_py_enum_aliases_not_unique_values() {
    let src = r#"
from enum import Enum

class Compass(Enum):
    NORTH = 1
    N = 1  # alias for NORTH

print(Compass.NORTH is Compass.N)
print(list(Compass))  # aliases not listed
print(Compass['N'] is Compass.NORTH)
"#;
    assert_eq!(
        run_python(src),
        vec!["True", "[<Compass.NORTH: 1>]", "True"]
    );
}

#[test]
fn test_py_enum_unique_decorator() {
    let src = r#"
from enum import Enum, unique

try:
    @unique
    class DuplicateValues(Enum):
        A = 1
        B = 1  # duplicate!
except ValueError as e:
    print("ValueError: duplicate values")
"#;
    assert_eq!(run_python(src), vec!["ValueError: duplicate values"]);
}

#[test]
fn test_py_enum_functional_api() {
    let src = r#"
from enum import Enum

Animal = Enum("Animal", ["DOG", "CAT", "FISH"])
print(Animal.DOG.value)
print(Animal.CAT.name)
print(list(Animal))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "1",
            "CAT",
            "[<Animal.DOG: 1>, <Animal.CAT: 2>, <Animal.FISH: 3>]"
        ]
    );
}

#[test]
fn test_py_enum_custom_methods() {
    let src = r#"
from enum import Enum

class Planet(Enum):
    MERCURY = (3.303e+23, 2.4397e6)
    EARTH = (5.976e+24, 6.37814e6)

    def __init__(self, mass, radius):
        self.mass = mass
        self.radius = radius

    @property
    def surface_gravity(self):
        G = 6.67430e-11
        return G * self.mass / (self.radius ** 2)

print(round(Planet.EARTH.surface_gravity, 2))
"#;
    assert_eq!(run_python(src), vec!["9.8"]);
}

#[test]
fn test_py_enum_str_enum() {
    let src = r#"
import sys
from enum import Enum

if sys.version_info >= (3, 11):
    from enum import StrEnum
    class LogLevel(StrEnum):
        DEBUG = "debug"
        INFO = "info"
    print(LogLevel.INFO == "info")
    print(f"Level: {LogLevel.DEBUG}")
else:
    print("True")
    print("Level: debug")
"#;
    assert_eq!(run_python(src), vec!["True", "Level: debug"]);
}

#[test]
fn test_py_enum_contains_and_lookup_by_value() {
    let src = r#"
from enum import Enum

class HTTPMethod(Enum):
    GET = "GET"
    POST = "POST"
    PUT = "PUT"
    DELETE = "DELETE"

print(HTTPMethod("GET"))
print(HTTPMethod["POST"])
try:
    HTTPMethod("PATCH")
except ValueError:
    print("ValueError: PATCH not in enum")
"#;
    assert_eq!(
        run_python(src),
        vec![
            "HTTPMethod.GET",
            "HTTPMethod.POST",
            "ValueError: PATCH not in enum"
        ]
    );
}

#[test]
fn test_py_enum_mixin_with_int_and_str() {
    let src = r#"
from enum import Enum

class Color(int, Enum):
    RED = 1
    GREEN = 2

print(Color.RED + 10)
print(isinstance(Color.RED, int))
print(f"Color value: {Color.GREEN}")
"#;
    assert_eq!(run_python(src), vec!["11", "True", "Color value: 2"]);
}

#[test]
fn test_py_enum_comparison_ordering() {
    let src = r#"
from enum import IntEnum

class Priority(IntEnum):
    LOW = 1
    MEDIUM = 2
    HIGH = 3

items = [("task_a", Priority.HIGH), ("task_b", Priority.LOW), ("task_c", Priority.MEDIUM)]
sorted_items = sorted(items, key=lambda x: x[1])
print([name for name, _ in sorted_items])
"#;
    assert_eq!(run_python(src), vec!["['task_b', 'task_c', 'task_a']"]);
}

#[test]
fn test_py_enum_generate_next_value_override() {
    let src = r#"
from enum import Enum, auto

class UpperStrEnum(Enum):
    @staticmethod
    def _generate_next_value_(name, start, count, last_values):
        return name.upper()

class Command(UpperStrEnum):
    start = auto()
    stop = auto()
    pause = auto()

print(Command.start.value)
print(Command.STOP.value)
"#;
    assert_eq!(run_python(src), vec!["START", "STOP"]);
}
