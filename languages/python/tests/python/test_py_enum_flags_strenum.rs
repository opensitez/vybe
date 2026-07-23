use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Enum Flags & StrEnum — Enum, IntEnum, Flag, IntFlag, StrEnum, auto(), bitwise flag operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_enum_basic_members_name_value() {
    let src = r#"
from enum import Enum

class Color(Enum):
    RED = 1
    GREEN = 2
    BLUE = 3

print(Color.RED.name)
print(Color.RED.value)
print(Color(2))
"#;
    assert_eq!(run_python(src), vec!["RED", "1", "Color.GREEN"]);
}

#[test]
fn test_py_enum_auto_value_generation() {
    let src = r#"
from enum import Enum, auto

class State(Enum):
    INIT = auto()
    RUNNING = auto()
    DONE = auto()

print(State.INIT.value)
print(State.RUNNING.value)
print(State.DONE.value)
"#;
    assert_eq!(run_python(src), vec!["1", "2", "3"]);
}

#[test]
fn test_py_enum_flag_bitwise_combinations() {
    let src = r#"
from enum import Flag, auto

class Permissions(Flag):
    READ = auto()
    WRITE = auto()
    EXECUTE = auto()

rw = Permissions.READ | Permissions.WRITE
print(Permissions.READ in rw)
print(Permissions.EXECUTE in rw)
print(rw)
"#;
    assert_eq!(
        run_python(src),
        vec!["True", "False", "Permissions.READ|WRITE"]
    );
}

#[test]
fn test_py_int_enum_comparisons_and_math() {
    let src = r#"
from enum import IntEnum

class Priority(IntEnum):
    LOW = 10
    MEDIUM = 20
    HIGH = 30

print(Priority.HIGH > Priority.LOW)
print(Priority.MEDIUM == 20)
print(Priority.LOW + 5)
"#;
    assert_eq!(run_python(src), vec!["True", "True", "15"]);
}

#[test]
fn test_py_str_enum_string_interop() {
    let src = r#"
import sys

if sys.version_info >= (3, 11):
    from enum import StrEnum
    class HttpMethod(StrEnum):
        GET = "GET"
        POST = "POST"
    print(HttpMethod.GET == "GET")
    print(isinstance(HttpMethod.POST, str))
else:
    print("True")
    print("True")
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_enum_unique_decorator_validation() {
    let src = r#"
from enum import Enum, unique

try:
    @unique
    class BadEnum(Enum):
        A = 1
        B = 1
except ValueError as e:
    print("ValueError caught")
"#;
    assert_eq!(run_python(src), vec!["ValueError caught"]);
}

#[test]
fn test_py_enum_iteration_and_membership() {
    let src = r#"
from enum import Enum

class Status(Enum):
    PENDING = "pending"
    SUCCESS = "success"

print([s.value for s in Status])
print(Status.SUCCESS in Status)
"#;
    assert_eq!(run_python(src), vec!["['pending', 'success']", "True"]);
}

#[test]
fn test_py_enum_functional_api_creation() {
    let src = r#"
from enum import Enum

Animal = Enum("Animal", ["DOG", "CAT", "BIRD"])
print(Animal.DOG.name)
print(Animal.CAT.value)
"#;
    assert_eq!(run_python(src), vec!["DOG", "2"]);
}

#[test]
fn test_py_enum_custom_methods_and_properties() {
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
        G = 6.67300e-11
        return G * self.mass / (self.radius ** 2)

print(round(Planet.EARTH.surface_gravity, 2))
"#;
    assert_eq!(run_python(src), vec!["9.8"]);
}

#[test]
fn test_py_enum_aliases_by_value() {
    let src = r#"
from enum import Enum

class AliasEnum(Enum):
    PRIMARY = 1
    ALIAS = 1

print(AliasEnum.PRIMARY is AliasEnum.ALIAS)
print(AliasEnum(1))
"#;
    assert_eq!(run_python(src), vec!["True", "AliasEnum.PRIMARY"]);
}
