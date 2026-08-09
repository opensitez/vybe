// Python enum advanced — Flag, IntEnum, StrEnum, auto(), aliases, iteration
use super::helpers::run_python;

#[test]
fn test_enum_flag_bitwise() {
    let script = r#"
from enum import Flag, auto

class Permission(Flag):
    READ = auto()
    WRITE = auto()
    EXEC = auto()

p = Permission.READ | Permission.WRITE
print(Permission.READ in p)
print(Permission.EXEC in p)
"#;
    assert_eq!(run_python(script), vec!["True", "False"]);
}

#[test]
fn test_enum_intenum_comparison() {
    let script = r#"
from enum import IntEnum

class Status(IntEnum):
    PENDING = 0
    ACTIVE = 1
    DONE = 2

print(Status.ACTIVE == 1)
print(Status.PENDING < Status.DONE)
print(Status.DONE + 1)
"#;
    assert_eq!(run_python(script), vec!["True", "True", "3"]);
}

#[test]
fn test_enum_auto_values() {
    let script = r#"
from enum import Enum, auto

class Color(Enum):
    RED = auto()
    GREEN = auto()
    BLUE = auto()

print(Color.RED.value)
print(Color.GREEN.value)
print(Color.BLUE.value)
"#;
    assert_eq!(run_python(script), vec!["1", "2", "3"]);
}

#[test]
fn test_enum_iteration() {
    let script = r#"
from enum import Enum

class Day(Enum):
    MON = 1
    TUE = 2
    WED = 3

for d in Day:
    print(d.name)
"#;
    assert_eq!(run_python(script), vec!["MON", "TUE", "WED"]);
}

#[test]
fn test_enum_alias() {
    let script = r#"
from enum import Enum

class Color(Enum):
    RED = 1
    ROUGE = 1  # alias

print(Color.ROUGE is Color.RED)
print(Color.ROUGE.value)
"#;
    assert_eq!(run_python(script), vec!["True", "1"]);
}

#[test]
fn test_enum_by_value() {
    let script = r#"
from enum import Enum

class Planet(Enum):
    EARTH = 3
    MARS = 4

p = Planet(3)
print(p)
print(p.name)
"#;
    assert_eq!(run_python(script), vec!["Planet.EARTH", "EARTH"]);
}

#[test]
fn test_enum_missing_handler() {
    let script = r#"
from enum import Enum

class Code(Enum):
    OK = 200
    NOT_FOUND = 404

    @classmethod
    def _missing_(cls, value):
        return None

print(Code(200))
print(Code(999))
"#;
    assert_eq!(run_python(script), vec!["Code.OK", "None"]);
}

#[test]
fn test_enum_custom_method() {
    let script = r#"
from enum import Enum

class Direction(Enum):
    NORTH = 0
    EAST = 90
    SOUTH = 180
    WEST = 270

    def opposite(self):
        return Direction((self.value + 180) % 360)

print(Direction.NORTH.opposite())
print(Direction.EAST.opposite())
"#;
    assert_eq!(
        run_python(script),
        vec!["Direction.SOUTH", "Direction.WEST"]
    );
}
