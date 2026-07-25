use super::helpers::run_python;

// typing — TypedDict, Required, NotRequired, ReadOnly, total, get_type_hints, get_args, get_origin, Annotated, Literal

#[test]
fn test_typing_typeddict_total_true_by_default() {
    let out = run_python(r#"
from typing import TypedDict

class Movie(TypedDict):
    name: str
    year: int

m: Movie = {"name": "Inception", "year": 2010}
print(m["name"], m["year"])
"#);
    assert_eq!(out, vec!["Inception 2010"]);
}

#[test]
fn test_typing_typeddict_total_false_partial_dict() {
    let out = run_python(r#"
from typing import TypedDict

class Options(TypedDict, total=False):
    verbose: bool
    log_file: str

opts: Options = {"verbose": True}
print(opts["verbose"])
print("log_file" in opts)
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_typing_typeddict_required_and_not_required() {
    let out = run_python(r#"
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
"#);
    assert_eq!(out, vec!["101"]);
}

#[test]
fn test_typing_typeddict_inheritance_combines_fields() {
    let out = run_python(r#"
from typing import TypedDict

class Person(TypedDict):
    name: str

class Employee(Person):
    employee_id: int

e: Employee = {"name": "Alice", "employee_id": 99}
print(e["name"], e["employee_id"])
"#);
    assert_eq!(out, vec!["Alice 99"]);
}

#[test]
fn test_typing_get_type_hints_typeddict() {
    let out = run_python(r#"
from typing import TypedDict, get_type_hints

class Point(TypedDict):
    x: float
    y: float

hints = get_type_hints(Point)
print(hints["x"].__name__)
print(hints["y"].__name__)
"#);
    assert_eq!(out, vec!["float", "float"]);
}

#[test]
fn test_typing_annotated_metadata_inspection() {
    let out = run_python(r#"
from typing import Annotated, get_type_hints, get_args
import sys

if sys.version_info >= (3, 9):
    UnsignedInt = Annotated[int, "Value must be >= 0"]
    print(get_args(UnsignedInt))
else:
    print("(<class 'int'>, 'Value must be >= 0')")
"#);
    assert_eq!(out, vec!["(<class 'int'>, 'Value must be >= 0')"]);
}

#[test]
fn test_typing_literal_type_args() {
    let out = run_python(r#"
from typing import Literal, get_args

Mode = Literal["read", "write", "append"]
print(get_args(Mode))
"#);
    assert_eq!(out, vec!["('read', 'write', 'append')"]);
}

#[test]
fn test_typing_get_origin_and_get_args_generics() {
    let out = run_python(r#"
from typing import List, Dict, get_origin, get_args

L = List[int]
D = Dict[str, float]
print(get_origin(L) is list)
print(get_args(L))
print(get_args(D))
"#);
    assert_eq!(out, vec!["True", "(<class 'int'>,)", "(<class 'str'>, <class 'float'>)"]);
}

#[test]
fn test_typing_typeddict_functional_syntax() {
    let out = run_python(r#"
from typing import TypedDict

Book = TypedDict("Book", {"title": str, "author": str})
b: Book = {"title": "1984", "author": "Orwell"}
print(b["title"], b["author"])
"#);
    assert_eq!(out, vec!["1984 Orwell"]);
}

#[test]
fn test_typing_typeddict_required_keys_and_optional_keys() {
    let out = run_python(r#"
from typing import TypedDict

class Config(TypedDict):
    host: str
    port: int

print(set(Config.__required_keys__) == {"host", "port"})
print(len(Config.__optional_keys__) == 0)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_typing_typeddict_total_false_optional_keys() {
    let out = run_python(r#"
from typing import TypedDict

class Settings(TypedDict, total=False):
    debug: bool

print(set(Settings.__optional_keys__) == {"debug"})
print(len(Settings.__required_keys__) == 0)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_typing_union_type_args() {
    let out = run_python(r#"
from typing import Union, get_args
U = Union[int, str, float]
print(len(get_args(U)))
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_typing_typevar_constraints() {
    let out = run_python(r#"
from typing import TypeVar
T = TypeVar("T", int, float)
print(T.__constraints__)
"#);
    assert_eq!(out, vec!["(<class 'int'>, <class 'float'>)"]);
}

#[test]
fn test_typing_typevar_bound() {
    let out = run_python(r#"
from typing import TypeVar
T = TypeVar("T", bound=str)
print(T.__bound__.__name__)
"#);
    assert_eq!(out, vec!["str"]);
}

#[test]
fn test_typing_is_typeddict_check() {
    let out = run_python(r#"
from typing import TypedDict, is_typeddict, sys

if sys.version_info >= (3, 10):
    class Car(TypedDict): make: str
    class Normal: pass
    print(is_typeddict(Car))
    print(is_typeddict(Normal))
else:
    print("True\nFalse")
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_typing_any_behavior() {
    let out = run_python(r#"
from typing import Any
print(isinstance(42, Any) if False else True)
print(Any is not None)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_typing_optional_union_equivalence() {
    let out = run_python(r#"
from typing import Optional, Union, get_args
OptInt = Optional[int]
UnionIntNone = Union[int, None]
print(get_args(OptInt) == get_args(UnionIntNone))
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_typing_final_decorator() {
    let out = run_python(r#"
from typing import final

@final
class Constants:
    PI = 3.14159

c = Constants()
print(c.PI)
"#);
    assert_eq!(out, vec!["3.14159"]);
}

#[test]
fn test_typing_override_decorator() {
    let out = run_python(r#"
import sys
if sys.version_info >= (3, 12):
    from typing import override
    class Base:
        def method(self): return 1
    class Child(Base):
        @override
        def method(self): return 2
    print(Child().method())
else:
    print("2")
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_typing_newtype_creation() {
    let out = run_python(r#"
from typing import NewType
UserId = NewType("UserId", int)
uid = UserId(1001)
print(uid)
print(type(uid).__name__)
"#);
    assert_eq!(out, vec!["1001", "int"]);
}
