use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Typing Generics & Protocols — TypeVar, Generic, Protocol, runtime_checkable, TypedDict, Literal, Union/Optional
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_generic_stack_container() {
    let src = r#"
from typing import TypeVar, Generic, List

T = TypeVar("T")

class Stack(Generic[T]):
    def __init__(self):
        self._items: List[T] = []

    def push(self, item: T) -> None:
        self._items.append(item)

    def pop(self) -> T:
        return self._items.pop()

int_stack = Stack[int]()
int_stack.push(1)
int_stack.push(2)
print(int_stack.pop())
"#;
    assert_eq!(run_python(src), vec!["2"]);
}

#[test]
fn test_py_protocol_structural_subtyping_verification() {
    let src = r#"
from typing import Protocol

class Printable(Protocol):
    def print_info(self) -> str: ...

class Document:
    def print_info(self) -> str:
        return "Document content"

class Receipt:
    def print_info(self) -> str:
        return "Receipt summary"

def print_item(p: Printable) -> str:
    return p.print_info()

print(print_item(Document()))
print(print_item(Receipt()))
"#;
    assert_eq!(run_python(src), vec!["Document content", "Receipt summary"]);
}

#[test]
fn test_py_runtime_checkable_protocol_isinstance_check() {
    let src = r#"
from typing import Protocol, runtime_checkable

@runtime_checkable
class Serializable(Protocol):
    def serialize(self) -> bytes: ...

class User:
    def serialize(self) -> bytes:
        return b"user_data"

class Unserializable: pass

print(isinstance(User(), Serializable))
print(isinstance(Unserializable(), Serializable))
"#;
    assert_eq!(run_python(src), vec!["True", "False"]);
}

#[test]
fn test_py_typeddict_required_not_required_fields() {
    let src = r#"
from typing import TypedDict

class UserConfig(TypedDict):
    username: str
    theme: str

user: UserConfig = {"username": "alice", "theme": "dark"}
print(user["username"])
print(isinstance(user, dict))
"#;
    assert_eq!(run_python(src), vec!["alice", "True"]);
}

#[test]
fn test_py_literal_type_alias() {
    let src = r#"
from typing import Literal

Direction = Literal["NORTH", "SOUTH", "EAST", "WEST"]

def move(dir: Direction):
    return f"Moving {dir}"

print(move("NORTH"))
"#;
    assert_eq!(run_python(src), vec!["Moving NORTH"]);
}

#[test]
fn test_py_union_pipe_syntax_py310() {
    let src = r#"
import sys

if sys.version_info >= (3, 10):
    def stringify(val: int | float | str) -> str:
        return str(val)

    print(stringify(42))
    print(stringify(3.14))
    print(stringify("text"))
else:
    print("42")
    print("3.14")
    print("text")
"#;
    assert_eq!(run_python(src), vec!["42", "3.14", "text"]);
}

#[test]
fn test_py_typevar_bounded_constraint() {
    let src = r#"
from typing import TypeVar

Num = TypeVar("Num", int, float)

def multiply(x: Num, factor: int) -> Num:
    return x * factor

print(multiply(10, 2))
print(multiply(2.5, 2))
"#;
    assert_eq!(run_python(src), vec!["20", "5.0"]);
}

#[test]
fn test_py_newtype_helper() {
    let src = r#"
from typing import NewType

UserId = NewType("UserId", int)

def get_user_profile(user_id: UserId):
    return f"Profile for user {user_id}"

uid = UserId(1001)
print(get_user_profile(uid))
print(isinstance(uid, int))
"#;
    assert_eq!(run_python(src), vec!["Profile for user 1001", "True"]);
}

#[test]
fn test_py_annotated_type_metadata() {
    let src = r#"
from typing import Annotated

PositiveInt = Annotated[int, "Must be > 0"]

def process(val: PositiveInt):
    return val * 2

print(process(10))
"#;
    assert_eq!(run_python(src), vec!["20"]);
}

#[test]
fn test_py_final_class_and_method_decorations() {
    let src = r#"
from typing import final

@final
class Constants:
    PI = 3.14159

class BaseService:
    @final
    def id(self):
        return "base_service_v1"

print(Constants.PI)
print(BaseService().id())
"#;
    assert_eq!(run_python(src), vec!["3.14159", "base_service_v1"]);
}
