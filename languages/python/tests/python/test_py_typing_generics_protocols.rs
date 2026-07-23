use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Typing, Generics & Protocols — Generic, TypeVar, Protocol, runtime_checkable, structural subtyping
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_generic_class_typevar() {
    let src = r#"
from typing import TypeVar, Generic, List

T = TypeVar("T")

class Stack(Generic[T]):
    def __init__(self):
        self.items: List[T] = []

    def push(self, item: T) -> None:
        self.items.append(item)

    def pop(self) -> T:
        return self.items.pop()

s = Stack[int]()
s.push(10)
s.push(20)
print(s.pop())
print(s.pop())
"#;
    assert_eq!(run_python(src), vec!["20", "10"]);
}

#[test]
fn test_py_protocol_structural_subtyping() {
    let src = r#"
from typing import Protocol

class Renderable(Protocol):
    def render(self) -> str: ...

class HTMLWidget:
    def render(self) -> str:
        return "<div>Widget</div>"

class TextWidget:
    def render(self) -> str:
        return "Widget Text"

def display(r: Renderable) -> str:
    return r.render()

print(display(HTMLWidget()))
print(display(TextWidget()))
"#;
    assert_eq!(run_python(src), vec!["<div>Widget</div>", "Widget Text"]);
}

#[test]
fn test_py_runtime_checkable_protocol_isinstance() {
    let src = r#"
from typing import Protocol, runtime_checkable

@runtime_checkable
class Closable(Protocol):
    def close(self) -> None: ...

class FileStream:
    def close(self): pass

class NonClosable: pass

print(isinstance(FileStream(), Closable))
print(isinstance(NonClosable(), Closable))
"#;
    assert_eq!(run_python(src), vec!["True", "False"]);
}

#[test]
fn test_py_typevar_bounded_and_constrained() {
    let src = r#"
from typing import TypeVar

Num = TypeVar("Num", int, float)

def add(x: Num, y: Num) -> Num:
    return x + y

print(add(1, 2))
print(add(1.5, 2.5))
"#;
    assert_eq!(run_python(src), vec!["3", "4.0"]);
}

#[test]
fn test_py_typeddict_optional_total_fields() {
    let src = r#"
from typing import TypedDict

class UserProfile(TypedDict):
    id: int
    name: str

u: UserProfile = {"id": 1, "name": "Alice"}
print(u["name"])
print(isinstance(u, dict))
"#;
    assert_eq!(run_python(src), vec!["Alice", "True"]);
}

#[test]
fn test_py_literal_type_annotations() {
    let src = r#"
from typing import Literal

Mode = Literal["r", "w", "a"]

def open_file(path: str, mode: Mode):
    return f"Opening {path} in mode {mode}"

print(open_file("test.txt", "r"))
"#;
    assert_eq!(run_python(src), vec!["Opening test.txt in mode r"]);
}

#[test]
fn test_py_union_optional_alias() {
    let src = r#"
from typing import Union, Optional

def process(val: Optional[Union[int, str]]) -> str:
    if val is None:
        return "None"
    return f"Val: {val}"

print(process(10))
print(process("hello"))
print(process(None))
"#;
    assert_eq!(run_python(src), vec!["Val: 10", "Val: hello", "None"]);
}

#[test]
fn test_py_callable_type_annotation() {
    let src = r#"
from typing import Callable

def apply_twice(func: Callable[[int], int], val: int) -> int:
    return func(func(val))

print(apply_twice(lambda x: x + 1, 5))
"#;
    assert_eq!(run_python(src), vec!["7"]);
}

#[test]
fn test_py_any_and_no_return_annotations() {
    let src = r#"
from typing import Any, NoReturn

def log(msg: Any) -> None:
    print(f"LOG: {msg}")

log("test message")
log(123)
"#;
    assert_eq!(run_python(src), vec!["LOG: test message", "LOG: 123"]);
}

#[test]
fn test_py_final_class_and_method_decorator() {
    let src = r#"
from typing import final

@final
class LeafClass:
    pass

class Base:
    @final
    def lock(self):
        return "locked"

print(Base().lock())
"#;
    assert_eq!(run_python(src), vec!["locked"]);
}
