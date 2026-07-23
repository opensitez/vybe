use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: typing — TypeVar, Generic, Protocol, Literal, Union, overload, TypeGuard
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_typing_optional_union_isinstance() {
    let src = r#"
from typing import Optional, Union

def greet(name: Optional[str] = None) -> str:
    return f"Hello, {name}" if name else "Hello!"

print(greet("Alice"))
print(greet())

def process(x: Union[int, str]) -> str:
    return f"{type(x).__name__}:{x}"

print(process(42))
print(process("hello"))
"#;
    assert_eq!(
        run_python(src),
        vec!["Hello, Alice", "Hello!", "int:42", "str:hello"]
    );
}

#[test]
fn test_py_typing_typevar_generic_function() {
    let src = r#"
from typing import TypeVar, List

T = TypeVar('T')

def first(items: List[T]) -> T:
    return items[0]

print(first([1, 2, 3]))
print(first(["a", "b"]))
print(type(first([1, 2])).__name__)
"#;
    assert_eq!(run_python(src), vec!["1", "a", "int"]);
}

#[test]
fn test_py_typing_generic_class() {
    let src = r#"
from typing import TypeVar, Generic, List

T = TypeVar('T')

class Stack(Generic[T]):
    def __init__(self):
        self._items: List[T] = []
    def push(self, item: T) -> None:
        self._items.append(item)
    def pop(self) -> T:
        return self._items.pop()
    def __len__(self) -> int:
        return len(self._items)

s = Stack()
s.push(1)
s.push(2)
print(s.pop())
print(len(s))
"#;
    assert_eq!(run_python(src), vec!["2", "1"]);
}

#[test]
fn test_py_typing_protocol_structural_subtyping() {
    let src = r#"
from typing import Protocol

class Drawable(Protocol):
    def draw(self) -> str: ...

class Circle:
    def draw(self) -> str:
        return "Drawing circle"

class Square:
    def draw(self) -> str:
        return "Drawing square"

def render(shape: Drawable) -> str:
    return shape.draw()

print(render(Circle()))
print(render(Square()))
"#;
    assert_eq!(run_python(src), vec!["Drawing circle", "Drawing square"]);
}

#[test]
fn test_py_typing_runtime_checkable_protocol() {
    let src = r#"
from typing import Protocol, runtime_checkable

@runtime_checkable
class Sized(Protocol):
    def __len__(self) -> int: ...

print(isinstance([], Sized))
print(isinstance({}, Sized))
print(isinstance(42, Sized))
"#;
    assert_eq!(run_python(src), vec!["True", "True", "False"]);
}

#[test]
fn test_py_typing_literal_type() {
    let src = r#"
from typing import Literal

Direction = Literal["north", "south", "east", "west"]

def move(direction: Direction) -> str:
    return f"Moving {direction}"

print(move("north"))
print(move("east"))
"#;
    assert_eq!(run_python(src), vec!["Moving north", "Moving east"]);
}

#[test]
fn test_py_typing_typeddict_keys_and_access() {
    let src = r#"
from typing import TypedDict

class Movie(TypedDict):
    title: str
    year: int
    rating: float

m: Movie = {"title": "Inception", "year": 2010, "rating": 8.8}
print(m["title"])
print(m["year"])
"#;
    assert_eq!(run_python(src), vec!["Inception", "2010"]);
}

#[test]
fn test_py_typing_typeddict_total_false() {
    let src = r#"
from typing import TypedDict

class PartialConfig(TypedDict, total=False):
    debug: bool
    timeout: int
    log_level: str

config: PartialConfig = {"debug": True}
print(config.get("debug"))
print(config.get("timeout", 30))
"#;
    assert_eq!(run_python(src), vec!["True", "30"]);
}

#[test]
fn test_py_typing_get_type_hints() {
    let src = r#"
from typing import get_type_hints, Optional

class User:
    name: str
    age: Optional[int]

    def __init__(self, name: str, age: Optional[int] = None) -> None:
        self.name = name
        self.age = age

hints = get_type_hints(User)
print('name' in hints)
print('age' in hints)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_typing_callable_type_annotation() {
    let src = r#"
from typing import Callable

def apply(func: Callable[[int, int], int], a: int, b: int) -> int:
    return func(a, b)

print(apply(lambda x, y: x + y, 3, 4))
print(apply(lambda x, y: x * y, 3, 4))
"#;
    assert_eq!(run_python(src), vec!["7", "12"]);
}

#[test]
fn test_py_typing_any_bypasses_type_checks() {
    let src = r#"
from typing import Any

def identity(x: Any) -> Any:
    return x

print(identity(42))
print(identity("hello"))
print(identity([1, 2, 3]))
"#;
    assert_eq!(run_python(src), vec!["42", "hello", "[1, 2, 3]"]);
}

#[test]
fn test_py_typing_overload_dispatch() {
    let src = r#"
from typing import overload

@overload
def process(x: int) -> str: ...
@overload
def process(x: str) -> int: ...

def process(x):
    if isinstance(x, int):
        return str(x)
    return len(x)

print(process(42))
print(process("hello"))
"#;
    assert_eq!(run_python(src), vec!["42", "5"]);
}

#[test]
fn test_py_typing_classvar_and_final() {
    let src = r#"
from typing import ClassVar, Final

class Config:
    MAX_RETRIES: ClassVar[int] = 3
    DEFAULT_HOST: Final = "localhost"

print(Config.MAX_RETRIES)
print(Config.DEFAULT_HOST)
"#;
    assert_eq!(run_python(src), vec!["3", "localhost"]);
}

#[test]
fn test_py_typing_new_union_syntax_py310() {
    let src = r#"
import sys

def describe(x: int | str | None) -> str:
    if x is None:
        return "none"
    return f"{type(x).__name__}:{x}"

print(describe(10))
print(describe("hi"))
print(describe(None))
"#;
    assert_eq!(run_python(src), vec!["int:10", "str:hi", "none"]);
}

#[test]
fn test_py_typing_annotated_metadata() {
    let src = r#"
from typing import Annotated, get_type_hints
import sys

Positive = Annotated[int, "must be > 0"]

def validate(x: Positive) -> bool:
    return x > 0

print(validate(5))
print(validate(-1))
"#;
    assert_eq!(run_python(src), vec!["True", "False"]);
}

#[test]
fn test_py_typing_namedtuple_with_defaults() {
    let src = r#"
from typing import NamedTuple

class Point(NamedTuple):
    x: float
    y: float
    z: float = 0.0

p1 = Point(1.0, 2.0)
p2 = Point(1.0, 2.0, 3.0)
print(p1)
print(p2)
print(p1._asdict())
"#;
    assert_eq!(
        run_python(src),
        vec![
            "Point(x=1.0, y=2.0, z=0.0)",
            "Point(x=1.0, y=2.0, z=3.0)",
            "{'x': 1.0, 'y': 2.0, 'z': 0.0}"
        ]
    );
}

#[test]
fn test_py_typing_typeguard() {
    let src = r#"
from typing import Union

def is_string(val: Union[str, int]) -> bool:
    return isinstance(val, str)

items = [1, "hello", 2, "world"]
strings = [x for x in items if is_string(x)]
print(strings)
"#;
    assert_eq!(run_python(src), vec!["['hello', 'world']"]);
}

#[test]
fn test_py_typing_tuple_fixed_and_variable() {
    let src = r#"
from typing import Tuple

def swap(pair: Tuple[int, str]) -> Tuple[str, int]:
    return pair[1], pair[0]

a, b = swap((1, "hello"))
print(a, b)
print(type(a).__name__, type(b).__name__)
"#;
    assert_eq!(run_python(src), vec!["hello 1", "str int"]);
}

#[test]
fn test_py_typing_typevar_bound_constraint() {
    let src = r#"
from typing import TypeVar

Numeric = TypeVar('Numeric', int, float, complex)

def square(x: Numeric) -> Numeric:
    return x * x

print(square(3))
print(square(2.5))
"#;
    assert_eq!(run_python(src), vec!["9", "6.25"]);
}

#[test]
fn test_py_typing_self_type_reference() {
    let src = r#"
from __future__ import annotations

class Builder:
    def __init__(self):
        self._items = []

    def add(self, item: str) -> Builder:
        self._items.append(item)
        return self

    def build(self) -> list:
        return self._items

result = Builder().add("a").add("b").add("c").build()
print(result)
"#;
    assert_eq!(run_python(src), vec!["['a', 'b', 'c']"]);
}
