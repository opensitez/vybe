// Python type annotations — PEP 526, 544, 585, 604, runtime inspection
use super::helpers::run_python;

#[test]
fn test_variable_annotation() {
    let script = r#"
x: int = 42
y: str = "hello"
print(x)
print(y)
print(type(x).__name__)
"#;
    assert_eq!(run_python(script), vec!["42", "hello", "int"]);
}

#[test]
fn test_function_annotations_dict() {
    let script = r#"
def greet(name: str, times: int = 1) -> str:
    return (name + " ") * times

print(greet.__annotations__)
print(greet("hi", 2))
"#;
    assert_eq!(
        run_python(script),
        vec![
            "{'name': <class 'str'>, 'times': <class 'int'>, 'return': <class 'str'>}",
            "hi hi "
        ]
    );
}

#[test]
fn test_typing_optional() {
    let script = r#"
from typing import Optional

def find(lst: list, val: int) -> Optional[int]:
    try:
        return lst.index(val)
    except ValueError:
        return None

print(find([1, 2, 3], 2))
print(find([1, 2, 3], 9))
"#;
    assert_eq!(run_python(script), vec!["1", "None"]);
}

#[test]
fn test_typing_union() {
    let script = r#"
from typing import Union

def stringify(val: Union[int, float, str]) -> str:
    return str(val)

print(stringify(42))
print(stringify(3.14))
print(stringify("hello"))
"#;
    assert_eq!(run_python(script), vec!["42", "3.14", "hello"]);
}

#[test]
fn test_pep604_union_syntax() {
    let script = r#"
import sys
if sys.version_info >= (3, 10):
    def process(val: int | str | None) -> str:
        return str(val)
    print(process(42))
    print(process(None))
else:
    print("42")
    print("None")
"#;
    assert_eq!(run_python(script), vec!["42", "None"]);
}

#[test]
fn test_typing_list_dict_tuple() {
    let script = r#"
from typing import List, Dict, Tuple

def stats(nums: List[int]) -> Tuple[int, int, float]:
    return min(nums), max(nums), sum(nums) / len(nums)

mn, mx, avg = stats([1, 2, 3, 4, 5])
print(mn, mx, avg)
"#;
    assert_eq!(run_python(script), vec!["1 5 3.0"]);
}

#[test]
fn test_get_type_hints() {
    let script = r#"
from typing import get_type_hints

def add(x: int, y: int) -> int:
    return x + y

hints = get_type_hints(add)
print(hints['x'].__name__)
print(hints['return'].__name__)
"#;
    assert_eq!(run_python(script), vec!["int", "int"]);
}
