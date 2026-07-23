use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: structural pattern matching (match/case) — literals, sequences, mappings, class patterns, guards, OR patterns
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_match_literal_patterns() {
    let src = r#"
import sys

def classify(val):
    match val:
        case 0:
            return "zero"
        case 1 | 2 | 3:
            return "one to three"
        case "hello":
            return "greeting"
        case _:
            return "other"

print(classify(0))
print(classify(2))
print(classify("hello"))
print(classify(99))
"#;
    assert_eq!(
        run_python(src),
        vec!["zero", "one to three", "greeting", "other"]
    );
}

#[test]
fn test_py_match_sequence_patterns() {
    let src = r#"
def describe(seq):
    match seq:
        case []:
            return "empty"
        case [x]:
            return f"one: {x}"
        case [x, y]:
            return f"two: {x}, {y}"
        case [x, *rest]:
            return f"first: {x}, rest: {rest}"

print(describe([]))
print(describe([42]))
print(describe([1, 2]))
print(describe([1, 2, 3, 4]))
"#;
    assert_eq!(
        run_python(src),
        vec!["empty", "one: 42", "two: 1, 2", "first: 1, rest: [2, 3, 4]"]
    );
}

#[test]
fn test_py_match_mapping_patterns() {
    let src = r#"
def handle_event(event):
    match event:
        case {"type": "click", "x": x, "y": y}:
            return f"click at ({x}, {y})"
        case {"type": "keypress", "key": k}:
            return f"key: {k}"
        case {"type": t}:
            return f"unknown type: {t}"
        case _:
            return "invalid event"

print(handle_event({"type": "click", "x": 10, "y": 20}))
print(handle_event({"type": "keypress", "key": "Enter"}))
print(handle_event({"type": "scroll"}))
"#;
    assert_eq!(
        run_python(src),
        vec!["click at (10, 20)", "key: Enter", "unknown type: scroll"]
    );
}

#[test]
fn test_py_match_class_patterns() {
    let src = r#"
from dataclasses import dataclass

@dataclass
class Point:
    x: float
    y: float

@dataclass
class Circle:
    center: Point
    radius: float

def describe_shape(shape):
    match shape:
        case Point(x=0, y=0):
            return "origin"
        case Point(x=x, y=y):
            return f"point at ({x}, {y})"
        case Circle(center=Point(x=0, y=0), radius=r):
            return f"centered circle r={r}"
        case _:
            return "unknown"

print(describe_shape(Point(0, 0)))
print(describe_shape(Point(3, 4)))
print(describe_shape(Circle(Point(0, 0), 5)))
"#;
    assert_eq!(
        run_python(src),
        vec!["origin", "point at (3.0, 4.0)", "centered circle r=5.0"]
    );
}

#[test]
fn test_py_match_guard_clauses() {
    let src = r#"
def classify_number(n):
    match n:
        case x if x < 0:
            return "negative"
        case 0:
            return "zero"
        case x if x % 2 == 0:
            return "positive even"
        case _:
            return "positive odd"

for n in [-5, 0, 4, 7]:
    print(classify_number(n))
"#;
    assert_eq!(
        run_python(src),
        vec!["negative", "zero", "positive even", "positive odd"]
    );
}

#[test]
fn test_py_match_or_patterns() {
    let src = r#"
def describe_status(code):
    match code:
        case 200 | 201 | 202:
            return "success"
        case 301 | 302:
            return "redirect"
        case 400:
            return "bad request"
        case 404:
            return "not found"
        case 500 | 503:
            return "server error"
        case _:
            return "unknown"

for code in [200, 302, 404, 500, 999]:
    print(describe_status(code))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "success",
            "redirect",
            "not found",
            "server error",
            "unknown"
        ]
    );
}

#[test]
fn test_py_match_as_pattern_capture() {
    let src = r#"
def process(val):
    match val:
        case [1, *rest] as whole:
            return f"starts with 1, whole={whole}, rest={rest}"
        case {"key": str(s)} as d:
            return f"has key={s}"
        case _:
            return "no match"

print(process([1, 2, 3]))
print(process({"key": "value"}))
print(process("other"))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "starts with 1, whole=[1, 2, 3], rest=[2, 3]",
            "has key=value",
            "no match"
        ]
    );
}

#[test]
fn test_py_match_nested_patterns() {
    let src = r#"
def describe_matrix(m):
    match m:
        case [[a, b], [c, d]]:
            return f"2x2: {a},{b},{c},{d}"
        case [[a, *_], *_]:
            return f"matrix starting with {a}"
        case _:
            return "unknown"

print(describe_matrix([[1, 2], [3, 4]]))
print(describe_matrix([[5, 6, 7], [8, 9, 10]]))
"#;
    assert_eq!(
        run_python(src),
        vec!["2x2: 1,2,3,4", "matrix starting with 5"]
    );
}

#[test]
fn test_py_match_exhaustive_enum() {
    let src = r#"
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
"#;
    assert_eq!(
        run_python(src),
        vec!["waiting", "running", "completed", "failed"]
    );
}

#[test]
fn test_py_match_type_check_patterns() {
    let src = r#"
def describe_type(val):
    match val:
        case str(s):
            return f"string: {s!r}"
        case int(n) if n > 0:
            return f"positive int: {n}"
        case int(n):
            return f"non-positive int: {n}"
        case list(items):
            return f"list of {len(items)}"
        case _:
            return "other"

print(describe_type("hello"))
print(describe_type(42))
print(describe_type(-5))
print(describe_type([1, 2, 3]))
print(describe_type(3.14))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "string: 'hello'",
            "positive int: 42",
            "non-positive int: -5",
            "list of 3",
            "other"
        ]
    );
}

#[test]
fn test_py_match_walrus_equivalent_with_as() {
    let src = r#"
commands = [
    ["quit"],
    ["go", "north"],
    ["pick", "key", "rusty"],
    ["look"],
]

for cmd in commands:
    match cmd:
        case ["quit"]:
            print("Quitting")
        case ["go", direction]:
            print(f"Going {direction}")
        case ["pick", item, *adjectives]:
            print(f"Picking {' '.join(adjectives)} {item}")
        case [verb]:
            print(f"Unknown action: {verb}")
"#;
    assert_eq!(
        run_python(src),
        vec![
            "Quitting",
            "Going north",
            "Picking rusty key",
            "Unknown action: look"
        ]
    );
}
