// Python match statement with guards — value, class, sequence, mapping patterns + guard clauses
use super::helpers::run_python;

#[test]
fn test_match_guard_condition() {
    let script = r#"
def classify(n):
    match n:
        case x if x < 0:
            return "negative"
        case 0:
            return "zero"
        case x if x % 2 == 0:
            return "positive even"
        case _:
            return "positive odd"

print(classify(-5))
print(classify(0))
print(classify(4))
print(classify(7))
"#;
    assert_eq!(
        run_python(script),
        vec!["negative", "zero", "positive even", "positive odd"]
    );
}

#[test]
fn test_match_sequence_pattern() {
    let script = r#"
def describe(lst):
    match lst:
        case []:
            return "empty"
        case [x]:
            return f"one: {x}"
        case [x, y]:
            return f"two: {x}, {y}"
        case [x, *rest]:
            return f"many: first={x}, rest={rest}"

print(describe([]))
print(describe([1]))
print(describe([1, 2]))
print(describe([1, 2, 3, 4]))
"#;
    assert_eq!(
        run_python(script),
        vec![
            "empty",
            "one: 1",
            "two: 1, 2",
            "many: first=1, rest=[2, 3, 4]"
        ]
    );
}

#[test]
fn test_match_mapping_pattern() {
    let script = r#"
def process(cmd):
    match cmd:
        case {"action": "move", "x": x, "y": y}:
            return f"move to ({x}, {y})"
        case {"action": "quit"}:
            return "quit"
        case _:
            return "unknown"

print(process({"action": "move", "x": 3, "y": 7}))
print(process({"action": "quit"}))
print(process({"action": "fire"}))
"#;
    assert_eq!(
        run_python(script),
        vec!["move to (3, 7)", "quit", "unknown"]
    );
}

#[test]
fn test_match_or_pattern() {
    let script = r#"
for status in [200, 201, 404, 500]:
    match status:
        case 200 | 201:
            print("ok")
        case 404:
            print("not found")
        case _:
            print("error")
"#;
    assert_eq!(run_python(script), vec!["ok", "ok", "not found", "error"]);
}

#[test]
fn test_match_class_pattern() {
    let script = r#"
class Point:
    def __init__(self, x, y):
        self.x = x
        self.y = y

def describe_point(p):
    match p:
        case Point(x=0, y=0):
            return "origin"
        case Point(x=0, y=y):
            return f"y-axis at {y}"
        case Point(x=x, y=0):
            return f"x-axis at {x}"
        case Point(x=x, y=y):
            return f"({x}, {y})"

print(describe_point(Point(0, 0)))
print(describe_point(Point(0, 5)))
print(describe_point(Point(3, 0)))
print(describe_point(Point(2, 4)))
"#;
    assert_eq!(
        run_python(script),
        vec!["origin", "y-axis at 5", "x-axis at 3", "(2, 4)"]
    );
}

#[test]
fn test_match_wildcard_capture() {
    let script = r#"
commands = ["start", "stop", "unknown"]
for cmd in commands:
    match cmd:
        case "start":
            print("starting")
        case "stop":
            print("stopping")
        case other:
            print(f"got: {other}")
"#;
    assert_eq!(
        run_python(script),
        vec!["starting", "stopping", "got: unknown"]
    );
}
