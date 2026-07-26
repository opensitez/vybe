// Python walrus operator (:=) — scope, while, comprehensions, conditionals
use super::helpers::run_python;

#[test]
fn test_walrus_in_while() {
    let script = r#"
import io
stream = io.StringIO("line1\nline2\nline3\n")
lines = []
while line := stream.readline():
    lines.append(line.strip())
print(lines)
"#;
    assert_eq!(run_python(script), vec!["['line1', 'line2', 'line3']"]);
}

#[test]
fn test_walrus_in_if() {
    let script = r#"
data = [1, 2, 3, 4, 5]
if (n := len(data)) > 3:
    print(f"long list: {n}")
else:
    print(f"short list: {n}")
"#;
    assert_eq!(run_python(script), vec!["long list: 5"]);
}

#[test]
fn test_walrus_in_comprehension() {
    let script = r#"
vals = [1, -2, 3, -4, 5]
positive = [y for x in vals if (y := x * 2) > 0]
print(positive)
"#;
    assert_eq!(run_python(script), vec!["[2, 6, 10]"]);
}

#[test]
fn test_walrus_avoids_double_call() {
    let script = r#"
calls = 0
def expensive():
    global calls
    calls += 1
    return calls * 10

results = []
if (v := expensive()) > 5:
    results.append(v)

print(calls)
print(results)
"#;
    assert_eq!(run_python(script), vec!["1", "[10]"]);
}

#[test]
fn test_walrus_escapes_comprehension_scope() {
    let script = r#"
last = None
data = [1, 2, 3, 4, 5]
result = [last := x for x in data]
print(last)
print(result)
"#;
    assert_eq!(run_python(script), vec!["5", "[1, 2, 3, 4, 5]"]);
}

#[test]
fn test_walrus_chained() {
    let script = r#"
text = "hello world"
if (words := text.split()) and (count := len(words)) > 1:
    print(count)
    print(words[0])
"#;
    assert_eq!(run_python(script), vec!["2", "hello"]);
}

#[test]
fn test_walrus_in_any_all() {
    let script = r#"
data = [0, 0, 3, 0, 5]
found = any((first_nonzero := x) for x in data if x != 0)
print(found)
print(first_nonzero)
"#;
    assert_eq!(run_python(script), vec!["True", "3"]);
}
