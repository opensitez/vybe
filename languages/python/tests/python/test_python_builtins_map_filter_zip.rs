// Python builtins: map, filter, zip, enumerate, zip_longest, chain
use super::helpers::run_python;

#[test]
fn test_map_basic() {
    let script = r#"
result = list(map(lambda x: x * 2, [1, 2, 3, 4]))
print(result)
"#;
    assert_eq!(run_python(script), vec!["[2, 4, 6, 8]"]);
}

#[test]
fn test_map_multiple_iterables() {
    let script = r#"
result = list(map(lambda x, y: x + y, [1, 2, 3], [10, 20, 30]))
print(result)
"#;
    assert_eq!(run_python(script), vec!["[11, 22, 33]"]);
}

#[test]
fn test_filter_basic() {
    let script = r#"
result = list(filter(lambda x: x % 2 == 0, range(10)))
print(result)
"#;
    assert_eq!(run_python(script), vec!["[0, 2, 4, 6, 8]"]);
}

#[test]
fn test_filter_none_removes_falsy() {
    let script = r#"
result = list(filter(None, [0, 1, "", "hi", None, False, True]))
print(result)
"#;
    assert_eq!(run_python(script), vec!["[1, 'hi', True]"]);
}

#[test]
fn test_zip_basic() {
    let script = r#"
pairs = list(zip([1, 2, 3], ['a', 'b', 'c']))
print(pairs)
"#;
    assert_eq!(run_python(script), vec!["[(1, 'a'), (2, 'b'), (3, 'c')]"]);
}

#[test]
fn test_zip_stops_at_shortest() {
    let script = r#"
pairs = list(zip([1, 2, 3, 4], ['a', 'b']))
print(pairs)
"#;
    assert_eq!(run_python(script), vec!["[(1, 'a'), (2, 'b')]"]);
}

#[test]
fn test_zip_strict_raises_on_mismatch() {
    let script = r#"
try:
    list(zip([1, 2, 3], [1, 2], strict=True))
    print("no_error")
except ValueError:
    print("ValueError")
"#;
    assert_eq!(run_python(script), vec!["ValueError"]);
}

#[test]
fn test_enumerate_with_start() {
    let script = r#"
for i, v in enumerate(['a', 'b', 'c'], start=5):
    print(i, v)
"#;
    assert_eq!(run_python(script), vec!["5 a", "6 b", "7 c"]);
}

#[test]
fn test_map_lazy_evaluation() {
    let script = r#"
calls = []
def track(x):
    calls.append(x)
    return x * 10

m = map(track, [1, 2, 3])
print(len(calls))  # 0 — not called yet
next(m)
print(len(calls))  # 1
"#;
    assert_eq!(run_python(script), vec!["0", "1"]);
}

#[test]
fn test_zip_three_iterables() {
    let script = r#"
a = [1, 2]
b = ['x', 'y']
c = [True, False]
print(list(zip(a, b, c)))
"#;
    assert_eq!(
        run_python(script),
        vec!["[(1, 'x', True), (2, 'y', False)]"]
    );
}
