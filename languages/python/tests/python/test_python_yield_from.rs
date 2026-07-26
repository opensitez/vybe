// Python yield from — delegation, send pass-through, return values, StopIteration
use super::helpers::run_python;

#[test]
fn test_yield_from_basic_delegation() {
    let script = r#"
def inner():
    yield 1
    yield 2
    yield 3

def outer():
    yield 0
    yield from inner()
    yield 4

print(list(outer()))
"#;
    assert_eq!(run_python(script), vec!["[0, 1, 2, 3, 4]"]);
}

#[test]
fn test_yield_from_return_value() {
    let script = r#"
def inner():
    yield 1
    return "inner_done"

def outer():
    result = yield from inner()
    yield result

print(list(outer()))
"#;
    assert_eq!(run_python(script), vec!["[1, 'inner_done']"]);
}

#[test]
fn test_yield_from_send_passthrough() {
    let script = r#"
def inner():
    received = yield "from_inner"
    yield f"inner_got:{received}"

def outer():
    yield from inner()

g = outer()
print(next(g))
print(g.send("hello"))
"#;
    assert_eq!(run_python(script), vec!["from_inner", "inner_got:hello"]);
}

#[test]
fn test_yield_from_any_iterable() {
    let script = r#"
def gen():
    yield from range(3)
    yield from "abc"
    yield from [10, 20]

print(list(gen()))
"#;
    assert_eq!(run_python(script), vec!["[0, 1, 2, 'a', 'b', 'c', 10, 20]"]);
}

#[test]
fn test_yield_from_chained_generators() {
    let script = r#"
def counter(start, stop):
    while start < stop:
        yield start
        start += 1

def merged(*ranges):
    for r in ranges:
        yield from r

result = list(merged(counter(0, 3), counter(10, 13)))
print(result)
"#;
    assert_eq!(run_python(script), vec!["[0, 1, 2, 10, 11, 12]"]);
}

#[test]
fn test_yield_from_exception_propagation() {
    let script = r#"
def inner():
    try:
        yield 1
        yield 2
    except RuntimeError as e:
        yield f"inner_caught: {e}"

def outer():
    yield from inner()

g = outer()
print(next(g))
print(g.throw(RuntimeError, "oops"))
"#;
    assert_eq!(run_python(script), vec!["1", "inner_caught: oops"]);
}

#[test]
fn test_yield_from_nested() {
    let script = r#"
def a():
    yield 1

def b():
    yield from a()
    yield 2

def c():
    yield from b()
    yield 3

print(list(c()))
"#;
    assert_eq!(run_python(script), vec!["[1, 2, 3]"]);
}
