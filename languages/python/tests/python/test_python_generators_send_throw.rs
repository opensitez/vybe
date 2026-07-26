// Python generators — .send(), .throw(), .close(), GeneratorExit
use super::helpers::run_python;

#[test]
fn test_generator_send_value() {
    let script = r#"
def accumulator():
    total = 0
    while True:
        value = yield total
        if value is None:
            break
        total += value

g = accumulator()
next(g)  # prime
print(g.send(10))
print(g.send(20))
print(g.send(5))
"#;
    assert_eq!(run_python(script), vec!["10", "30", "35"]);
}

#[test]
fn test_generator_throw() {
    let script = r#"
def gen():
    try:
        yield 1
        yield 2
    except ValueError as e:
        yield f"caught: {e}"

g = gen()
print(next(g))
print(g.throw(ValueError, "oops"))
"#;
    assert_eq!(run_python(script), vec!["1", "caught: oops"]);
}

#[test]
fn test_generator_close() {
    let script = r#"
def gen():
    try:
        yield 1
        yield 2
    finally:
        print("cleanup")

g = gen()
print(next(g))
g.close()
"#;
    assert_eq!(run_python(script), vec!["1", "cleanup"]);
}

#[test]
fn test_generator_return_value() {
    let script = r#"
def gen():
    yield 1
    yield 2
    return "done"

g = gen()
print(next(g))
print(next(g))
try:
    next(g)
except StopIteration as e:
    print(e.value)
"#;
    assert_eq!(run_python(script), vec!["1", "2", "done"]);
}

#[test]
fn test_yield_from_delegation() {
    let script = r#"
def inner():
    yield 1
    yield 2

def outer():
    yield 0
    yield from inner()
    yield 3

print(list(outer()))
"#;
    assert_eq!(run_python(script), vec!["[0, 1, 2, 3]"]);
}

#[test]
fn test_generator_send_priming() {
    let script = r#"
def echo():
    while True:
        received = yield
        print(f"got: {received}")

g = echo()
next(g)  # prime
g.send("hello")
g.send("world")
g.close()
"#;
    assert_eq!(run_python(script), vec!["got: hello", "got: world"]);
}

#[test]
fn test_generator_pipeline() {
    let script = r#"
def producer():
    for i in range(5):
        yield i

def doubler(gen):
    for x in gen:
        yield x * 2

def consumer(gen):
    return list(gen)

result = consumer(doubler(producer()))
print(result)
"#;
    assert_eq!(run_python(script), vec!["[0, 2, 4, 6, 8]"]);
}
