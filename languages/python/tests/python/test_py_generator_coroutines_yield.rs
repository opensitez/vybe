use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Generator Coroutines & Yield — yield from, send, throw, close, StopIteration, subgenerators
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_generator_send_bidirectional_data() {
    let src = r#"
def echo_coroutine():
    val = yield "ready"
    while True:
        val = yield f"echo: {val}"

co = echo_coroutine()
print(next(co))
print(co.send("first"))
print(co.send("second"))
"#;
    assert_eq!(
        run_python(src),
        vec!["ready", "echo: first", "echo: second"]
    );
}

#[test]
fn test_py_yield_from_subgenerator_return_value() {
    let src = r#"
def subgenerator():
    yield 1
    yield 2
    return "subgen_result"

def main_generator():
    res = yield from subgenerator()
    yield f"captured: {res}"

g = main_generator()
print(list(g))
"#;
    assert_eq!(run_python(src), vec!["[1, 2, 'captured: subgen_result']"]);
}

#[test]
fn test_py_generator_throw_exception_into_generator() {
    let src = r#"
def resilient_generator():
    try:
        yield "start"
        yield "running"
    except ValueError as e:
        yield f"recovered from: {e}"

g = resilient_generator()
print(next(g))
print(g.throw(ValueError("bad input")))
"#;
    assert_eq!(run_python(src), vec!["start", "recovered from: bad input"]);
}

#[test]
fn test_py_generator_close_cleanup_finally() {
    let src = r#"
events = []

def cleanup_gen():
    try:
        yield 1
        yield 2
    finally:
        events.append("cleanup executed")

g = cleanup_gen()
print(next(g))
g.close()
print(events)
"#;
    assert_eq!(run_python(src), vec!["1", "['cleanup executed']"]);
}

#[test]
fn test_py_generator_expression_statefulness() {
    let src = r#"
squares = (x * x for x in range(5))
print(next(squares))
print(next(squares))
print(list(squares))  # consumes remainder
print(list(squares))  # empty now
"#;
    assert_eq!(run_python(src), vec!["0", "1", "[4, 9, 16]", "[]"]);
}

#[test]
fn test_py_stopiteration_value_extraction() {
    let src = r#"
def fn():
    yield 10
    return "final_value"

g = fn()
next(g)
try:
    next(g)
except StopIteration as e:
    print(e.value)
"#;
    assert_eq!(run_python(src), vec!["final_value"]);
}

#[test]
fn test_py_generator_pipeline_composition() {
    let src = r#"
def numbers(n):
    for i in range(n):
        yield i

def evens(seq):
    for x in seq:
        if x % 2 == 0:
            yield x

def doubled(seq):
    for x in seq:
        yield x * 2

pipeline = doubled(evens(numbers(10)))
print(list(pipeline))
"#;
    assert_eq!(run_python(src), vec!["[0, 4, 8, 12, 16]"]);
}

#[test]
fn test_py_yield_from_iterable_delegation() {
    let src = r#"
def delegate_all():
    yield from [1, 2, 3]
    yield from (x * 10 for x in range(1, 4))
    yield from "AB"

print(list(delegate_all()))
"#;
    assert_eq!(run_python(src), vec!["[1, 2, 3, 10, 20, 30, 'A', 'B']"]);
}

#[test]
fn test_py_generator_isinstance_inspect() {
    let src = r#"
import inspect

def gen_func():
    yield 1

g = gen_func()
print(inspect.isgeneratorfunction(gen_func))
print(inspect.isgenerator(g))
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_generator_accumulator_coroutine() {
    let src = r#"
def running_total():
    total = 0
    while True:
        val = yield total
        if val is None:
            break
        total += val

t = running_total()
next(t)  # prime
print(t.send(10))
print(t.send(20))
print(t.send(30))
"#;
    assert_eq!(run_python(src), vec!["10", "30", "60"]);
}
