use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: generators + iterators — yield, yield from, send, throw, StopIteration, generator expressions, __iter__, __next__
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_generator_basic_yield() {
    let src = r#"
def count_up(n):
    for i in range(n):
        yield i

gen = count_up(4)
print(next(gen))
print(next(gen))
print(list(gen))  # exhaust remaining
"#;
    assert_eq!(run_python(src), vec!["0", "1", "[2, 3]"]);
}

#[test]
fn test_py_generator_infinite_sequence() {
    let src = r#"
def fibonacci():
    a, b = 0, 1
    while True:
        yield a
        a, b = b, a + b

import itertools
print(list(itertools.islice(fibonacci(), 8)))
"#;
    assert_eq!(run_python(src), vec!["[0, 1, 1, 2, 3, 5, 8, 13]"]);
}

#[test]
fn test_py_generator_send_value() {
    let src = r#"
def accumulator():
    total = 0
    while True:
        value = yield total
        if value is None:
            break
        total += value

gen = accumulator()
next(gen)  # prime the generator
print(gen.send(10))
print(gen.send(20))
print(gen.send(5))
"#;
    assert_eq!(run_python(src), vec!["10", "30", "35"]);
}

#[test]
fn test_py_generator_throw_exception() {
    let src = r#"
def guarded():
    try:
        yield "running"
        yield "still going"
    except ValueError:
        yield "caught ValueError"

gen = guarded()
print(next(gen))
print(gen.throw(ValueError, "oops"))
"#;
    assert_eq!(run_python(src), vec!["running", "caught ValueError"]);
}

#[test]
fn test_py_generator_return_value() {
    let src = r#"
def gen_with_return():
    yield 1
    yield 2
    return "done"

g = gen_with_return()
print(next(g))
print(next(g))
try:
    next(g)
except StopIteration as e:
    print(e.value)
"#;
    assert_eq!(run_python(src), vec!["1", "2", "done"]);
}

#[test]
fn test_py_generator_yield_from_delegates() {
    let src = r#"
def inner():
    yield 1
    yield 2
    return "inner_done"

def outer():
    result = yield from inner()
    print(f"inner returned: {result}")
    yield 3

print(list(outer()))
"#;
    assert_eq!(
        run_python(src),
        vec!["inner returned: inner_done", "[1, 2, 3]"]
    );
}

#[test]
fn test_py_generator_yield_from_flattens_nested() {
    let src = r#"
def flatten(nested):
    for item in nested:
        if isinstance(item, list):
            yield from flatten(item)
        else:
            yield item

data = [1, [2, [3, 4], 5], [6, 7]]
print(list(flatten(data)))
"#;
    assert_eq!(run_python(src), vec!["[1, 2, 3, 4, 5, 6, 7]"]);
}

#[test]
fn test_py_generator_expression() {
    let src = r#"
squares = (x ** 2 for x in range(5))
print(type(squares).__name__)
print(next(squares))
print(sum(squares))  # 1+4+9+16 = 30
"#;
    assert_eq!(run_python(src), vec!["generator", "0", "30"]);
}

#[test]
fn test_py_generator_close_cleanup() {
    let src = r#"
log = []

def gen():
    try:
        yield 1
        yield 2
    except GeneratorExit:
        log.append("cleaned_up")

g = gen()
next(g)
g.close()  # triggers GeneratorExit inside generator
print(log)
"#;
    assert_eq!(run_python(src), vec!["['cleaned_up']"]);
}

#[test]
fn test_py_iterator_protocol_custom_class() {
    let src = r#"
class Range:
    def __init__(self, start, stop):
        self.current = start
        self.stop = stop

    def __iter__(self):
        return self

    def __next__(self):
        if self.current >= self.stop:
            raise StopIteration
        val = self.current
        self.current += 1
        return val

print(list(Range(2, 6)))
"#;
    assert_eq!(run_python(src), vec!["[2, 3, 4, 5]"]);
}

#[test]
fn test_py_iterable_vs_iterator_distinction() {
    let src = r#"
class NumberList:
    def __init__(self, data):
        self.data = data

    def __iter__(self):
        return iter(self.data)

nl = NumberList([10, 20, 30])
for v in nl:
    print(v)
# Can iterate multiple times:
print(sum(nl))
"#;
    assert_eq!(run_python(src), vec!["10", "20", "30", "60"]);
}

#[test]
fn test_py_generator_pipeline() {
    let src = r#"
def integers():
    n = 1
    while True:
        yield n
        n += 1

def squares(it):
    for n in it:
        yield n * n

def take(n, it):
    for _ in range(n):
        yield next(it)

pipeline = take(5, squares(integers()))
print(list(pipeline))
"#;
    assert_eq!(run_python(src), vec!["[1, 4, 9, 16, 25]"]);
}

#[test]
fn test_py_generator_stateful_transformation() {
    let src = r#"
def running_average():
    total = 0
    count = 0
    while True:
        val = yield (total / count) if count else 0
        if val is not None:
            total += val
            count += 1

g = running_average()
next(g)
g.send(10)
g.send(20)
result = g.send(30)
print(result)
"#;
    assert_eq!(run_python(src), vec!["20.0"]);
}
