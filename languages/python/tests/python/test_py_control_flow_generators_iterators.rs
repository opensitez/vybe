use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Control Flow, Generators & Iterators — custom iterator protocol, iter() 2-arg sentinel, StopIteration, itertools.islice
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_custom_iterator_class_protocol() {
    let src = r#"
class Fibonacci:
    def __init__(self, limit):
        self.limit = limit
        self.a, self.b = 0, 1
        self.count = 0

    def __iter__(self):
        return self

    def __next__(self):
        if self.count >= self.limit:
            raise StopIteration
        val = self.a
        self.a, self.b = self.b, self.a + self.b
        self.count += 1
        return val

fibs = list(Fibonacci(6))
print(fibs)
"#;
    assert_eq!(run_python(src), vec!["[0, 1, 1, 2, 3, 5]"]);
}

#[test]
fn test_py_iter_callable_sentinel_pattern() {
    let src = r#"
count = 0
def get_next():
    global count
    count += 1
    return count

# iter calls get_next until it returns 4
it = iter(get_next, 4)
print(list(it))
"#;
    assert_eq!(run_python(src), vec!["[1, 2, 3]"]);
}

#[test]
fn test_py_stop_iteration_value_payload() {
    let src = r#"
def gen():
    yield "a"
    yield "b"
    return "done_value"

g = gen()
print(next(g))
print(next(g))
try:
    next(g)
except StopIteration as e:
    print(e.value)
"#;
    assert_eq!(run_python(src), vec!["a", "b", "done_value"]);
}

#[test]
fn test_py_generator_expression_chaining() {
    let src = r#"
nums = range(10)
evens = (x for x in nums if x % 2 == 0)
squared = (x * x for x in evens)
print(list(squared))
"#;
    assert_eq!(run_python(src), vec!["[0, 4, 16, 36, 64]"]);
}

#[test]
fn test_py_itertools_islice_range_slicing() {
    let src = r#"
from itertools import islice

def infinite_counter():
    n = 0
    while True:
        yield n
        n += 1

slice_out = list(islice(infinite_counter(), 5, 10))
print(slice_out)
"#;
    assert_eq!(run_python(src), vec!["[5, 6, 7, 8, 9]"]);
}

#[test]
fn test_py_generator_send_return_flow() {
    let src = r#"
def accumulator():
    total = 0
    while True:
        val = yield total
        if val is None:
            break
        total += val

acc = accumulator()
print(next(acc))  # prime
print(acc.send(10))
print(acc.send(20))
"#;
    assert_eq!(run_python(src), vec!["0", "10", "30"]);
}

#[test]
fn test_py_generator_throw_exception_recovery() {
    let src = r#"
def resilient():
    try:
        yield "working"
    except ValueError as e:
        yield f"recovered: {e}"

g = resilient()
print(next(g))
print(g.throw(ValueError("bad input")))
"#;
    assert_eq!(run_python(src), vec!["working", "recovered: bad input"]);
}

#[test]
fn test_py_generator_close_finally_block() {
    let src = r#"
events = []

def cleanup_gen():
    try:
        yield "step1"
        yield "step2"
    finally:
        events.append("cleaned up")

g = cleanup_gen()
print(next(g))
g.close()
print(events)
"#;
    assert_eq!(run_python(src), vec!["step1", "['cleaned up']"]);
}

#[test]
fn test_py_yield_from_subgenerator_delegation() {
    let src = r#"
def sub():
    yield 10
    yield 20
    return "sub_done"

def parent():
    res = yield from sub()
    yield f"parent_got:{res}"

print(list(parent()))
"#;
    assert_eq!(run_python(src), vec!["[10, 20, 'parent_got:sub_done']"]);
}

#[test]
fn test_py_iter_returns_self_for_iterators() {
    let src = r#"
g = (x for x in range(3))
print(iter(g) is g)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_next_default_fallback_argument() {
    let src = r#"
g = (x for x in [1, 2])
print(next(g, "end"))
print(next(g, "end"))
print(next(g, "end"))
"#;
    assert_eq!(run_python(src), vec!["1", "2", "end"]);
}

#[test]
fn test_py_generator_closure_scope_state() {
    let src = r#"
def make_gen(factor):
    return (x * factor for x in range(3))

g = make_gen(10)
print(list(g))
"#;
    assert_eq!(run_python(src), vec!["[0, 10, 20]"]);
}

#[test]
fn test_py_itertools_chain_generator_sequences() {
    let src = r#"
from itertools import chain

g1 = (x for x in range(2))
g2 = (x * 10 for x in range(1, 3))
chained = list(chain(g1, g2))
print(chained)
"#;
    assert_eq!(run_python(src), vec!["[0, 1, 10, 20]"]);
}

#[test]
fn test_py_custom_iterable_multiple_passes() {
    let src = r#"
class MultiPass:
    def __init__(self, data):
        self.data = data

    def __iter__(self):
        return iter(self.data)

mp = MultiPass([1, 2, 3])
print(list(mp))
print(list(mp))  # reusable!
"#;
    assert_eq!(run_python(src), vec!["[1, 2, 3]", "[1, 2, 3]"]);
}

#[test]
fn test_py_generator_recursion_tree_walk() {
    let src = r#"
tree = [1, [2, [3, 4]], 5]

def flatten(nested):
    for item in nested:
        if isinstance(item, list):
            yield from flatten(item)
        else:
            yield item

print(list(flatten(tree)))
"#;
    assert_eq!(run_python(src), vec!["[1, 2, 3, 4, 5]"]);
}
