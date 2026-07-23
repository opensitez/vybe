use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: functools — lru_cache, cache, partial, reduce, singledispatch, cached_property, cmp_to_key, wraps
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_functools_lru_cache_basic() {
    let src = r#"
import functools

call_count = [0]

@functools.lru_cache(maxsize=128)
def fib(n):
    call_count[0] += 1
    if n < 2:
        return n
    return fib(n - 1) + fib(n - 2)

print(fib(10))
print(fib(10))  # from cache
info = fib.cache_info()
print(info.hits > 0)
print(info.misses == 11)  # 0..10
"#;
    assert_eq!(run_python(src), vec!["55", "55", "True", "True"]);
}

#[test]
fn test_py_functools_cache_unbounded() {
    let src = r#"
import functools

@functools.cache
def square(n):
    return n * n

print(square(4))
print(square(4))
print(square.cache_info().hits)
"#;
    assert_eq!(run_python(src), vec!["16", "16", "1"]);
}

#[test]
fn test_py_functools_lru_cache_maxsize_eviction() {
    let src = r#"
import functools

@functools.lru_cache(maxsize=2)
def identity(x):
    return x

identity(1)
identity(2)
identity(3)  # evicts 1
identity(4)  # evicts 2
info = identity.cache_info()
print(info.currsize)  # should be 2
"#;
    assert_eq!(run_python(src), vec!["2"]);
}

#[test]
fn test_py_functools_partial_positional_binding() {
    let src = r#"
import functools

def power(base, exp):
    return base ** exp

square = functools.partial(power, exp=2)
cube = functools.partial(power, exp=3)
print(square(4))
print(cube(3))
"#;
    assert_eq!(run_python(src), vec!["16", "27"]);
}

#[test]
fn test_py_functools_partial_keyword_binding() {
    let src = r#"
import functools

def greet(greeting, name, punctuation="!"):
    return f"{greeting}, {name}{punctuation}"

hello = functools.partial(greet, "Hello", punctuation=".")
print(hello("Alice"))
print(hello("Bob"))
"#;
    assert_eq!(run_python(src), vec!["Hello, Alice.", "Hello, Bob."]);
}

#[test]
fn test_py_functools_reduce_sum() {
    let src = r#"
import functools

total = functools.reduce(lambda acc, x: acc + x, [1, 2, 3, 4, 5])
print(total)

product = functools.reduce(lambda acc, x: acc * x, [1, 2, 3, 4], 1)
print(product)
"#;
    assert_eq!(run_python(src), vec!["15", "24"]);
}

#[test]
fn test_py_functools_reduce_with_initial_value() {
    let src = r#"
import functools

result = functools.reduce(lambda acc, x: acc + x, [], 42)
print(result)  # empty sequence with initializer returns initializer

nested = functools.reduce(lambda d, k: d[k], ["a", "b", "c"], {"a": {"b": {"c": 99}}})
print(nested)
"#;
    assert_eq!(run_python(src), vec!["42", "99"]);
}

#[test]
fn test_py_functools_singledispatch() {
    let src = r#"
import functools

@functools.singledispatch
def process(arg):
    return f"generic: {arg}"

@process.register(int)
def _(arg):
    return f"int: {arg * 2}"

@process.register(str)
def _(arg):
    return f"str: {arg.upper()}"

@process.register(list)
def _(arg):
    return f"list length: {len(arg)}"

print(process(5))
print(process("hello"))
print(process([1, 2, 3]))
print(process(3.14))
"#;
    assert_eq!(
        run_python(src),
        vec!["int: 10", "str: HELLO", "list length: 3", "generic: 3.14"]
    );
}

#[test]
fn test_py_functools_cached_property() {
    let src = r#"
import functools

calls = [0]

class DataSet:
    def __init__(self, data):
        self.data = data

    @functools.cached_property
    def stats(self):
        calls[0] += 1
        return {"sum": sum(self.data), "count": len(self.data)}

ds = DataSet([1, 2, 3, 4, 5])
print(ds.stats)
print(ds.stats)  # cached — no recompute
print(calls[0])
"#;
    assert_eq!(
        run_python(src),
        vec!["{'sum': 15, 'count': 5}", "{'sum': 15, 'count': 5}", "1"]
    );
}

#[test]
fn test_py_functools_cmp_to_key() {
    let src = r#"
import functools

def compare_lengths(a, b):
    if len(a) < len(b): return -1
    if len(a) > len(b): return 1
    return 0

words = ["banana", "kiwi", "apple", "fig", "cherry"]
print(sorted(words, key=functools.cmp_to_key(compare_lengths)))
"#;
    assert_eq!(
        run_python(src),
        vec!["['fig', 'kiwi', 'apple', 'banana', 'cherry']"]
    );
}

#[test]
fn test_py_functools_wraps_preserves_metadata() {
    let src = r#"
import functools

def decorator(func):
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        return func(*args, **kwargs)
    return wrapper

@decorator
def compute(x: int) -> int:
    """Computes x squared."""
    return x ** 2

print(compute.__name__)
print(compute.__doc__)
print(compute.__annotations__)
"#;
    assert_eq!(
        run_python(src),
        vec![
            "compute",
            "Computes x squared.",
            "{'x': <class 'int'>, 'return': <class 'int'>}"
        ]
    );
}

#[test]
fn test_py_functools_total_ordering() {
    let src = r#"
import functools

@functools.total_ordering
class Version:
    def __init__(self, major, minor):
        self.major = major
        self.minor = minor

    def __eq__(self, other):
        return (self.major, self.minor) == (other.major, other.minor)

    def __lt__(self, other):
        return (self.major, self.minor) < (other.major, other.minor)

v1 = Version(1, 0)
v2 = Version(2, 0)
print(v1 < v2)
print(v1 > v2)
print(v1 <= v1)
print(v1 >= v2)
"#;
    assert_eq!(run_python(src), vec!["True", "False", "True", "False"]);
}

#[test]
fn test_py_functools_partial_method_in_class() {
    let src = r#"
import functools

class Formatter:
    def format_value(self, prefix, value):
        return f"{prefix}: {value}"

    format_price = functools.partialmethod(format_value, "Price")
    format_qty = functools.partialmethod(format_value, "Qty")

f = Formatter()
print(f.format_price(99.99))
print(f.format_qty(42))
"#;
    assert_eq!(run_python(src), vec!["Price: 99.99", "Qty: 42"]);
}

#[test]
fn test_py_functools_reduce_flatten_nested() {
    let src = r#"
import functools, operator

nested = [[1, 2], [3, 4], [5, 6]]
flat = functools.reduce(operator.add, nested)
print(flat)
"#;
    assert_eq!(run_python(src), vec!["[1, 2, 3, 4, 5, 6]"]);
}

#[test]
fn test_py_functools_singledispatch_abc_registration() {
    let src = r#"
import functools
from numbers import Number

@functools.singledispatch
def describe(x):
    return f"unknown: {type(x).__name__}"

@describe.register(Number)
def _(x):
    return f"number: {x}"

print(describe(42))
print(describe(3.14))
print(describe("hello"))
"#;
    assert_eq!(
        run_python(src),
        vec!["number: 42", "number: 3.14", "unknown: str"]
    );
}
