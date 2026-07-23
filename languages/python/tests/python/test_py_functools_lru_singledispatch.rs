use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Functools LRU & SingleDispatch — lru_cache, cache, singledispatch, partial, reduce, wraps, total_ordering
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_functools_lru_cache_clear_and_info() {
    let src = r#"
from functools import lru_cache

call_count = 0

@lru_cache(maxsize=2)
def compute(x):
    global call_count
    call_count += 1
    return x * 2

print(compute(1))
print(compute(1))
print(compute(2))
print(compute(3))  # evicts 1
print(compute(1))  # recomputes 1

info = compute.cache_info()
print(info.hits)
print(info.misses)

compute.cache_clear()
info_cleared = compute.cache_info()
print(info_cleared.currsize)
"#;
    assert_eq!(
        run_python(src),
        vec!["2", "2", "4", "6", "2", "1", "4", "0"]
    );
}

#[test]
fn test_py_functools_cache_unbounded_decorator() {
    let src = r#"
from functools import cache

calls = 0

@cache
def factorial(n):
    global calls
    calls += 1
    if n == 0: return 1
    return n * factorial(n - 1)

print(factorial(5))
print(calls)
print(factorial(5))  # cached
print(calls)
"#;
    assert_eq!(run_python(src), vec!["120", "6", "120", "6"]);
}

#[test]
fn test_py_functools_singledispatch_generic_function() {
    let src = r#"
from functools import singledispatch

@singledispatch
def format_val(val):
    return f"str:{val}"

@format_val.register(int)
def _(val):
    return f"int:{val}"

@format_val.register(list)
def _(val):
    return f"list:{len(val)}"

print(format_val(10))
print(format_val([1, 2, 3]))
print(format_val(3.14))
"#;
    assert_eq!(run_python(src), vec!["int:10", "list:3", "str:3.14"]);
}

#[test]
fn test_py_functools_partial_keyword_and_positional() {
    let src = r#"
from functools import partial

def power(base, exponent):
    return base ** exponent

square = partial(power, exponent=2)
cube = partial(power, exponent=3)

print(square(4))
print(cube(3))
"#;
    assert_eq!(run_python(src), vec!["16", "27"]);
}

#[test]
fn test_py_functools_reduce_cumulative_aggregation() {
    let src = r#"
from functools import reduce

data = [1, 2, 3, 4, 5]
sum_all = reduce(lambda acc, x: acc + x, data)
max_all = reduce(lambda acc, x: acc if acc > x else x, data)

print(sum_all)
print(max_all)
"#;
    assert_eq!(run_python(src), vec!["15", "5"]);
}

#[test]
fn test_py_functools_total_ordering_class_decorator() {
    let src = r#"
from functools import total_ordering

@total_ordering
class Card:
    def __init__(self, rank):
        self.rank = rank

    def __eq__(self, other):
        return self.rank == other.rank

    def __lt__(self, other):
        return self.rank < other.rank

c1 = Card(5)
c2 = Card(10)

print(c1 < c2)
print(c1 <= c2)
print(c1 > c2)
print(c1 >= c2)
"#;
    assert_eq!(run_python(src), vec!["True", "True", "False", "False"]);
}

#[test]
fn test_py_functools_cached_property_descriptor() {
    let src = r#"
from functools import cached_property

class Circle:
    def __init__(self, radius):
        self.radius = radius

    @cached_property
    def area(self):
        print("Calculating area")
        return 3.14159 * (self.radius ** 2)

c = Circle(5)
print(round(c.area, 2))
print(round(c.area, 2))  # cached
"#;
    assert_eq!(run_python(src), vec!["Calculating area", "78.54", "78.54"]);
}

#[test]
fn test_py_functools_cmp_to_key_conversion() {
    let src = r#"
from functools import cmp_to_key

def custom_compare(a, b):
    # Sort strings by length, then alphabetically
    if len(a) != len(b):
        return len(a) - len(b)
    return (a > b) - (a < b)

words = ["banana", "apple", "fig", "date", "cherry"]
print(sorted(words, key=cmp_to_key(custom_compare)))
"#;
    assert_eq!(
        run_python(src),
        vec!["['fig', 'date', 'apple', 'banana', 'cherry']"]
    );
}

#[test]
fn test_py_functools_partialmethod_class_bound() {
    let src = r#"
from functools import partialmethod

class Logger:
    def log(self, level, message):
        return f"[{level}] {message}"

    debug = partialmethod(log, "DEBUG")
    error = partialmethod(log, "ERROR")

l = Logger()
print(l.debug("Starting"))
print(l.error("Failed"))
"#;
    assert_eq!(run_python(src), vec!["[DEBUG] Starting", "[ERROR] Failed"]);
}

#[test]
fn test_py_functools_wraps_update_wrapper() {
    let src = r#"
from functools import wraps

def my_dec(f):
    @wraps(f)
    def wrapper(*args): return f(*args)
    return wrapper

@my_dec
def target(x: int) -> int:
    """Target function docstring"""
    return x

print(target.__name__)
print(target.__doc__)
"#;
    assert_eq!(run_python(src), vec!["target", "Target function docstring"]);
}
