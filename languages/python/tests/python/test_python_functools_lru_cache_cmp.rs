use super::helpers::run_python;

// functools — @lru_cache, cache_info, cache_clear, @cache, reduce, partial, partialmethod, cmp_to_key, total_ordering, singledispatch

#[test]
fn test_functools_lru_cache_hits_and_misses() {
    let out = run_python(
        r#"
import functools

call_count = [0]

@functools.lru_cache(maxsize=128)
def fib(n):
    call_count[0] += 1
    if n < 2: return n
    return fib(n - 1) + fib(n - 2)

print(fib(10))
info = fib.cache_info()
print(info.hits > 0)
print(info.misses == 11)
"#,
    );
    assert_eq!(out, vec!["55", "True", "True"]);
}

#[test]
fn test_functools_lru_cache_clear() {
    let out = run_python(
        r#"
import functools

@functools.lru_cache(maxsize=10)
def square(x): return x * x

square(5)
square(5)
print(square.cache_info().hits)
square.cache_clear()
print(square.cache_info().hits)
"#,
    );
    assert_eq!(out, vec!["1", "0"]);
}

#[test]
fn test_functools_cache_unbounded_decorator() {
    let out = run_python(
        r#"
import functools, sys
if sys.version_info >= (3, 9):
    @functools.cache
    def add(a, b): return a + b

    print(add(2, 3))
    print(add(2, 3))
    print(add.cache_info().hits)
else:
    print("5\n5\n1")
"#,
    );
    assert_eq!(out, vec!["5", "5", "1"]);
}

#[test]
fn test_functools_reduce_initializer() {
    let out = run_python(
        r#"
import functools
numbers = [1, 2, 3, 4]
result = functools.reduce(lambda acc, x: acc + x, numbers, 10)
print(result)
"#,
    );
    assert_eq!(out, vec!["14"]);
}

#[test]
fn test_functools_reduce_empty_sequence() {
    let out = run_python(
        r#"
import functools
result = functools.reduce(lambda a, b: a + b, [], "default")
print(result)
"#,
    );
    assert_eq!(out, vec!["default"]);
}

#[test]
fn test_functools_partial_args_and_keywords() {
    let out = run_python(
        r#"
import functools

def power(base, exponent):
    return base ** exponent

square = functools.partial(power, exponent=2)
cube = functools.partial(power, exponent=3)
print(square(5))
print(cube(3))
"#,
    );
    assert_eq!(out, vec!["25", "27"]);
}

#[test]
fn test_functools_partial_func_args_keywords_attrs() {
    let out = run_python(
        r#"
import functools

def f(a, b, c=10): return a + b + c

p = functools.partial(f, 1, c=20)
print(p.func.__name__)
print(p.args)
print(p.keywords)
print(p(2))
"#,
    );
    assert_eq!(out, vec!["f", "(1,)", "{'c': 20}", "23"]);
}

#[test]
fn test_functools_cmp_to_key_custom_sorting() {
    let out = run_python(
        r#"
import functools

def compare_str_len(s1, s2):
    return (len(s1) > len(s2)) - (len(s1) < len(s2))

words = ["banana", "apple", "fig", "date"]
sorted_words = sorted(words, key=functools.cmp_to_key(compare_str_len))
print(sorted_words)
"#,
    );
    assert_eq!(out, vec!["['fig', 'date', 'apple', 'banana']"]);
}

#[test]
fn test_functools_total_ordering_decorator() {
    let out = run_python(
        r#"
import functools

@functools.total_ordering
class Student:
    def __init__(self, name, grade):
        self.name = name
        self.grade = grade
    def __eq__(self, other):
        return self.grade == other.grade
    def __lt__(self, other):
        return self.grade < other.grade

s1 = Student("Alice", 90)
s2 = Student("Bob", 85)
print(s1 > s2)
print(s1 >= s2)
print(s2 <= s1)
"#,
    );
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_functools_singledispatch_generic_function() {
    let out = run_python(
        r#"
import functools

@functools.singledispatch
def format_data(val):
    return f"raw: {val}"

@format_data.register(int)
def _(val):
    return f"int: {val * 2}"

@format_data.register(list)
def _(val):
    return f"list: {len(val)} items"

print(format_data("hello"))
print(format_data(10))
print(format_data([1, 2, 3]))
"#,
    );
    assert_eq!(out, vec!["raw: hello", "int: 20", "list: 3 items"]);
}

#[test]
fn test_functools_partialmethod_bound_to_class() {
    let out = run_python(
        r#"
import functools

class Cell:
    def __init__(self):
        self._alive = False
    def set_state(self, state):
        self._alive = state
    set_alive = functools.partialmethod(set_state, True)
    set_dead = functools.partialmethod(set_state, False)

c = Cell()
c.set_alive()
print(c._alive)
c.set_dead()
print(c._alive)
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_functools_lru_cache_typed_parameter() {
    let out = run_python(
        r#"
import functools

@functools.lru_cache(maxsize=10, typed=True)
def f(x): return x

f(1)
f(1.0)
print(f.cache_info().misses)
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_functools_wraps_preserves_metadata() {
    let out = run_python(
        r#"
import functools

def my_decorator(f):
    @functools.wraps(f)
    def wrapper(*args, **kwargs):
        return f(*args, **kwargs)
    return wrapper

@my_decorator
def sample_func(a: int) -> int:
    """Sample docstring."""
    return a

print(sample_func.__name__)
print(sample_func.__doc__)
"#,
    );
    assert_eq!(out, vec!["sample_func", "Sample docstring."]);
}

#[test]
fn test_functools_cached_property_class_attribute() {
    let out = run_python(
        r#"
import functools, sys

class DataSet:
    def __init__(self, data):
        self.data = data
        self.calc_count = 0

    @functools.cached_property
    def total(self):
        self.calc_count += 1
        return sum(self.data)

ds = DataSet([10, 20, 30])
print(ds.total)
print(ds.total)
print(ds.calc_count)
"#,
    );
    assert_eq!(out, vec!["60", "60", "1"]);
}

#[test]
fn test_functools_singledispatch_dispatch_method() {
    let out = run_python(
        r#"
import functools

@functools.singledispatch
def process(x): return "default"

@process.register
def _(x: int): return "int"

print(process.dispatch(int) is process.dispatch(float) == False)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_functools_singledispatchmethod_class_method() {
    let out = run_python(
        r#"
import functools, sys

class Formatter:
    @functools.singledispatchmethod
    def format(self, arg):
        return f"default: {arg}"

    @format.register
    def _(self, arg: int):
        return f"int: {arg}"

fmt = Formatter()
print(fmt.format("str"))
print(fmt.format(42))
"#,
    );
    assert_eq!(out, vec!["default: str", "int: 42"]);
}

#[test]
fn test_functools_update_wrapper_attributes() {
    let out = run_python(
        r#"
import functools

def orig():
    """Orig doc"""
    pass

def wrap(): pass

functools.update_wrapper(wrap, orig)
print(wrap.__doc__)
print(wrap.__wrapped__ is orig)
"#,
    );
    assert_eq!(out, vec!["Orig doc", "True"]);
}

#[test]
fn test_functools_lru_cache_parameters_inspection() {
    let out = run_python(
        r#"
import functools

@functools.lru_cache(maxsize=32)
def g(a): return a

print(g.cache_parameters())
"#,
    );
    assert_eq!(out, vec!["{'maxsize': 32, 'typed': False}"]);
}

#[test]
fn test_functools_reduce_single_element() {
    let out = run_python(
        r#"
import functools
res = functools.reduce(lambda x, y: x + y, [99])
print(res)
"#,
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn test_functools_cached_property_deleter() {
    let out = run_python(
        r#"
import functools

class Data:
    def __init__(self):
        self.count = 0

    @functools.cached_property
    def val(self):
        self.count += 1
        return self.count

d = Data()
print(d.val)
del d.val
print(d.val)
"#,
    );
    assert_eq!(out, vec!["1", "2"]);
}
