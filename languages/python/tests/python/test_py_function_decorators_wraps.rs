use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Function Decorators & Wraps — function decorators, class decorators, parametric decorators, wraps
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_basic_function_decorator() {
    let src = r#"
def logger(func):
    def wrapper(*args, **kwargs):
        print(f"Calling {func.__name__}")
        return func(*args, **kwargs)
    return wrapper

@logger
fn greet(name):
    return f"Hello {name}"

print(greet("Alice"))
"#;
    assert_eq!(run_python(src), vec!["Calling greet", "Hello Alice"]);
}

#[test]
fn test_py_functools_wraps_preserves_metadata() {
    let src = r#"
from functools import wraps

def my_decorator(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        """Wrapper doc"""
        return func(*args, **kwargs)
    return wrapper

@my_decorator
def add(a: int, b: int) -> int:
    """Add two numbers"""
    return a + b

print(add.__name__)
print(add.__doc__)
print(add.__annotations__)
"#;
    assert_eq!(
        run_python(src),
        vec![
            "add",
            "Add two numbers",
            "{'a': <class 'int'>, 'b': <class 'int'>, 'return': <class 'int'>}"
        ]
    );
}

#[test]
fn test_py_decorator_with_arguments_factory() {
    let src = r#"
from functools import wraps

def repeat(num_times):
    def decorator_repeat(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            for _ in range(num_times):
                result = func(*args, **kwargs)
            return result
        return wrapper
    return decorator_repeat

log = []

@repeat(num_times=3)
def ping():
    log.append("pong")

ping()
print(log)
"#;
    assert_eq!(run_python(src), vec!["['pong', 'pong', 'pong']"]);
}

#[test]
fn test_py_class_decorator_modifying_class() {
    let src = r#"
def add_str_repr(cls):
    cls.__str__ = lambda self: f"{cls.__name__}({self.__dict__})"
    return cls

@add_str_repr
class Product:
    def __init__(self, name, price):
        self.name = name
        self.price = price

p = Product("book", 15)
print(str(p))
"#;
    assert_eq!(
        run_python(src),
        vec!["Product({'name': 'book', 'price': 15})"]
    );
}

#[test]
fn test_py_decorator_chaining_order() {
    let src = r#"
def dec1(func):
    def wrapper():
        return "dec1(" + func() + ")"
    return wrapper

def dec2(func):
    def wrapper():
        return "dec2(" + func() + ")"
    return wrapper

# Applied bottom-up: dec1(dec2(base))
@dec1
@dec2
def base():
    return "base"

print(base())
"#;
    assert_eq!(run_python(src), vec!["dec1(dec2(base))"]);
}

#[test]
fn test_py_class_as_decorator_stateful() {
    let src = r#"
class CountCalls:
    def __init__(self, func):
        self.func = func
        self.num_calls = 0

    def __call__(self, *args, **kwargs):
        self.num_calls += 1
        return self.func(*args, **kwargs)

@CountCalls
def say_hello():
    return "hello"

say_hello()
say_hello()
say_hello()
print(say_hello.num_calls)
"#;
    assert_eq!(run_python(src), vec!["3"]);
}

#[test]
fn test_py_method_decorator_binding() {
    let src = r#"
from functools import wraps

def trace_method(func):
    @wraps(func)
    def wrapper(self, *args, **kwargs):
        print(f"Method {func.__name__} called on {self.name}")
        return func(self, *args, **kwargs)
    return wrapper

class Robot:
    def __init__(self, name):
        self.name = name

    @trace_method
    def work(self):
        return "working"

r = Robot("R2D2")
print(r.work())
"#;
    assert_eq!(
        run_python(src),
        vec!["Method work called on R2D2", "working"]
    );
}

#[test]
fn test_py_decorator_optional_arguments_pattern() {
    let src = r#"
from functools import wraps, partial

def smart_decorator(func=None, *, prefix="LOG"):
    if func is None:
        return partial(smart_decorator, prefix=prefix)

    @wraps(func)
    def wrapper(*args, **kwargs):
        print(f"[{prefix}] Calling {func.__name__}")
        return func(*args, **kwargs)
    return wrapper

@smart_decorator
def f1(): return "f1"

@smart_decorator(prefix="CUSTOM")
def f2(): return "f2"

f1()
f2()
"#;
    assert_eq!(
        run_python(src),
        vec!["[LOG] Calling f1", "[CUSTOM] Calling f2"]
    );
}

#[test]
fn test_py_functools_lru_cache_decorator() {
    let src = r#"
from functools import lru_cache

call_count = 0

@lru_cache(maxsize=32)
def fib(n):
    global call_count
    call_count += 1
    if n < 2: return n
    return fib(n - 1) + fib(n - 2)

print(fib(10))
print(call_count)  # memoized, only 11 calls
"#;
    assert_eq!(run_python(src), vec!["55", "11"]);
}

#[test]
fn test_py_functools_singledispatch_decorator() {
    let src = r#"
from functools import singledispatch

@singledispatch
def process(val):
    return f"default: {val}"

@process.register(int)
def _(val):
    return f"int: {val * 2}"

@process.register(str)
def _(val):
    return f"str: {val.upper()}"

print(process(10))
print(process("hello"))
print(process(3.14))
"#;
    assert_eq!(
        run_python(src),
        vec!["int: 20", "str: HELLO", "default: 3.14"]
    );
}
