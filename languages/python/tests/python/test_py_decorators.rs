use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Function & Class Decorators, `@functools.wraps`, Stateful Closures
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_simple_function_decorator() {
    let src = r#"
def my_decorator(func):
    def wrapper(*args, **kwargs):
        print("before")
        res = func(*args, **kwargs)
        print("after")
        return res
    return wrapper

@my_decorator
def test():
    print("inside")

test()
"#;
    assert_eq!(run_python(src), vec!["before", "inside", "after"]);
}

#[test]
fn test_py_functools_wraps_preserves_doc_and_name() {
    let src = r#"
import functools

def log_calls(func):
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        return func(*args, **kwargs)
    return wrapper

@log_calls
def add(a, b):
    """Adds two numbers."""
    return a + b

print(add.__name__)
print(add.__doc__)
"#;
    assert_eq!(run_python(src), vec!["add", "Adds two numbers."]);
}

#[test]
fn test_py_decorator_with_arguments_factory() {
    let src = r#"
def repeat(num_times):
    def decorator(func):
        def wrapper(*args, **kwargs):
            for _ in range(num_times):
                func(*args, **kwargs)
        return wrapper
    return decorator

@repeat(num_times=3)
def greet():
    print("Hello!")

greet()
"#;
    assert_eq!(run_python(src), vec!["Hello!", "Hello!", "Hello!"]);
}

#[test]
fn test_py_stacked_decorators_execution_order() {
    let src = r#"
def dec_a(func):
    def wrapper():
        return "A(" + func() + ")"
    return wrapper

def dec_b(func):
    def wrapper():
        return "B(" + func() + ")"
    return wrapper

@dec_a
@dec_b
def base():
    return "Base"

print(base())
"#;
    assert_eq!(run_python(src), vec!["A(B(Base))"]); // Stacked decorators apply bottom-up!
}

#[test]
fn test_py_class_decorator_modifying_attributes() {
    let src = r#"
def add_metadata(cls):
    cls.version = "1.0.0"
    cls.get_version = lambda self: cls.version
    return cls

@add_metadata
class Service:
    pass

s = Service()
print(s.version)
print(s.get_version())
"#;
    assert_eq!(run_python(src), vec!["1.0.0", "1.0.0"]);
}

#[test]
fn test_py_stateful_callable_class_decorator() {
    let src = r#"
class CountCalls:
    def __init__(self, func):
        self.func = func
        self.num_calls = 0

    def __call__(self, *args, **kwargs):
        self.num_calls += 1
        return f"Call {self.num_calls}: {self.func(*args, **kwargs)}"

@CountCalls
def say_hi(name):
    return f"Hi {name}"

print(say_hi("Alice"))
print(say_hi("Bob"))
"#;
    assert_eq!(run_python(src), vec!["Call 1: Hi Alice", "Call 2: Hi Bob"]);
}

#[test]
fn test_py_decorator_preserving_wrapped_attribute() {
    let src = r#"
import functools

def my_dec(func):
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        return func(*args, **kwargs)
    return wrapper

@my_dec
def original():
    return "original_result"

print(original.__wrapped__())
"#;
    assert_eq!(run_python(src), vec!["original_result"]);
}

#[test]
fn test_py_method_decorator_in_class_body() {
    let src = r#"
def double_return(func):
    def wrapper(self, *args, **kwargs):
        return func(self, *args, **kwargs) * 2
    return wrapper

class Calculator:
    @double_return
    def compute(self, x):
        return x + 5

c = Calculator()
print(c.compute(10))
"#;
    assert_eq!(run_python(src), vec!["30"]); // (10 + 5) * 2 = 30
}

#[test]
fn test_py_property_decorator_getter_setter_deleter() {
    let src = r#"
class User:
    def __init__(self, name):
        self._name = name

    @property
    def name(self):
        return self._name.upper()

    @name.setter
    def name(self, value):
        self._name = value

    @name.deleter
    def name(self):
        self._name = "DELETED"

u = User("alice")
print(u.name)
u.name = "bob"
print(u.name)
del u.name
print(u._name)
"#;
    assert_eq!(run_python(src), vec!["ALICE", "BOB", "DELETED"]);
}

#[test]
fn test_py_classmethod_and_staticmethod_decorators() {
    let src = r#"
class Utility:
    @staticmethod
    def add(a, b):
        return a + b

    @classmethod
    def get_class_name(cls):
        return cls.__name__

print(Utility.add(3, 4))
print(Utility.get_class_name())
"#;
    assert_eq!(run_python(src), vec!["7", "Utility"]);
}

#[test]
fn test_py_decorator_optional_arguments_pattern() {
    let src = r#"
import functools

def smart_decorator(_func=None, *, prefix="[LOG]"):
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            return f"{prefix} {func(*args, **kwargs)}"
        return wrapper

    if _func is None:
        return decorator
    else:
        return decorator(_func)

@smart_decorator
def f1():
    return "f1"

@smart_decorator(prefix="[CUSTOM]")
def f2():
    return "f2"

print(f1())
print(f2())
"#;
    assert_eq!(run_python(src), vec!["[LOG] f1", "[CUSTOM] f2"]);
}

#[test]
fn test_py_closure_cell_mutation_nonlocal() {
    let src = r#"
def make_counter(start=0):
    count = start
    def counter():
        nonlocal count
        count += 1
        return count
    return counter

c = make_counter(10)
print(c())
print(c())
"#;
    assert_eq!(run_python(src), vec!["11", "12"]);
}

#[test]
fn test_py_class_decorator_registering_in_global_registry() {
    let src = r#"
REGISTRY = {}

def register(name):
    def decorator(cls):
        REGISTRY[name] = cls
        return cls
    return decorator

@register("plugin_a")
class PluginA:
    pass

@register("plugin_b")
class PluginB:
    pass

print(sorted(REGISTRY.keys()))
"#;
    assert_eq!(run_python(src), vec!["['plugin_a', 'plugin_b']"]);
}

#[test]
fn test_py_functools_wraps_assigned_and_updated_defaults() {
    let src = r#"
import functools

def custom_wraps(func):
    return functools.wraps(func, assigned=('__name__',), updated=())

def dec(func):
    @custom_wraps(func)
    def wrapper():
        return func()
    return wrapper

@dec
def target():
    """Docstring not copied"""
    return 100

print(target.__name__)
print(target.__doc__ is None)
"#;
    assert_eq!(run_python(src), vec!["target", "True"]);
}

#[test]
fn test_py_decorator_returning_generator_function() {
    let src = r#"
def multiply_yields(factor):
    def decorator(gen_func):
        def wrapper(*args, **kwargs):
            for val in gen_func(*args, **kwargs):
                yield val * factor
        return wrapper
    return decorator

@multiply_yields(10)
def generate_numbers():
    yield 1
    yield 2
    yield 3

print(list(generate_numbers()))
"#;
    assert_eq!(run_python(src), vec!["[10, 20, 30]"]);
}

#[test]
fn test_py_decorator_validating_argument_types() {
    let src = r#"
def enforce_types(*expected_types):
    def decorator(func):
        def wrapper(*args):
            for arg, expected in zip(args, expected_types):
                if not isinstance(arg, expected):
                    raise TypeError(f"Expected {expected.__name__}, got {type(arg).__name__}")
            return func(*args)
        return wrapper
    return decorator

@enforce_types(int, str)
def repeat_str(count, text):
    return text * count

print(repeat_str(3, "Hi"))
try:
    repeat_str("invalid", "Hi")
except TypeError as e:
    print("TypeError caught")
"#;
    assert_eq!(run_python(src), vec!["HiHiHi", "TypeError caught"]);
}

#[test]
fn test_py_decorator_caching_return_values() {
    let src = r#"
def memoize(func):
    cache = {}
    def wrapper(*args):
        if args not in cache:
            cache[args] = func(*args)
        return cache[args]
    return wrapper

@memoize
def fib(n):
    if n < 2:
        return n
    return fib(n - 1) + fib(n - 2)

print(fib(10))
"#;
    assert_eq!(run_python(src), vec!["55"]);
}

#[test]
fn test_py_async_coroutine_decorator() {
    let src = r#"
import asyncio

def async_log(func):
    async def wrapper(*args, **kwargs):
        print("async start")
        res = await func(*args, **kwargs)
        print("async end")
        return res
    return wrapper

@async_log
async def fetch_data():
    return "data"

print(asyncio.run(fetch_data()))
"#;
    assert_eq!(run_python(src), vec!["async start", "async end", "data"]);
}

#[test]
fn test_py_decorator_inspecting_function_annotations() {
    let src = r#"
def print_annotations(func):
    print(func.__annotations__)
    return func

@print_annotations
def process(x: int, y: str) -> bool:
    return True
"#;
    assert_eq!(
        run_python(src),
        vec!["{'x': <class 'int'>, 'y': <class 'str'>, 'return': <class 'bool'>}"]
    );
}

#[test]
fn test_py_closure_late_binding_fix_with_default_arg() {
    let src = r#"
# Without default arg, late binding captures final loop variable
funcs_late = [lambda: i for i in range(3)]
print([f() for f in funcs_late])

# Fixed with default parameter default arg
funcs_fixed = [lambda i=i: i for i in range(3)]
print([f() for f in funcs_fixed])
"#;
    assert_eq!(run_python(src), vec!["[2, 2, 2]", "[0, 1, 2]"]);
}
