use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: closures + higher-order functions — closure over variables, lambda, partial application, currying, HOF patterns
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_closure_captures_enclosing_scope() {
    let src = r#"
def make_adder(n):
    def adder(x):
        return x + n
    return adder

add5 = make_adder(5)
add10 = make_adder(10)
print(add5(3))
print(add10(3))
print(add5.__closure__[0].cell_contents)
"#;
    assert_eq!(run_python(src), vec!["8", "13", "5"]);
}

#[test]
fn test_py_closure_nonlocal_mutation() {
    let src = r#"
def make_counter(start=0):
    count = start

    def increment(by=1):
        nonlocal count
        count += by
        return count

    def reset():
        nonlocal count
        count = start

    return increment, reset

inc, reset = make_counter(10)
print(inc())
print(inc(5))
reset()
print(inc())
"#;
    assert_eq!(run_python(src), vec!["11", "16", "11"]);
}

#[test]
fn test_py_closure_loop_capture_gotcha() {
    let src = r#"
# Late binding gotcha
fns_bad = [lambda x: x + i for i in range(4)]
print(fns_bad[0](0))  # all capture same i=3

# Fixed with default argument
fns_good = [lambda x, i=i: x + i for i in range(4)]
print([f(0) for f in fns_good])
"#;
    assert_eq!(run_python(src), vec!["3", "[0, 1, 2, 3]"]);
}

#[test]
fn test_py_higher_order_map_filter_reduce() {
    let src = r#"
from functools import reduce

nums = [1, 2, 3, 4, 5]
doubled = list(map(lambda x: x * 2, nums))
evens = list(filter(lambda x: x % 2 == 0, nums))
total = reduce(lambda a, b: a + b, nums)
print(doubled)
print(evens)
print(total)
"#;
    assert_eq!(run_python(src), vec!["[2, 4, 6, 8, 10]", "[2, 4]", "15"]);
}

#[test]
fn test_py_lambda_key_functions() {
    let src = r#"
data = [{"name": "Charlie", "age": 35}, {"name": "Alice", "age": 25}, {"name": "Bob", "age": 30}]
by_age = sorted(data, key=lambda p: p["age"])
print([p["name"] for p in by_age])
by_name = sorted(data, key=lambda p: p["name"])
print([p["name"] for p in by_name])
"#;
    assert_eq!(
        run_python(src),
        vec!["['Alice', 'Bob', 'Charlie']", "['Alice', 'Bob', 'Charlie']"]
    );
}

#[test]
fn test_py_closure_memoize_decorator() {
    let src = r#"
def memoize(func):
    cache = {}
    def wrapper(*args):
        if args not in cache:
            cache[args] = func(*args)
        return cache[args]
    return wrapper

calls = [0]

@memoize
def expensive(n):
    calls[0] += 1
    return n ** 2

print(expensive(5))
print(expensive(5))
print(expensive(6))
print(calls[0])  # only 2 unique calls
"#;
    assert_eq!(run_python(src), vec!["25", "25", "36", "2"]);
}

#[test]
fn test_py_currying_with_closures() {
    let src = r#"
def curry(func):
    import inspect
    n = len(inspect.signature(func).parameters)

    def accumulate(args):
        if len(args) >= n:
            return func(*args)
        return lambda *new_args: accumulate(args + new_args)

    return lambda *args: accumulate(args)

@curry
def add3(a, b, c):
    return a + b + c

print(add3(1)(2)(3))
print(add3(1, 2)(3))
print(add3(1)(2, 3))
"#;
    assert_eq!(run_python(src), vec!["6", "6", "6"]);
}

#[test]
fn test_py_function_composition() {
    let src = r#"
from functools import reduce

def compose(*funcs):
    def composed(x):
        return reduce(lambda v, f: f(v), reversed(funcs), x)
    return composed

double = lambda x: x * 2
add_one = lambda x: x + 1
square = lambda x: x ** 2

pipeline = compose(double, add_one, square)
print(pipeline(3))  # double(add_one(square(3))) = double(add_one(9)) = double(10) = 20
"#;
    assert_eq!(run_python(src), vec!["20"]);
}

#[test]
fn test_py_closure_factory_pattern() {
    let src = r#"
def validator_factory(**rules):
    def validate(data: dict) -> list:
        errors = []
        for field, rule in rules.items():
            if field not in data:
                errors.append(f"missing: {field}")
            elif not rule(data[field]):
                errors.append(f"invalid: {field}")
        return errors
    return validate

check = validator_factory(
    name=lambda v: isinstance(v, str) and len(v) > 0,
    age=lambda v: isinstance(v, int) and 0 < v < 150
)
print(check({"name": "Alice", "age": 30}))
print(check({"name": "", "age": 200}))
"#;
    assert_eq!(
        run_python(src),
        vec!["[]", "['invalid: name', 'invalid: age']"]
    );
}

#[test]
fn test_py_higher_order_sorted_itemgetter() {
    let src = r#"
from operator import itemgetter, attrgetter

data = [{"x": 3, "y": 1}, {"x": 1, "y": 4}, {"x": 2, "y": 2}]
print(sorted(data, key=itemgetter("x")))

class Point:
    def __init__(self, x, y):
        self.x = x
        self.y = y
    def __repr__(self):
        return f"Point({self.x},{self.y})"

points = [Point(3, 1), Point(1, 4), Point(2, 2)]
print(sorted(points, key=attrgetter("x")))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "[{'x': 1, 'y': 4}, {'x': 2, 'y': 2}, {'x': 3, 'y': 1}]",
            "[Point(1,4), Point(2,2), Point(3,1)]"
        ]
    );
}

#[test]
fn test_py_lambda_conditional_expression() {
    let src = r#"
classify = lambda x: "positive" if x > 0 else ("zero" if x == 0 else "negative")
print(classify(5))
print(classify(0))
print(classify(-3))
print(list(map(classify, [1, -2, 0, 3, -4])))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "positive",
            "zero",
            "negative",
            "['positive', 'negative', 'zero', 'positive', 'negative']"
        ]
    );
}

#[test]
fn test_py_closure_shared_state_between_closures() {
    let src = r#"
def make_stack():
    data = []

    def push(val):
        data.append(val)

    def pop():
        return data.pop() if data else None

    def peek():
        return data[-1] if data else None

    return push, pop, peek

push, pop, peek = make_stack()
push(1)
push(2)
push(3)
print(peek())
print(pop())
print(peek())
"#;
    assert_eq!(run_python(src), vec!["3", "3", "2"]);
}
