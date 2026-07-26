// Python functools.partial and functools.reduce
use super::helpers::run_python;

#[test]
fn test_partial_basic() {
    let script = r#"
from functools import partial

def power(base, exp):
    return base ** exp

square = partial(power, exp=2)
cube = partial(power, exp=3)
print(square(4))
print(cube(3))
"#;
    assert_eq!(run_python(script), vec!["16", "27"]);
}

#[test]
fn test_partial_positional() {
    let script = r#"
from functools import partial

def greet(greeting, name):
    return f"{greeting}, {name}!"

say_hello = partial(greet, "Hello")
print(say_hello("Alice"))
print(say_hello("Bob"))
"#;
    assert_eq!(run_python(script), vec!["Hello, Alice!", "Hello, Bob!"]);
}

#[test]
fn test_partial_keywords() {
    let script = r#"
from functools import partial

def log(level, message, prefix="LOG"):
    return f"[{prefix}:{level}] {message}"

error_log = partial(log, "ERROR", prefix="APP")
print(error_log("Something failed"))
"#;
    assert_eq!(run_python(script), vec!["[APP:ERROR] Something failed"]);
}

#[test]
fn test_reduce_sum() {
    let script = r#"
from functools import reduce
total = reduce(lambda a, b: a + b, [1, 2, 3, 4, 5])
print(total)
"#;
    assert_eq!(run_python(script), vec!["15"]);
}

#[test]
fn test_reduce_with_initial() {
    let script = r#"
from functools import reduce
total = reduce(lambda a, b: a + b, [1, 2, 3], 100)
print(total)
"#;
    assert_eq!(run_python(script), vec!["106"]);
}

#[test]
fn test_reduce_product() {
    let script = r#"
from functools import reduce
product = reduce(lambda a, b: a * b, range(1, 6))
print(product)
"#;
    assert_eq!(run_python(script), vec!["120"]);
}

#[test]
fn test_reduce_max() {
    let script = r#"
from functools import reduce
biggest = reduce(lambda a, b: a if a > b else b, [3, 1, 4, 1, 5, 9, 2, 6])
print(biggest)
"#;
    assert_eq!(run_python(script), vec!["9"]);
}

#[test]
fn test_partial_func_attribute() {
    let script = r#"
from functools import partial

def add(x, y):
    return x + y

add5 = partial(add, 5)
print(add5.func is add)
print(add5.args)
print(add5(3))
"#;
    assert_eq!(run_python(script), vec!["True", "(5,)", "8"]);
}
