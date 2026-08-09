// Python lambda and closures — captures, HOF, currying, recursion via variable
use super::helpers::run_python;

#[test]
fn test_lambda_basic() {
    let script = r#"
square = lambda x: x * x
print(square(7))
"#;
    assert_eq!(run_python(script), vec!["49"]);
}

#[test]
fn test_lambda_in_sorted() {
    let script = r#"
words = ["banana", "apple", "cherry", "date"]
print(sorted(words, key=lambda w: len(w)))
"#;
    assert_eq!(
        run_python(script),
        vec!["['date', 'apple', 'banana', 'cherry']"]
    );
}

#[test]
fn test_closure_captures_outer() {
    let script = r#"
def make_adder(n):
    return lambda x: x + n

add5 = make_adder(5)
add10 = make_adder(10)
print(add5(3))
print(add10(3))
"#;
    assert_eq!(run_python(script), vec!["8", "13"]);
}

#[test]
fn test_closure_counter() {
    let script = r#"
def counter(start=0):
    count = [start]
    def inc():
        count[0] += 1
        return count[0]
    return inc

c = counter(10)
print(c())
print(c())
print(c())
"#;
    assert_eq!(run_python(script), vec!["11", "12", "13"]);
}

#[test]
fn test_lambda_map_filter() {
    let script = r#"
nums = list(range(1, 11))
evens = list(filter(lambda x: x % 2 == 0, nums))
doubled = list(map(lambda x: x * 2, evens))
print(doubled)
"#;
    assert_eq!(run_python(script), vec!["[4, 8, 12, 16, 20]"]);
}

#[test]
fn test_closure_late_binding_fixed() {
    let script = r#"
# classic loop closure bug — fix using default arg
funcs = [lambda x, i=i: x + i for i in range(3)]
print(funcs[0](10))
print(funcs[1](10))
print(funcs[2](10))
"#;
    assert_eq!(run_python(script), vec!["10", "11", "12"]);
}

#[test]
fn test_lambda_immediate_call() {
    let script = r#"
result = (lambda a, b: a ** b)(2, 10)
print(result)
"#;
    assert_eq!(run_python(script), vec!["1024"]);
}

#[test]
fn test_closure_nonlocal() {
    let script = r#"
def outer():
    x = 0
    def inner():
        nonlocal x
        x += 1
        return x
    return inner

inc = outer()
print(inc())
print(inc())
print(inc())
"#;
    assert_eq!(run_python(script), vec!["1", "2", "3"]);
}
