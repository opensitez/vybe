// Python scope — LEGB, global, nonlocal, class scope, comprehension scope
use super::helpers::run_python;

#[test]
fn test_legb_local_shadows_global() {
    let script = r#"
x = "global"

def func():
    x = "local"
    return x

print(func())
print(x)
"#;
    assert_eq!(run_python(script), vec!["local", "global"]);
}

#[test]
fn test_global_keyword() {
    let script = r#"
count = 0

def increment():
    global count
    count += 1

increment()
increment()
increment()
print(count)
"#;
    assert_eq!(run_python(script), vec!["3"]);
}

#[test]
fn test_nonlocal_keyword() {
    let script = r#"
def outer():
    x = 0
    def inner():
        nonlocal x
        x += 10
    inner()
    inner()
    return x

print(outer())
"#;
    assert_eq!(run_python(script), vec!["20"]);
}

#[test]
fn test_enclosing_scope_read() {
    let script = r#"
def make_greeting(name):
    def greet():
        return f"Hello, {name}"
    return greet

greet = make_greeting("Alice")
print(greet())
"#;
    assert_eq!(run_python(script), vec!["Hello, Alice"]);
}

#[test]
fn test_class_scope_not_inherited_by_methods() {
    let script = r#"
x = "global"

class MyClass:
    x = "class"
    def method(self):
        return x  # sees global, not class x

obj = MyClass()
print(obj.method())
print(MyClass.x)
"#;
    assert_eq!(run_python(script), vec!["global", "class"]);
}

#[test]
fn test_comprehension_scope_isolated() {
    let script = r#"
x = 10
result = [x for x in range(3)]
print(x)  # x unchanged after comprehension
print(result)
"#;
    assert_eq!(run_python(script), vec!["10", "[0, 1, 2]"]);
}

#[test]
fn test_unbound_local_error() {
    let script = r#"
x = 10

def bad():
    try:
        print(x)
        x = 20
    except UnboundLocalError:
        print("UnboundLocalError")

bad()
"#;
    assert_eq!(run_python(script), vec!["UnboundLocalError"]);
}
