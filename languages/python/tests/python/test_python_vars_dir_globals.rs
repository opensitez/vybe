// Python vars(), dir(), globals(), locals() introspection
use super::helpers::run_python;

#[test]
fn test_vars_object() {
    let script = r#"
class Point:
    def __init__(self, x, y):
        self.x = x
        self.y = y

p = Point(3, 4)
d = vars(p)
print(d['x'])
print(d['y'])
"#;
    assert_eq!(run_python(script), vec!["3", "4"]);
}

#[test]
fn test_vars_no_arg_is_locals() {
    let script = r#"
def func():
    a = 10
    b = 20
    v = vars()
    return sorted(v.keys())

print('a' in func())
print('b' in func())
"#;
    assert_eq!(run_python(script), vec!["True", "True"]);
}

#[test]
fn test_dir_lists_attributes() {
    let script = r#"
class Foo:
    x = 1
    def bar(self):
        pass

f = Foo()
attrs = dir(f)
print('x' in attrs)
print('bar' in attrs)
print('__class__' in attrs)
"#;
    assert_eq!(run_python(script), vec!["True", "True", "True"]);
}

#[test]
fn test_globals_contains_builtins() {
    let script = r#"
g = globals()
print('__name__' in g)
print(isinstance(g, dict))
"#;
    assert_eq!(run_python(script), vec!["True", "True"]);
}

#[test]
fn test_locals_in_function() {
    let script = r#"
def func(a, b):
    c = a + b
    l = locals()
    return sorted(l.keys())

print(func(1, 2))
"#;
    assert_eq!(run_python(script), vec!["['a', 'b', 'c', 'l']"]);
}

#[test]
fn test_dir_on_module() {
    let script = r#"
import math
attrs = dir(math)
print('pi' in attrs)
print('sqrt' in attrs)
print('cos' in attrs)
"#;
    assert_eq!(run_python(script), vec!["True", "True", "True"]);
}

#[test]
fn test_vars_modifies_object() {
    let script = r#"
class Obj:
    pass

o = Obj()
vars(o)['dynamic'] = 42
print(o.dynamic)
"#;
    assert_eq!(run_python(script), vec!["42"]);
}
