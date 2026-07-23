use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Scope & Namespaces — LEGB rule, global, nonlocal, exec(), eval(), custom globals/locals
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_global_variable_modification() {
    let src = r#"
counter = 0

def increment():
    global counter
    counter += 1

increment()
increment()
print(counter)
"#;
    assert_eq!(run_python(src), vec!["2"]);
}

#[test]
fn test_py_nonlocal_enclosing_scope_modification() {
    let src = r#"
def outer():
    x = "initial"
    def inner():
        nonlocal x
        x = "modified"
    inner()
    return x

print(outer())
"#;
    assert_eq!(run_python(src), vec!["modified"]);
}

#[test]
fn test_py_legb_scope_shadowing() {
    let src = r#"
x = "global"

def outer():
    x = "enclosing"
    def inner():
        x = "local"
        return x
    return inner(), x

res_inner, res_outer = outer()
print(res_inner)
print(res_outer)
print(x)
"#;
    assert_eq!(run_python(src), vec!["local", "enclosing", "global"]);
}

#[test]
fn test_py_unbound_local_error() {
    let src = r#"
x = 10

def bad_increment():
    try:
        x += 1  # python treats x as local due to assignment, causing UnboundLocalError
    except UnboundLocalError:
        print("UnboundLocalError caught")

bad_increment()
"#;
    assert_eq!(run_python(src), vec!["UnboundLocalError caught"]);
}

#[test]
fn test_py_eval_with_custom_globals_and_locals() {
    let src = r#"
g = {"x": 10, "y": 20}
l = {"y": 100}
result = eval("x + y", g, l)
print(result)
"#;
    assert_eq!(run_python(src), vec!["110"]);
}

#[test]
fn test_py_exec_code_string_with_dictionary_locals() {
    let src = r#"
code = """
total = sum(items)
msg = f"Total: {total}"
"""

loc = {"items": [1, 2, 3, 4, 5]}
exec(code, {}, loc)
print(loc["total"])
print(loc["msg"])
"#;
    assert_eq!(run_python(src), vec!["15", "Total: 15"]);
}

#[test]
fn test_py_globals_and_locals_builtins_inspection() {
    let src = r#"
def fn(arg):
    local_var = arg * 2
    print("arg" in locals())
    print("local_var" in locals())

fn(5)
print("fn" in globals())
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True"]);
}

#[test]
fn test_py_closure_cell_contents_inspection() {
    let src = r#"
def make_multiplier(factor):
    def multiply(number):
        return number * factor
    return multiply

double = make_multiplier(2)
print(double(5))
print(double.__closure__[0].cell_contents)
"#;
    assert_eq!(run_python(src), vec!["10", "2"]);
}

#[test]
fn test_py_class_scope_does_not_enclose_methods() {
    let src = r#"
x = "global"

class Foo:
    x = "class"
    def get_x(self):
        # Methods look in local then global, NOT class body scope!
        return x

f = Foo()
print(f.get_x())
"#;
    assert_eq!(run_python(src), vec!["global"]);
}

#[test]
fn test_py_comprehension_scope_isolation_py3() {
    let src = r#"
x = "outside"
lst = [x for x in range(3)]
print(lst)
print(x)  # x in outer scope not leaked in Python 3!
"#;
    assert_eq!(run_python(src), vec!["[0, 1, 2]", "outside"]);
}
