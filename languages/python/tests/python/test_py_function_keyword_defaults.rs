use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Function Keywords & Defaults — positional-only, keyword-only, default arg evaluation, forwarding, function attributes
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_positional_only_slash_delimiter() {
    let src = r#"
def format_point(x, y, /, label="Point"):
    return f"{label}({x}, {y})"

print(format_point(10, 20))
print(format_point(10, 20, label="P1"))
try:
    format_point(x=10, y=20)
except TypeError:
    print("TypeError: positional-only")
"#;
    assert_eq!(
        run_python(src),
        vec!["Point(10, 20)", "P1(10, 20)", "TypeError: positional-only"]
    );
}

#[test]
fn test_py_keyword_only_star_delimiter() {
    let src = r#"
def connect(host, port, *, timeout=30, retry=True):
    return f"{host}:{port} (timeout={timeout}, retry={retry})"

print(connect("localhost", 8080, timeout=10))
try:
    connect("localhost", 8080, 10)
except TypeError:
    print("TypeError: keyword-only required")
"#;
    assert_eq!(
        run_python(src),
        vec![
            "localhost:8080 (timeout=10, retry=True)",
            "TypeError: keyword-only required"
        ]
    );
}

#[test]
fn test_py_mutable_default_idiomatic_none_fix() {
    let src = r#"
def append_to_list(element, target=None):
    if target is None:
        target = []
    target.append(element)
    return target

l1 = append_to_list(1)
l2 = append_to_list(2)
print(l1)
print(l2)
"#;
    assert_eq!(run_python(src), vec!["[1]", "[2]"]);
}

#[test]
fn test_py_forwarding_args_and_kwargs() {
    let src = r#"
def target(a, b, c=0, verbose=False):
    return f"a={a}, b={b}, c={c}, verbose={verbose}"

def wrapper(*args, **kwargs):
    return target(*args, **kwargs)

print(wrapper(1, 2))
print(wrapper(1, 2, 3, verbose=True))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "a=1, b=2, c=0, verbose=False",
            "a=1, b=2, c=3, verbose=True"
        ]
    );
}

#[test]
fn test_py_function_attributes_assignment() {
    let src = r#"
def worker():
    worker.calls += 1
    return worker.calls

worker.calls = 0
print(worker())
print(worker())
print(worker.calls)
"#;
    assert_eq!(run_python(src), vec!["1", "2", "2"]);
}

#[test]
fn test_py_default_values_tuple_inspection() {
    let src = r#"
def fn(a, b=10, c="test"):
    pass

print(fn.__defaults__)
"#;
    assert_eq!(run_python(src), vec!["(10, 'test')"]);
}

#[test]
fn test_py_kwdefaults_dictionary_inspection() {
    let src = r#"
def fn(a, *, b=20, c="kw"):
    pass

print(fn.__kwdefaults__)
"#;
    assert_eq!(run_python(src), vec!["{'b': 20, 'c': 'kw'}"]);
}

#[test]
fn test_py_dictionary_unpacking_matching_kw_params() {
    let src = r#"
def configure(host="localhost", port=8080):
    return f"{host}:{port}"

config_dict = {"port": 9000, "host": "127.0.0.1"}
print(configure(**config_dict))
"#;
    assert_eq!(run_python(src), vec!["127.0.0.1:9000"]);
}

#[test]
fn test_py_function_code_object_co_argcount() {
    let src = r#"
def sample(a, b, c=1, *, d=2):
    pass

code = sample.__code__
print(code.co_argcount)      # 3 positional/kw args
print(code.co_kwonlyargcount)  # 1 kw-only arg
"#;
    assert_eq!(run_python(src), vec!["3", "1"]);
}

#[test]
fn test_py_partial_application_parameter_binding() {
    let src = r#"
from functools import partial

def multiply(x, y):
    return x * y

double = partial(multiply, 2)
triple = partial(multiply, 3)

print(double(5))
print(triple(5))
"#;
    assert_eq!(run_python(src), vec!["10", "15"]);
}
