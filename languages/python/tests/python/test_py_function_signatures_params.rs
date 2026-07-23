use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Function Signatures & Parameters — positional-only, keyword-only, *args, **kwargs, defaults, signature binding
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_positional_only_slash_syntax() {
    let src = r#"
def pos_only(a, b, /, c=10):
    return a + b + c

print(pos_only(1, 2))
print(pos_only(1, 2, c=30))

try:
    pos_only(a=1, b=2)
except TypeError:
    print("TypeError: positional-only")
"#;
    assert_eq!(
        run_python(src),
        vec!["13", "33", "TypeError: positional-only"]
    );
}

#[test]
fn test_py_keyword_only_asterisk_syntax() {
    let src = r#"
def kw_only(a, *, b, c=100):
    return a + b + c

print(kw_only(1, b=2))
print(kw_only(1, b=2, c=3))

try:
    kw_only(1, 2)
except TypeError:
    print("TypeError: keyword-only required")
"#;
    assert_eq!(
        run_python(src),
        vec!["103", "6", "TypeError: keyword-only required"]
    );
}

#[test]
fn test_py_combined_positional_and_keyword_only() {
    let src = r#"
def full(pos, /, pos_or_kw, *, kw):
    return f"{pos}-{pos_or_kw}-{kw}"

print(full(1, 2, kw=3))
print(full(1, pos_or_kw=2, kw=3))
"#;
    assert_eq!(run_python(src), vec!["1-2-3", "1-2-3"]);
}

#[test]
fn test_py_var_args_and_var_kwargs_forwarding() {
    let src = r#"
def target(a, b, c=0, debug=False):
    return f"a={a}, b={b}, c={c}, debug={debug}"

def proxy(*args, **kwargs):
    return target(*args, **kwargs)

print(proxy(1, 2))
print(proxy(1, 2, 3, debug=True))
"#;
    assert_eq!(
        run_python(src),
        vec!["a=1, b=2, c=0, debug=False", "a=1, b=2, c=3, debug=True"]
    );
}

#[test]
fn test_py_mutable_default_argument_trap_and_idiom() {
    let src = r#"
def bad_add(item, target=[]):
    target.append(item)
    return target

def good_add(item, target=None):
    if target is None:
        target = []
    target.append(item)
    return target

print(bad_add(1))
print(bad_add(2))  # state persists!

print(good_add(1))
print(good_add(2))  # independent lists
"#;
    assert_eq!(run_python(src), vec!["[1]", "[1, 2]", "[1]", "[2]"]);
}

#[test]
fn test_py_default_values_evaluated_at_definition() {
    let src = r#"
count = 0
def get_count():
    global count
    count += 1
    return count

def fn(val=get_count()):
    return val

print(fn())
print(fn())  # get_count not called again
"#;
    assert_eq!(run_python(src), vec!["1", "1"]);
}

#[test]
fn test_py_inspect_signature_bind_partial() {
    let src = r#"
import inspect

def greet(name, age=30, *, city="NY"):
    pass

sig = inspect.signature(greet)
bound = sig.bind_partial("Alice")
bound.apply_defaults()
print(dict(bound.arguments))
"#;
    assert_eq!(
        run_python(src),
        vec!["{'name': 'Alice', 'age': 30, 'city': 'NY'}"]
    );
}

#[test]
fn test_py_unpacking_dict_keys_matching_parameters() {
    let src = r#"
def config(host="localhost", port=8080, timeout=30):
    return f"{host}:{port} (timeout={timeout}s)"

params = {"port": 9090, "timeout": 60}
print(config(**params))
"#;
    assert_eq!(run_python(src), vec!["localhost:9090 (timeout=60s)"]);
}

#[test]
fn test_py_parameter_annotations_access() {
    let src = r#"
def add(x: int, y: int = 0) -> int:
    return x + y

print(add.__annotations__)
"#;
    assert_eq!(
        run_python(src),
        vec!["{'x': <class 'int'>, 'y': <class 'int'>, 'return': <class 'int'>}"]
    );
}

#[test]
fn test_py_kwarg_name_clash_prevention() {
    let src = r#"
def build_dict(**kwargs):
    return kwargs

print(build_dict(name="item", value=42, type="widget"))
"#;
    assert_eq!(
        run_python(src),
        vec!["{'name': 'item', 'value': 42, 'type': 'widget'}"]
    );
}
