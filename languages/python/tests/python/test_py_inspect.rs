use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: inspect module — signatures, source, frame introspection, module members, is* predicates, annotations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_inspect_signature_and_parameters() {
    let src = r#"
import inspect

def create_user(name: str, age: int = 18, *, admin: bool = False) -> dict:
    return {"name": name, "age": age, "admin": admin}

sig = inspect.signature(create_user)
params = sig.parameters
print(list(params.keys()))
print(params["age"].default)
print(params["admin"].kind.name)
"#;
    assert_eq!(
        run_python(src),
        vec!["['name', 'age', 'admin']", "18", "KEYWORD_ONLY"]
    );
}

#[test]
fn test_py_inspect_getmembers() {
    let src = r#"
import inspect

class MyClass:
    class_var = 42

    def method(self):
        pass

    @classmethod
    def class_method(cls):
        pass

    @staticmethod
    def static_method():
        pass

methods = [name for name, m in inspect.getmembers(MyClass, inspect.isfunction)]
print("method" in methods)
print("class_var" not in methods)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_inspect_isfunction_ismethod_isclass() {
    let src = r#"
import inspect

class Foo:
    def method(self):
        pass

def func():
    pass

print(inspect.isfunction(func))
print(inspect.isfunction(Foo.method))
print(inspect.isclass(Foo))
print(inspect.ismethod(Foo().method))
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True", "True"]);
}

#[test]
fn test_py_inspect_get_annotations() {
    let src = r#"
import inspect

def compute(x: int, y: float = 1.0) -> float:
    return x * y

hints = inspect.get_annotations(compute)
print(hints)
"#;
    assert_eq!(
        run_python(src),
        vec!["{'x': <class 'int'>, 'y': <class 'float'>, 'return': <class 'float'>}"]
    );
}

#[test]
fn test_py_inspect_getdoc() {
    let src = r#"
import inspect

def documented():
    """This is a docstring.

    With multiple lines.
    """
    pass

doc = inspect.getdoc(documented)
print(doc.splitlines()[0])
print(len(doc.splitlines()) > 1)
"#;
    assert_eq!(run_python(src), vec!["This is a docstring.", "True"]);
}

#[test]
fn test_py_inspect_currentframe_and_stack() {
    let src = r#"
import inspect

def inner():
    frame = inspect.currentframe()
    return frame.f_code.co_name

def outer():
    return inner()

print(outer())
print(inner())
"#;
    assert_eq!(run_python(src), vec!["inner", "inner"]);
}

#[test]
fn test_py_inspect_parameter_kinds() {
    let src = r#"
import inspect

def complex_func(pos_only, /, regular, *args, kw_only, **kwargs):
    pass

sig = inspect.signature(complex_func)
for name, param in sig.parameters.items():
    print(f"{name}: {param.kind.name}")
"#;
    assert_eq!(
        run_python(src),
        vec![
            "pos_only: POSITIONAL_ONLY",
            "regular: POSITIONAL_OR_KEYWORD",
            "args: VAR_POSITIONAL",
            "kw_only: KEYWORD_ONLY",
            "kwargs: VAR_KEYWORD"
        ]
    );
}

#[test]
fn test_py_inspect_isgeneratorfunction() {
    let src = r#"
import inspect

def regular():
    return 1

def generator():
    yield 1

async def coroutine():
    pass

async def async_gen():
    yield 1

print(inspect.isgeneratorfunction(regular))
print(inspect.isgeneratorfunction(generator))
print(inspect.iscoroutinefunction(coroutine))
print(inspect.isasyncgenfunction(async_gen))
"#;
    assert_eq!(run_python(src), vec!["False", "True", "True", "True"]);
}

#[test]
fn test_py_inspect_getmodule() {
    let src = r#"
import inspect, math

print(inspect.getmodule(math.sqrt).__name__)
print(inspect.getmodule(len))  # builtin - might be None
"#;
    assert_eq!(run_python(src), vec!["math", "None"]);
}

#[test]
fn test_py_inspect_signature_bind() {
    let src = r#"
import inspect

def func(a, b, c=10, *args, key=None):
    pass

sig = inspect.signature(func)
bound = sig.bind(1, 2, 3, 4, 5, key="val")
bound.apply_defaults()
print(dict(bound.arguments))
"#;
    assert_eq!(
        run_python(src),
        vec!["{'a': 1, 'b': 2, 'c': 3, 'args': (4, 5), 'key': 'val'}"]
    );
}
