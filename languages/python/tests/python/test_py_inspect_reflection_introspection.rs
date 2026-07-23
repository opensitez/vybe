use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Inspect & Reflection Introspection — signature, getmembers, iscoroutinefunction, stack, currentframe, annotations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_inspect_signature_parameters_kind() {
    let src = r#"
import inspect

def func(a, /, b, *, c=10): pass

sig = inspect.signature(func)
params = sig.parameters
print(params["a"].kind.name)
print(params["b"].kind.name)
print(params["c"].kind.name)
print(params["c"].default)
"#;
    assert_eq!(
        run_python(src),
        vec![
            "POSITIONAL_ONLY",
            "POSITIONAL_OR_KEYWORD",
            "KEYWORD_ONLY",
            "10"
        ]
    );
}

#[test]
fn test_py_inspect_getmembers_predicates() {
    let src = r#"
import inspect

class MyClass:
    class_var = 42
    def method(self): pass

funcs = [name for name, _ in inspect.getmembers(MyClass, inspect.isfunction)]
print(funcs)
"#;
    assert_eq!(run_python(src), vec!["['method']"]);
}

#[test]
fn test_py_inspect_coroutine_generator_predicates() {
    let src = r#"
import inspect

async def coro(): pass
def gen(): yield 1
def normal(): pass

print(inspect.iscoroutinefunction(coro))
print(inspect.isgeneratorfunction(gen))
print(inspect.isfunction(normal))
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True"]);
}

#[test]
fn test_py_inspect_currentframe_f_code_name() {
    let src = r#"
import inspect

def test_fn():
    frame = inspect.currentframe()
    return frame.f_code.co_name

print(test_fn())
"#;
    assert_eq!(run_python(src), vec!["test_fn"]);
}

#[test]
fn test_py_inspect_stack_frame_traversal() {
    let src = r#"
import inspect

def level2():
    stack = inspect.stack()
    return [frame.function for frame in stack[:3]]

def level1():
    return level2()

funcs = level1()
print("level2" in funcs and "level1" in funcs)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_inspect_get_annotations_eval() {
    let src = r#"
import inspect

def compute(x: int, y: float = 1.0) -> float:
    return x * y

annotations = inspect.get_annotations(compute)
print(annotations["x"] is int)
print(annotations["return"] is float)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_inspect_getdoc_clean_docstring() {
    let src = r#"
import inspect

def documented():
    """
    First line of doc.
    Second line of doc.
    """
    pass

doc = inspect.getdoc(documented)
print(doc)
"#;
    assert_eq!(
        run_python(src),
        vec!["First line of doc.\nSecond line of doc."]
    );
}

#[test]
fn test_py_inspect_isclass_ismethod_ismodule() {
    let src = r#"
import inspect, sys

class Sample:
    def method(self): pass

s = Sample()
print(inspect.isclass(Sample))
print(inspect.ismethod(s.method))
print(inspect.ismodule(sys))
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True"]);
}

#[test]
fn test_py_inspect_signature_bind_partial() {
    let src = r#"
import inspect

def connect(host, port=8080, *, timeout=30): pass

sig = inspect.signature(connect)
bound = sig.bind_partial("localhost")
bound.apply_defaults()
print(dict(bound.arguments))
"#;
    assert_eq!(
        run_python(src),
        vec!["{'host': 'localhost', 'port': 8080, 'timeout': 30}"]
    );
}

#[test]
fn test_py_inspect_getmodule_function_owner() {
    let src = r#"
import inspect, json

mod = inspect.getmodule(json.dumps)
print(mod.__name__)
"#;
    assert_eq!(run_python(src), vec!["json"]);
}
