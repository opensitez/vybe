use vybe_ast::{Expression, Statement, StmtKind};
use vybe_compiler::Compiler;
use vybe_runtime::{Value, VM};

fn result(source: &str) -> Vec<Value> {
    vybe_language_python::register();
    let mut module = vybe_language_python::parse(source).unwrap();
    assert_eq!(module.directives.missing_arg_is_undefined, Some(true));
    module.body.push(Statement::new(StmtKind::Return(Some(Expression::ident("result")))));
    let profile = vybe_compiler::profile::parse_profile(vybe_language_python::profile_source()).unwrap();
    let chunks = Compiler::with_profile(profile).compile(&module).unwrap();
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    let output = vm.run(chunks).unwrap();
    let Value::Object(object) = output else { panic!("expected result list, got {output:?}") };
    let object = object.lock().unwrap();
    let vybe_runtime::value::ObjectKind::Array(items) = &object.kind else { panic!("expected result list") };
    items.clone()
}

#[test]
fn omitted_stream_constructor_and_seek_arguments_use_defaults() {
    let values = result("import io\nb = io.StringIO()\nempty = b.getvalue()\nb.write('Hello ')\nb.write('World')\nb.seek(0)\nresult = [empty, b.read(), b.getvalue(), b.tell()]\n");
    assert_eq!(values.len(), 4);
    for (actual, expected) in values[..3].iter().zip(["", "Hello World", "Hello World"]) {
        assert!(matches!(actual, Value::String(s) if s.as_ref() == expected), "{actual:?}");
    }
    assert_eq!(values[3].as_f64(), 11.0);
}

#[test]
fn reported_stringio_corpus_case_passes_its_original_check() {
    let source = format!("{}\nresult = [__buf]\n", include_str!("../../../tests/python/py_io_streams_buffers_bytes/test_py_io_stringio_read_write_seek.py"));
    let values = result(&source);
    assert!(matches!(&values[0], Value::String(s) if s.as_ref() == "Hello World\nHello World\n"));
}

#[test]
fn explicit_stream_initial_values_and_empty_bytes_are_preserved() {
    let values = result("import io\na = io.StringIO('abc')\nb = io.StringIO(None)\nc = io.BytesIO()\nresult = [a.getvalue(), b.getvalue(), len(c.getvalue())]\n");
    assert!(matches!(&values[0], Value::String(s) if s.as_ref() == "abc"));
    assert!(matches!(&values[1], Value::String(s) if s.is_empty()));
    assert_eq!(values[2].as_f64(), 0.0);
}

#[test]
fn constructor_and_method_defaults_distinguish_none_and_zero() {
    let values = result("class Box:\n    def __init__(self, value=9):\n        self.value = value\n    def echo(self, value=7):\n        return value\na = Box()\nb = Box(None)\nc = Box(0)\nresult = [a.value, b.value, c.value, a.echo(), a.echo(None), a.echo(0)]\n");
    assert_eq!(values[0].as_f64(), 9.0);
    assert!(matches!(values[1], Value::Null));
    assert_eq!(values[2].as_f64(), 0.0);
    assert_eq!(values[3].as_f64(), 7.0);
    assert!(matches!(values[4], Value::Null));
    assert_eq!(values[5].as_f64(), 0.0);
}

#[test]
fn inherited_constructor_defaults_keep_explicit_none() {
    let values = result("class Base:\n    def __init__(self, value=5):\n        self.value = value\nclass Child(Base):\n    pass\nresult = [Child().value, Child(None).value, Child(0).value]\n");
    assert_eq!(values[0].as_f64(), 5.0);
    assert!(matches!(values[1], Value::Null));
    assert_eq!(values[2].as_f64(), 0.0);
}
