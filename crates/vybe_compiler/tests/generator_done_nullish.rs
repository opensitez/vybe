use vybe_runtime::{Value, VM};

fn execute(source: &str) -> Value {
    static INIT: std::sync::Once = std::sync::Once::new();
    INIT.call_once(vybe_language_python::register);
    let ast = vybe_language_python::parse(source).unwrap();
    let profile = vybe_compiler::profile::parse_profile(vybe_language_python::profile_source()).unwrap();
    let chunks = vybe_compiler::Compiler::with_profile(profile).compile(&ast).unwrap();
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    vm.run(chunks).unwrap();
    vm.global("answer").expect("fixture stores its result").clone()
}

#[test]
fn missing_completion_marker_resumes_real_python_generator() {
    let result = execute("def values():\n    yield 1\n    yield 2\nanswer = list(values())\n");
    let Value::Object(array) = result else { panic!("expected yielded array: {result:?}"); };
    let array = array.lock().unwrap();
    let vybe_runtime::value::ObjectKind::Array(items) = &array.kind else { panic!("expected array"); };
    assert_eq!(items.len(), 2);
    assert_eq!(items[0].as_f64(), 1.0);
    assert_eq!(items[1].as_f64(), 2.0);
}

#[test]
fn completion_remains_exhausted_on_second_drain() {
    let result = execute("def values():\n    yield 1\n    yield 2\ng = values()\nfirst = list(g)\nsecond = list(g)\nanswer = len(first) * 10 + len(second)\n");
    assert_eq!(result.as_f64(), 20.0);
}

#[test]
fn sum_and_nested_loop_resume_all_yields() {
    let result = execute("def values():\n    for i in range(4):\n        yield i\nanswer = sum(values())\n");
    assert_eq!(result.as_f64(), 6.0);
}

#[test]
fn concatenated_generator_lists_keep_list_kind() {
    let result = execute("def a():\n    yield 1\ndef b():\n    yield 2\nanswer = repr(list(a()) + list(b()))\n");
    assert_eq!(result, Value::String("[1, 2]".into()));
}

#[test]
fn absent_tuple_tag_differs_from_present_tuple_tag() {
    let result = execute("answer = [isinstance([], tuple), isinstance([1, 2], tuple), isinstance((1, 2), tuple)]\n");
    let Value::Object(array) = result else { panic!("expected flags"); };
    let array = array.lock().unwrap();
    let vybe_runtime::value::ObjectKind::Array(items) = &array.kind else { panic!("expected array"); };
    assert_eq!(items, &vec![Value::Bool(false), Value::Bool(false), Value::Bool(true)]);
}
