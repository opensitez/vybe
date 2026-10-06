//! ECMA-native string adapter checks without guest compiler dependencies.
use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::VM;

fn text(input: &str) -> Value {
    Value::String(std::sync::Arc::from(input))
}

fn invoke(vm: &mut VM, name: &str, args: &[Value]) -> Value {
    let index = *vm.host_registry
        .get(&("ecma:string".to_owned(), name.to_owned()))
        .expect("string adapter registered");
    let mut callable = Object::new();
    callable.kind = ObjectKind::HostFunction(index);
    vm.invoke(&Value::Object(vybe_runtime::heap::alloc(callable)), args)
        .expect("native adapter invocation succeeds")
}

fn main() {
    let mut vm = VM::new();
    vybe_platform_ecma::string::register(&mut vm);
    for (input, pattern, expected) in [
        ("xa", "(a)(b)?", Some(vec![text("a"), text("a"), Value::Undefined])),
        ("xab", "(a)(b)?", Some(vec![text("ab"), text("a"), text("b")])),
        ("bbb", "a*", Some(vec![text("")])),
        ("item-42", "(?P<word>[a-z]+)-([0-9]+)", Some(vec![text("item-42"), text("item"), text("42")])),
        ("x\u{e9}\u{e9}", "\u{e9}+(\u{1f600})?", Some(vec![text("\u{e9}\u{e9}"), Value::Undefined])),
        ("not present", "word", None),
        ("anything", "[", None),
    ] {
        let result = invoke(&mut vm, "match", &[text(input), text(pattern)]);
        if let Some(expected) = expected {
            let Value::Object(object) = result else { panic!("match array expected"); };
            let guard = object.lock().unwrap();
            let ObjectKind::Array(values) = &guard.kind else { panic!("array storage expected"); };
            assert_eq!(values, &expected, "pattern={pattern:?}");
        } else {
            assert_eq!(result, Value::Null, "pattern={pattern:?}");
        }
    }
    for (name, input, expected) in [
        ("encodeURI", "https://example.test/path?field=value#anchor", "https://example.test/path?field=value#anchor"),
        ("encodeURIComponent", "/?:@&=+$#", "%2F%3F%3A%40%26%3D%2B%24%23"),
        ("encodeURI", "\u{e9}\u{1f600}", "%C3%A9%F0%9F%98%80"),
        ("encodeURIComponent", "\u{e9}\u{1f600}", "%C3%A9%F0%9F%98%80"),
        ("encodeURIComponent", "", ""),
    ] {
        assert_eq!(invoke(&mut vm, name, &[text(input)]), text(expected));
    }
    println!("Registered string capture and URI adapter checks passed");
}
