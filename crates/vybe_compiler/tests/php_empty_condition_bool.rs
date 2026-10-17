use std::sync::{Arc, Mutex};
use vybe_runtime::{
    VM, Value,
    value::{Object, ObjectKind},
};
#[test]
fn actual_php_empty_value_condition_and_negation_are_boolean_checked() {
    vybe_language_php::register();
    let module=vybe_language_php::parse("<?php function empty_controls($x){$v=empty($x);$n=!empty($x);if(empty($x))$c=1;else$c=2;return [$v,$n,$c];}").unwrap();
    let chunks = vybe_compiler::Compiler::with_profile(
        vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap(),
    )
    .compile(&module)
    .unwrap();
    let bytes = vybe_platform_wasm::write_wasm(&chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .unwrap();
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    vm.run(chunks).unwrap();
    let f = vm.global("__vybe_func$empty_controls").unwrap().clone();
    for (input, empty) in [
        (Value::F64(0.0), true),
        (Value::String("0".into()), true),
        (Value::Null, true),
        (
            Value::Object(Arc::new(Mutex::new(Object::new_array(vec![Value::F64(
                1.0,
            )])))),
            false,
        ),
        (Value::Object(Arc::new(Mutex::new(Object::new()))), false),
    ] {
        let result = vm.invoke_callback(&f, &[input]);
        assert!(vm.last_exception.is_none());
        let Value::Object(o) = result else {
            panic!("expected result Array")
        };
        let o = o.lock().unwrap();
        let ObjectKind::Array(a) = &o.kind else {
            panic!("expected Array kind")
        };
        assert_eq!(a[0], Value::Bool(empty));
        assert_eq!(a[1], Value::Bool(!empty));
        assert_eq!(a[2].as_f64(), if empty { 1.0 } else { 2.0 });
    }
}
