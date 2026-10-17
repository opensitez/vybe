//! PHP scalar/array serialization wire-format regressions.
use vybe_runtime::{VM, Value};

fn run(source: &str) -> Value {
    static INIT: std::sync::Once = std::sync::Once::new();
    INIT.call_once(vybe_language_php::register);
    let ast = vybe_language_php::parse(source).unwrap();
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
    let chunks = vybe_compiler::Compiler::with_profile(profile).compile(&ast).unwrap();
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    vm.run(chunks).unwrap()
}

#[test]
fn nested_php_arrays_use_php_wire_format() {
    assert_eq!(
        run(r#"<?php return serialize(['args'=>['x',3], 'schedule'=>'hourly']);"#),
        Value::String(r#"a:2:{s:4:"args";a:2:{i:0;s:1:"x";i:1;i:3;}s:8:"schedule";s:6:"hourly";}"#.into())
    );
}

#[test]
fn cron_array_keeps_numeric_keys_and_nested_empty_arrays() {
    assert_eq!(
        run(r#"<?php return serialize([1710000000=>['hook'=>['args'=>[], 'interval'=>3600]]]);"#),
        Value::String(r#"a:1:{i:1710000000;a:1:{s:4:"hook";a:2:{s:4:"args";a:0:{}s:8:"interval";i:3600;}}}"#.into())
    );
}

#[test]
fn primitive_values_use_php_type_tags() {
    assert_eq!(
        run(r#"<?php return serialize(null).'|'.serialize(false).'|'.serialize(1.5).'|'.serialize('abc');"#),
        Value::String(r#"N;|b:0;|d:1.5;|s:3:"abc";"#.into())
    );
}

#[test]
fn integral_float_keeps_its_php_float_type() {
    assert_eq!(
        run(r#"<?php return serialize([1, 1.0]);"#),
        Value::String(r#"a:2:{i:0;i:1;i:1;d:1;}"#.into())
    );
}
