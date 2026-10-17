//! Distinguish builtin output assignment from user-function reference passing.
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
fn empty_parse_str_populates_a_fresh_output_variable() {
    assert_eq!(
        run(r#"<?php parse_str('', $result); return is_array($result) ? 'array' : 'missing';"#),
        Value::String("array".into())
    );
}

#[test]
fn empty_parse_str_updates_a_user_function_reference_parameter() {
    assert_eq!(
        run(r#"<?php
            function parse_wrapper($input, &$output) { parse_str($input, $output); }
            $result = null;
            parse_wrapper('', $result);
            return is_array($result) ? 'array' : 'missing';
        "#),
        Value::String("array".into())
    );
}

#[test]
fn parsed_output_reaches_the_callers_array_merge() {
    assert_eq!(
        run(r#"<?php
            function parse_wrapper($input, &$output) { parse_str($input, $output); }
            function options($input, $defaults) {
                parse_wrapper($input, $parsed);
                return array_merge($defaults, $parsed);
            }
            $result = options('', ['optimize' => false]);
            return $result['optimize'] ? 'bad' : 'ok';
        "#),
        Value::String("ok".into())
    );
}
