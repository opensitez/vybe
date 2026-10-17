//! The numeric fast path must preserve PHP predicate values and conditions.

use std::sync::Arc;
use vybe_compiler::Compiler;
use vybe_runtime::{VM, Value};

#[test]
fn typed_numeric_comparisons_return_booleans_and_keep_condition_semantics() {
    let mut setup = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut setup);
    vybe_language_php::register();
    let language = vybe_compiler::languages::find_by_name("php").unwrap();
    let profile = Arc::new(
        vybe_runtime::profile::parse_profile((language.profile_source)()).unwrap(),
    );
    let source = r#"<?php
function lt(float $a, float $b) { return $a < $b; }
function gt(float $a, float $b) { return $a > $b; }
function le(float $a, float $b) { return $a <= $b; }
function ge(float $a, float $b) { return $a >= $b; }
function choose(float $a, float $b) {
    if ($a < $b) { return 7; }
    return 9;
}
function negate($value) { return !$value; }
function choose_negated($value) {
    if (!$value) { return 11; }
    return 13;
}
$a = 1.5;
$b = 2.5;
return json_encode([
    gettype(lt($a, $b)), gettype(gt($a, $b)),
    gettype(le($a, $b)), gettype(ge($a, $b)),
    lt($a, $b), gt($a, $b), le($a, $b), ge($a, $b),
    choose($a, $b), choose($b, $a),
    gettype(negate(0)), negate(0), negate(1),
    choose_negated(0), choose_negated(1)
]);
"#;
    let module = vybe_language_php::parse(source).unwrap();
    let chunks = Compiler::with_shared_profile(profile)
        .with_run_module_end_hooks(false)
        .compile(&module)
        .unwrap();
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    let value = vm.run(chunks).unwrap();
    let Value::String(actual) = value else {
        panic!("expected JSON string, got {value:?}");
    };
    assert_eq!(
        actual.as_ref(),
        r#"["boolean","boolean","boolean","boolean",true,false,true,false,7,9,"boolean",true,false,11,13]"#,
    );
}
