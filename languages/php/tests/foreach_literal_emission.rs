//! Literal arrays do not require runtime Iterator protocol dispatch.
use std::sync::Arc;
use vybe_compiler::Compiler;
use vybe_runtime::{VM, Value};

#[test]
fn literal_foreach_keeps_keys_evaluation_order_and_continue() {
    let mut setup = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut setup);
    vybe_language_php::register();
    let language = vybe_compiler::languages::find_by_name("php").unwrap();
    let profile = Arc::new(
        vybe_runtime::profile::parse_profile((language.profile_source)()).unwrap(),
    );
    let module = vybe_language_php::parse(r#"<?php
$counter = 0;
$rows = [];
foreach (['first' => ++$counter, 'skip' => ++$counter, 'last' => ++$counter] as $key => $value) {
    if ($key === 'skip') { continue; }
    $rows[] = $key . ':' . $value;
}
$total = 0;
foreach ([1, 2, 3, 4] as $value) { $total += $value; }
return json_encode([$counter, $rows, $total]);
"#).unwrap();
    let chunks = Compiler::with_shared_profile(profile)
        .with_run_module_end_hooks(false)
        .compile(&module).unwrap();
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    let result = vm.run(chunks).unwrap();
    let Value::String(actual) = result else {
        panic!("expected JSON, got {result:?}");
    };
    assert_eq!(actual.as_ref(), r#"[3,["first:1","last:3"],10]"#);
}
