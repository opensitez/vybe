//! Missing destructors skip finalization without changing real destructor order.
use std::sync::Arc;
use vybe_compiler::Compiler;
use vybe_runtime::{VM, Value};

#[test]
fn scalar_rebinds_and_object_aliases_preserve_destructor_order() {
    let mut setup = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut setup);
    vybe_language_php::register();
    let language = vybe_compiler::languages::find_by_name("php").unwrap();
    let profile = Arc::new(
        vybe_runtime::profile::parse_profile((language.profile_source)()).unwrap(),
    );
    let module = vybe_language_php::parse(r#"<?php
$GLOBALS['events'] = [];
class RebindProbe {
    public function __destruct() { $GLOBALS['events'][] = 'D'; }
}
function rebind_probe() {
    $number = 0;
    for ($i = 0; $i < 20; ++$i) { $number = $i; }
    $first = new RebindProbe();
    $alias = $first;
    $first = 0;
    $GLOBALS['events'][] = 'A';
    $alias = null;
    $GLOBALS['events'][] = 'B';
    $after = 7;
    return $number + $after;
}
$result = rebind_probe();
return json_encode([$result, $GLOBALS['events']]);
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
    assert_eq!(actual.as_ref(), r#"[26,["A","D","B"]]"#);
}
