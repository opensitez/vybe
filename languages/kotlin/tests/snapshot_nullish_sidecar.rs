use vybe_compiler::primitives::platforms;
use vybe_runtime::{VM, Value};
fn compiled() -> Vec<vybe_runtime::Chunk> {
    vybe_language_kotlin::register();
    vybe_platform_jvm::register();
    let ast =
        vybe_language_kotlin::parse(include_str!("fixtures/snapshot_nullish_sidecar.kt")).unwrap();
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_kotlin::profile_source()).unwrap();
    vybe_compiler::Compiler::with_profile(profile)
        .compile(&ast)
        .unwrap()
}
#[test]
fn actual_source_fresh_snapshot_set_preserves_duplicate_and_remove_results() {
    let chunks = compiled();
    let arity = chunks
        .iter()
        .find(|c| c.name == "kotlin_controls")
        .unwrap()
        .arity;
    assert!(arity <= 1, "unexpected zero-user-argument callable ABI");
    let mut v = VM::new();
    platforms::register_platforms_all(&mut v);
    v.run(chunks).unwrap();
    let f = v.global("kotlin_controls").unwrap().clone();
    let args = if arity == 1 {
        vec![Value::Null]
    } else {
        vec![]
    };
    for _ in 0..3 {
        let result = v.invoke(&f, &args).unwrap();
        assert_eq!(result.as_i32(), 58);
        assert!(v.last_exception.is_none());
    }
}
