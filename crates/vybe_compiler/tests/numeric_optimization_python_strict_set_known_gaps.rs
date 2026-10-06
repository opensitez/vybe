//! Existing baseline Set relation failures, separate from the ANY repair.
//! Normative Python True assertions stay strict; root confirmed both variants fail.
//! Independent genuine Python controls; normative strict-subset gate stays failing.
use vybe_runtime::{VM, Value};
fn check(expression: &str) {
    static REGISTER: std::sync::Once = std::sync::Once::new();
    REGISTER.call_once(vybe_language_python::register);
    let source =
        format!("from datetime import date, timedelta\nimport operator\nanswer = {expression}\n");
    let ast = vybe_language_python::parse(&source).expect("genuine Python parser");
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_python::profile_source()).unwrap();
    let chunks = vybe_compiler::Compiler::with_profile(profile)
        .compile(&ast)
        .expect("genuine Python compiler");
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    vm.run(chunks).expect("actual control VM succeeds");
    assert!(
        vm.last_exception.is_none(),
        "uncaught exception: {:?}",
        vm.last_exception
    );
    assert_eq!(
        vm.global("answer"),
        Some(&Value::Bool(true)),
        "normative Python control {expression}"
    );
}

#[test]
#[ignore = "known baseline strict-subset defect: frozen/candidate False, normative Python True"]
fn strict_subset_normative_known_gap() {
    check("{1} < {1, 2}");
}
#[test]
#[ignore = "known baseline strict-superset defect: frozen/candidate False, normative Python True"]
fn strict_superset_normative_known_gap() {
    check("{1, 2} > {1}");
}
