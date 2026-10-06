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
macro_rules! control {
    ($name:ident, $expression:literal) => {
        #[test]
        fn $name() {
            check($expression);
        }
    };
}
control!(numeric_lt, "1 < 2");
control!(numeric_gt, "2 > 1");
control!(numeric_le, "2 <= 2");
control!(numeric_ge, "2 >= 2");
control!(string_lt, "'a' < 'b'");
control!(list_lt, "[1] < [1, 2]");
control!(tuple_lt, "(1, 2) < (1, 3)");
control!(subset_equal, "{1} <= {1}");
control!(superset_equal, "{1} >= {1}");
control!(date_lt, "date(2020, 1, 1) < date(2021, 1, 1)");
control!(timedelta_lt, "timedelta(days=1) < timedelta(days=2)");
control!(operator_lt, "operator.lt(1, 2)");
