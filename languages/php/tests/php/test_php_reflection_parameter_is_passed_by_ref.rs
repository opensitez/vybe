use super::helpers::run_prints;

#[test]
fn test_reflection_parameter_is_passed_by_reference() {
    assert_eq!(
        run_prints(
            r#"<?php
function modify_val(&$arg) { $arg = 100; }
$rf = new ReflectionFunction('modify_val');
$param = $rf->getParameters()[0];
echo $param->isPassedByReference() ? 'by_ref' : 'by_val', "\n";
"#
        ),
        vec!["by_ref"]
    );
}

#[test]
fn test_reflection_parameter_is_variadic() {
    assert_eq!(
        run_prints(
            r#"<?php
function collect_items(...$items) {}
$rf = new ReflectionFunction('collect_items');
$param = $rf->getParameters()[0];
echo $param->isVariadic() ? 'variadic' : 'fixed', "\n";
"#
        ),
        vec!["variadic"]
    );
}
