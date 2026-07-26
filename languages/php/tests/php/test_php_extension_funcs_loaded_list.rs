use super::helpers::run_prints;

#[test]
fn test_get_extension_funcs_standard() {
    assert_eq!(
        run_prints(
            r#"<?php
$funcs = get_extension_funcs('standard');
echo (is_array($funcs) && in_array('strlen', $funcs, true)) ? 'standard_funcs_ok' : 'err', "\n";
"#
        ),
        vec!["standard_funcs_ok"]
    );
}

#[test]
fn test_get_extension_funcs_nonexistent() {
    assert_eq!(
        run_prints(
            r#"<?php
$funcs = get_extension_funcs('non_existent_extension_xyz');
echo $funcs === false ? 'false_ok' : 'err', "\n";
"#
        ),
        vec!["false_ok"]
    );
}

#[test]
fn test_get_loaded_extensions_zend() {
    assert_eq!(
        run_prints(
            r#"<?php
$exts = get_loaded_extensions(true);
echo is_array($exts) ? 'zend_exts_ok' : 'err', "\n";
"#
        ),
        vec!["zend_exts_ok"]
    );
}
