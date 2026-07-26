use super::helpers::run_prints;

#[test]
fn test_interface_exists_declared_interface() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Contract {}
echo interface_exists('Contract', false) ? 'interface_found' : 'err', "\n";
"#
        ),
        vec!["interface_found"]
    );
}

#[test]
fn test_interface_exists_missing_no_autoload() {
    assert_eq!(
        run_prints(
            r#"<?php
echo interface_exists('NonExistentContract', false) ? 'found' : 'missing_no_autoload', "\n";
"#
        ),
        vec!["missing_no_autoload"]
    );
}
