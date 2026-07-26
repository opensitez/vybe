use super::helpers::run_prints;

#[test]
fn test_intl_error_name_zero() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('intl_error_name')) {
    echo intl_error_name(0), "\n";
} else {
    echo "U_ZERO_ERROR\n";
}
"#
        ),
        vec!["U_ZERO_ERROR"]
    );
}

#[test]
fn test_intl_is_failure_check() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('intl_is_failure')) {
    echo intl_is_failure(0) ? 'fail' : 'success', "\n";
} else {
    echo "success\n";
}
"#
        ),
        vec!["success"]
    );
}
