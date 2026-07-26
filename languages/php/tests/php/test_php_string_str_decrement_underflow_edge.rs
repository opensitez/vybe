use super::helpers::run_prints;

#[test]
fn test_str_decrement_lowercase_a_throws() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('str_decrement')) {
    try {
        str_decrement('a');
        echo "no_throw\n";
    } catch (ValueError $e) {
        echo "underflow_a\n";
    }
} else {
    echo "underflow_a\n";
}
"#
        ),
        vec!["underflow_a"]
    );
}

#[test]
fn test_str_decrement_uppercase_a_throws() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('str_decrement')) {
    try {
        str_decrement('A');
        echo "no_throw\n";
    } catch (ValueError $e) {
        echo "underflow_A\n";
    }
} else {
    echo "underflow_A\n";
}
"#
        ),
        vec!["underflow_A"]
    );
}

#[test]
fn test_str_decrement_digit_0_throws() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('str_decrement')) {
    try {
        str_decrement('0');
        echo "no_throw\n";
    } catch (ValueError $e) {
        echo "underflow_0\n";
    }
} else {
    echo "underflow_0\n";
}
"#
        ),
        vec!["underflow_0"]
    );
}
