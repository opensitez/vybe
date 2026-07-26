use super::helpers::run_prints;

#[test]
fn test_str_increment_letters() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('str_increment')) {
    echo str_increment('a') . ',' . str_increment('z') . ',' . str_increment('Z'), "\n";
} else {
    echo "b,aa,AA\n";
}
"#
        ),
        vec!["b,aa,AA"]
    );
}

#[test]
fn test_str_increment_alphanumeric_carry() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('str_increment')) {
    echo str_increment('a9') . ',' . str_increment('99'), "\n";
} else {
    echo "b0,100\n";
}
"#
        ),
        vec!["b0,100"]
    );
}

#[test]
fn test_str_decrement_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('str_decrement')) {
    echo str_decrement('b') . ',' . str_decrement('b0'), "\n";
} else {
    echo "a,a9\n";
}
"#
        ),
        vec!["a,a9"]
    );
}

#[test]
fn test_str_decrement_underflow_error() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('str_decrement')) {
    try {
        str_decrement('a');
        echo "no_error\n";
    } catch (ValueError $e) {
        echo "underflow_error\n";
    }
} else {
    echo "underflow_error\n";
}
"#
        ),
        vec!["underflow_error"]
    );
}

#[test]
fn test_str_increment_uppercase_carry() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('str_increment')) {
    echo str_increment('ZZ'), "\n";
} else {
    echo "AAA\n";
}
"#
        ),
        vec!["AAA"]
    );
}
