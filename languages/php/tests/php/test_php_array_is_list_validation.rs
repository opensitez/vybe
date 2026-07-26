use super::helpers::run_prints;

#[test]
fn test_array_is_list_empty() {
    assert_eq!(
        run_prints(
            r#"<?php
echo array_is_list([]) ? 'true' : 'false', "\n";
"#
        ),
        vec!["true"]
    );
}

#[test]
fn test_array_is_list_sequential_integer_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
echo array_is_list(['a', 'b', 'c']) ? 'true' : 'false', "\n";
"#
        ),
        vec!["true"]
    );
}

#[test]
fn test_array_is_list_out_of_order_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$arr = [0 => 'a', 2 => 'b', 1 => 'c'];
echo array_is_list($arr) ? 'true' : 'false', "\n";
"#
        ),
        vec!["false"]
    );
}

#[test]
fn test_array_is_list_string_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$arr = ['name' => 'Alice', 'age' => 30];
echo array_is_list($arr) ? 'true' : 'false', "\n";
"#
        ),
        vec!["false"]
    );
}

#[test]
fn test_array_is_list_non_zero_start() {
    assert_eq!(
        run_prints(
            r#"<?php
$arr = [1 => 'a', 2 => 'b'];
echo array_is_list($arr) ? 'true' : 'false', "\n";
"#
        ),
        vec!["false"]
    );
}

#[test]
fn test_array_is_list_with_sparse_holes() {
    assert_eq!(
        run_prints(
            r#"<?php
$arr = [];
$arr[] = 'a';
$arr[2] = 'b';
echo array_is_list($arr) ? 'true' : 'false';
"#
        ),
        vec!["false"]
    );
}

#[test]
fn test_array_is_list_with_explicit_int_zero() {
    assert_eq!(
        run_prints(
            r#"<?php
$arr = [0 => 'a'];
echo array_is_list($arr) ? 'true' : 'false';
"#
        ),
        vec!["true"]
    );
}

#[test]
fn test_array_is_list_string_numeric_keys_are_not_list() {
    assert_eq!(
        run_prints(
            r#"<?php
$arr = ['0' => 'a', '1' => 'b'];
echo array_is_list($arr) ? 'true' : 'false';
"#
        ),
        vec!["false"]
    );
}

#[test]
fn test_array_is_list_with_bool_zero_false_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$arr = [false => 'a', true => 'b'];
echo array_is_list($arr) ? 'true' : 'false', '|', count($arr);
"#
        ),
        vec!["false|2"]
    );
}
