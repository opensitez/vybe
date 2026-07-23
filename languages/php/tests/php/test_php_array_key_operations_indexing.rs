use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Array Keys & Internal Pointer Operations — array_keys, array_values, array_flip, array_reverse, array_key_first, array_key_last, pointer functions
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_array_key_first_and_key_last_php73() {
    let out = run_prints(
        r#"<?php
$data = ["a" => 10, "b" => 20, "c" => 30];
echo array_key_first($data) . " | " . array_key_last($data);
"#,
    );
    assert_eq!(out, vec!["a | c"]);
}

#[test]
fn test_php_array_flip_swapping_keys_and_values() {
    let out = run_prints(
        r#"<?php
$input = ["oranges", "apples", "pears"];
$flipped = array_flip($input);
echo "oranges=" . $flipped["oranges"] . " pears=" . $flipped["pears"];
"#,
    );
    assert_eq!(out, vec!["oranges=0 pears=2"]);
}

#[test]
fn test_php_array_reverse_preserving_keys() {
    let out = run_prints(
        r#"<?php
$input = ["a" => 1, "b" => 2, "c" => 3];
$reversed = array_reverse($input, preserve_keys: true);
echo implode(",", array_keys($reversed));
"#,
    );
    assert_eq!(out, vec!["c,b,a"]);
}

#[test]
fn test_php_array_internal_pointer_functions() {
    let out = run_prints(
        r#"<?php
$transport = ["foot", "bike", "car", "plane"];
$mode = current($transport);
$next = next($transport);
$end = end($transport);
$prev = prev($transport);
$reset = reset($transport);

echo "$mode | $next | $end | $prev | $reset";
"#,
    );
    assert_eq!(out, vec!["foot | bike | plane | car | foot"]);
}

#[test]
fn test_php_array_change_key_case_upper_lower() {
    compile_ok(
        r#"<?php
$input = ["First" => 1, "SecOND" => 4];
$lower = array_change_key_case($input, CASE_LOWER);
$upper = array_change_key_case($input, CASE_UPPER);
echo implode(",", array_keys($lower)) . " | " . implode(",", array_keys($upper));
"#,
    );
}

#[test]
fn test_php_array_keys_value_filtering() {
    compile_ok(
        r#"<?php
$array = ["blue", "red", "green", "blue", "blue"];
$blueKeys = array_keys($array, "blue", strict: true);
echo implode(",", $blueKeys);
"#,
    );
}

#[test]
fn test_php_key_exists_and_array_key_exists() {
    compile_ok(
        r#"<?php
$arr = ["name" => "Alice", "null_val" => null];
echo array_key_exists("null_val", $arr) ? "KEY_EXISTS" : "NO";
"#,
    );
}

#[test]
fn test_php_array_values_reindexing() {
    compile_ok(
        r#"<?php
$arr = [10 => "a", 20 => "b", 30 => "c"];
$reindexed = array_values($arr);
echo $reindexed[0] . "-" . $reindexed[1];
"#,
    );
}

#[test]
fn test_php_array_fill_zero_based_index() {
    compile_ok(
        r#"<?php
$a = array_fill(5, 3, "banana");
echo implode(",", array_keys($a));
"#,
    );
}

#[test]
fn test_php_array_is_list_check_php81() {
    compile_ok(
        r#"<?php
if (function_exists('array_is_list')) {
    echo array_is_list(["a", "b", "c"]) ? "LIST" : "ASSOC";
    echo array_is_list(["a" => 1]) ? "LIST" : "ASSOC";
} else {
    echo "LISTASSOC";
}
"#,
    );
}
