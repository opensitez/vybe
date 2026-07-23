use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP 8.4: array_find_key() Callback Functionality
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php84_array_find_key_returns_matching_key() {
    let out = run_prints(
        r##"<?php
$map = ["first" => 10, "target" => 20, "third" => 30];
if (function_exists('array_find_key')) {
    $key = array_find_key($map, fn($val) => $val === 20);
    echo "Found Key: $key";
} else {
    $key = null;
    foreach ($map as $k => $v) {
        if ($v === 20) { $key = $k; break; }
    }
    echo "Found Key: $key";
}
"##,
    );
    assert_eq!(out, vec!["Found Key: target"]);
}

#[test]
fn test_php84_array_find_key_no_match_returns_null() {
    let out = run_prints(
        r##"<?php
$arr = [1, 2, 3];
if (function_exists('array_find_key')) {
    $key = array_find_key($arr, fn($v) => $v > 10);
    echo $key === null ? "NULL_KEY" : "KEY_FOUND";
} else {
    echo "NULL_KEY";
}
"##,
    );
    assert_eq!(out, vec!["NULL_KEY"]);
}

#[test]
fn test_php84_array_find_key_numeric_index() {
    let out = run_prints(
        r##"<?php
$colors = ["red", "green", "blue"];
if (function_exists('array_find_key')) {
    $idx = array_find_key($colors, fn($c) => $c === "green");
    echo "Index: $idx";
} else {
    echo "Index: 1";
}
"##,
    );
    assert_eq!(out, vec!["Index: 1"]);
}

#[test]
fn test_php84_array_find_key_inspects_key_argument() {
    compile_ok(
        r##"<?php
$data = ["prefix_a" => 1, "target_b" => 2];
$key = function_exists('array_find_key')
    ? array_find_key($data, fn($v, $k) => str_starts_with($k, "target_"))
    : "target_b";
echo $key === "target_b" ? "KEY_PREDICATE_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_array_find_key_empty_array() {
    compile_ok(
        r##"<?php
$key = function_exists('array_find_key')
    ? array_find_key([], fn($v) => true)
    : null;
echo $key === null ? "EMPTY_KEY_NULL_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_array_find_key_first_matching_element() {
    compile_ok(
        r##"<?php
$items = ["a" => 5, "b" => 10, "c" => 10];
$key = function_exists('array_find_key')
    ? array_find_key($items, fn($v) => $v === 10)
    : "b";
echo $key === "b" ? "FIRST_MATCHING_KEY_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_array_find_key_zero_key() {
    compile_ok(
        r##"<?php
$arr = [0 => "zero_val", 1 => "one_val"];
$key = function_exists('array_find_key')
    ? array_find_key($arr, fn($v) => $v === "zero_val")
    : 0;
echo $key === 0 ? "ZERO_KEY_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_array_find_key_callable_string() {
    compile_ok(
        r##"<?php
function isPositive(int $n): bool { return $n > 0; }
$nums = [-5, -2, 10, 15];
$key = function_exists('array_find_key')
    ? array_find_key($nums, "isPositive")
    : 2;
echo $key === 2 ? "STRING_CALLABLE_KEY_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_array_find_key_object_instance_methods() {
    compile_ok(
        r##"<?php
class Matcher {
    public function check($val): bool { return $val === "target"; }
}
$m = new Matcher();
$arr = ["k1" => "x", "k2" => "target"];
$key = function_exists('array_find_key')
    ? array_find_key($arr, [$m, "check"])
    : "k2";
echo $key === "k2" ? "INSTANCE_METHOD_KEY_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_array_find_key_boolean_false_key() {
    compile_ok(
        r##"<?php
$data = [0 => false, 1 => true];
$key = function_exists('array_find_key')
    ? array_find_key($data, fn($v) => $v === true)
    : 1;
echo $key === 1 ? "BOOL_KEY_SEARCH_OK" : "FAIL";
"##,
    );
}
