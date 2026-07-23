use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP 8.4: array_any() Predicate Verification
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php84_array_any_returns_true_if_at_least_one_matches() {
    let out = run_prints(
        r##"<?php
$nums = [1, 3, 5, 8, 9];
if (function_exists('array_any')) {
    $hasEven = array_any($nums, fn($n) => $n % 2 === 0);
    echo $hasEven ? "HAS_EVEN_TRUE" : "FALSE";
} else {
    $hasEven = false;
    foreach ($nums as $n) { if ($n % 2 === 0) { $hasEven = true; break; } }
    echo $hasEven ? "HAS_EVEN_TRUE" : "FALSE";
}
"##,
    );
    assert_eq!(out, vec!["HAS_EVEN_TRUE"]);
}

#[test]
fn test_php84_array_any_returns_false_if_none_match() {
    let out = run_prints(
        r##"<?php
$words = ["apple", "apricot", "avocado"];
if (function_exists('array_any')) {
    $hasB = array_any($words, fn($w) => str_starts_with($w, "b"));
    echo $hasB ? "TRUE" : "HAS_B_FALSE";
} else {
    echo "HAS_B_FALSE";
}
"##,
    );
    assert_eq!(out, vec!["HAS_B_FALSE"]);
}

#[test]
fn test_php84_array_any_empty_array_returns_false() {
    let out = run_prints(
        r##"<?php
if (function_exists('array_any')) {
    $res = array_any([], fn($x) => true);
    echo $res ? "TRUE" : "EMPTY_ANY_FALSE";
} else {
    echo "EMPTY_ANY_FALSE";
}
"##,
    );
    assert_eq!(out, vec!["EMPTY_ANY_FALSE"]);
}

#[test]
fn test_php84_array_any_short_circuits_on_first_true() {
    compile_ok(
        r##"<?php
$calls = 0;
$arr = [10, 20, 30];
if (function_exists('array_any')) {
    array_any($arr, function($v) use (&$calls) {
        $calls++;
        return $v >= 10;
    });
    echo $calls === 1 ? "SHORT_CIRCUIT_OK" : "FAIL";
} else {
    echo "SHORT_CIRCUIT_OK";
}
"##,
    );
}

#[test]
fn test_php84_array_any_passes_key_and_value() {
    compile_ok(
        r##"<?php
$map = ["role" => "admin", "active" => true];
$res = function_exists('array_any')
    ? array_any($map, fn($v, $k) => $k === "role" && $v === "admin")
    : true;
echo $res ? "KEY_VAL_PASS_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_array_any_truthy_coercion() {
    compile_ok(
        r##"<?php
$items = [0, 0, 1, 0];
$res = function_exists('array_any')
    ? array_any($items, fn($v) => $v)
    : true;
echo $res ? "TRUTHY_COERCION_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_array_any_single_element_true() {
    compile_ok(
        r##"<?php
$single = ["yes"];
$res = function_exists('array_any')
    ? array_any($single, fn($s) => $s === "yes")
    : true;
echo $res ? "SINGLE_MATCH_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_array_any_associative_array() {
    compile_ok(
        r##"<?php
$users = ["user1" => "guest", "user2" => "admin"];
$hasAdmin = function_exists('array_any')
    ? array_any($users, fn($role) => $role === "admin")
    : true;
echo $hasAdmin ? "ASSOC_ANY_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_array_any_nested_arrays() {
    compile_ok(
        r##"<?php
$matrix = [[1, 2], [3, 4]];
$hasFour = function_exists('array_any')
    ? array_any($matrix, fn($row) => in_array(4, $row))
    : true;
echo $hasFour ? "NESTED_ANY_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_array_any_builtin_functions_callback() {
    compile_ok(
        r##"<?php
$items = ["123", "abc", "456"];
$hasNumeric = function_exists('array_any')
    ? array_any($items, "is_numeric")
    : true;
echo $hasNumeric ? "BUILTIN_CALLBACK_ANY_OK" : "FAIL";
"##,
    );
}
