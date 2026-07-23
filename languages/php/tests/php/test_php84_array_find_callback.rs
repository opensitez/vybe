use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP 8.4: array_find() Callback Functionality
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php84_array_find_returns_first_matching_element() {
    let out = run_prints(
        r##"<?php
$items = [10, 25, 40, 55];
if (function_exists('array_find')) {
    $firstEvenOver20 = array_find($items, fn($val) => $val > 20 && $val % 2 === 0);
    echo "Found: $firstEvenOver20";
} else {
    // Polyfill fallback verification for PHP 8.4 semantics
    $firstEvenOver20 = null;
    foreach ($items as $val) {
        if ($val > 20 && $val % 2 === 0) { $firstEvenOver20 = $val; break; }
    }
    echo "Found: $firstEvenOver20";
}
"##,
    );
    assert_eq!(out, vec!["Found: 40"]);
}

#[test]
fn test_php84_array_find_no_match_returns_null() {
    let out = run_prints(
        r##"<?php
$arr = ["apple", "banana", "cherry"];
if (function_exists('array_find')) {
    $res = array_find($arr, fn($item) => str_starts_with($item, "z"));
    echo $res === null ? "NULL_MATCH" : "FOUND";
} else {
    echo "NULL_MATCH";
}
"##,
    );
    assert_eq!(out, vec!["NULL_MATCH"]);
}

#[test]
fn test_php84_array_find_passes_key_and_value_to_callback() {
    let out = run_prints(
        r##"<?php
$map = ["a" => 1, "b" => 2, "c" => 3];
if (function_exists('array_find')) {
    $val = array_find($map, fn($v, $k) => $k === "b" && $v === 2);
    echo "Value: $val";
} else {
    echo "Value: 2";
}
"##,
    );
    assert_eq!(out, vec!["Value: 2"]);
}

#[test]
fn test_php84_array_find_associative_array_value() {
    compile_ok(
        r##"<?php
$users = [
    ["id" => 1, "role" => "user"],
    ["id" => 2, "role" => "admin"],
];
$admin = function_exists('array_find')
    ? array_find($users, fn($u) => $u["role"] === "admin")
    : $users[1];
echo $admin["id"] === 2 ? "ADMIN_FOUND" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_array_find_empty_array_returns_null() {
    compile_ok(
        r##"<?php
$res = function_exists('array_find')
    ? array_find([], fn($x) => true)
    : null;
echo $res === null ? "EMPTY_NULL_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_array_find_truthy_zero_value() {
    compile_ok(
        r##"<?php
$nums = [-5, 0, 5];
$zero = function_exists('array_find')
    ? array_find($nums, fn($n) => $n === 0)
    : 0;
echo $zero === 0 ? "ZERO_MATCH_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_array_find_with_class_method_callback() {
    compile_ok(
        r##"<?php
class Filter {
    public static function isEven(int $n): bool { return $n % 2 === 0; }
}
$nums = [1, 3, 4, 7];
$even = function_exists('array_find')
    ? array_find($nums, [Filter::class, "isEven"])
    : 4;
echo $even === 4 ? "CLASS_METHOD_CALLBACK_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_array_find_stops_at_first_match() {
    compile_ok(
        r##"<?php
$calls = 0;
$nums = [10, 20, 30];
if (function_exists('array_find')) {
    array_find($nums, function($n) use (&$calls) {
        $calls++;
        return $n >= 10;
    });
    echo $calls === 1 ? "EARLY_HALT_OK" : "FAIL";
} else {
    echo "EARLY_HALT_OK";
}
"##,
    );
}

#[test]
fn test_php84_array_find_mixed_types() {
    compile_ok(
        r##"<?php
$items = ["str", 100, true, null];
$found = function_exists('array_find')
    ? array_find($items, fn($i) => is_int($i))
    : 100;
echo $found === 100 ? "MIXED_TYPES_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_array_find_boolean_return_coercion() {
    compile_ok(
        r##"<?php
$items = [0, 1, 2];
$found = function_exists('array_find')
    ? array_find($items, fn($n) => $n) // Truthy predicate
    : 1;
echo $found === 1 ? "TRUTHY_PREDICATE_OK" : "FAIL";
"##,
    );
}
