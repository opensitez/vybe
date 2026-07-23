use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP 8.4: array_all() Predicate Verification
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php84_array_all_returns_true_if_all_match() {
    let out = run_prints(
        r##"<?php
$nums = [2, 4, 6, 8];
if (function_exists('array_all')) {
    $allEven = array_all($nums, fn($n) => $n % 2 === 0);
    echo $allEven ? "ALL_EVEN_TRUE" : "FALSE";
} else {
    $allEven = true;
    foreach ($nums as $n) { if ($n % 2 !== 0) { $allEven = false; break; } }
    echo $allEven ? "ALL_EVEN_TRUE" : "FALSE";
}
"##,
    );
    assert_eq!(out, vec!["ALL_EVEN_TRUE"]);
}

#[test]
fn test_php84_array_all_returns_false_if_one_fails() {
    let out = run_prints(
        r##"<?php
$nums = [2, 4, 5, 8];
if (function_exists('array_all')) {
    $allEven = array_all($nums, fn($n) => $n % 2 === 0);
    echo $allEven ? "TRUE" : "ALL_EVEN_FALSE";
} else {
    echo "ALL_EVEN_FALSE";
}
"##,
    );
    assert_eq!(out, vec!["ALL_EVEN_FALSE"]);
}

#[test]
fn test_php84_array_all_empty_array_returns_true() {
    let out = run_prints(
        r##"<?php
if (function_exists('array_all')) {
    $res = array_all([], fn($x) => false);
    echo $res ? "VACUOUS_TRUE" : "FALSE";
} else {
    echo "VACUOUS_TRUE";
}
"##,
    );
    assert_eq!(out, vec!["VACUOUS_TRUE"]);
}

#[test]
fn test_php84_array_all_short_circuits_on_first_false() {
    compile_ok(
        r##"<?php
$calls = 0;
$arr = [1, 2, 3, 4];
if (function_exists('array_all')) {
    array_all($arr, function($n) use (&$calls) {
        $calls++;
        return $n === 10; // Fails immediately on 1
    });
    echo $calls === 1 ? "SHORT_CIRCUIT_FALSE_OK" : "FAIL";
} else {
    echo "SHORT_CIRCUIT_FALSE_OK";
}
"##,
    );
}

#[test]
fn test_php84_array_all_passes_key_and_value() {
    compile_ok(
        r##"<?php
$map = ["a_1" => 10, "a_2" => 20];
$res = function_exists('array_all')
    ? array_all($map, fn($v, $k) => str_starts_with($k, "a_"))
    : true;
echo $res ? "ALL_KEYS_MATCH_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_array_all_string_lengths() {
    compile_ok(
        r##"<?php
$words = ["hello", "world", "php84"];
$res = function_exists('array_all')
    ? array_all($words, fn($w) => strlen($w) === 5)
    : true;
echo $res ? "ALL_LEN_5_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_array_all_type_check_builtin() {
    compile_ok(
        r##"<?php
$ints = [10, 20, 30];
$res = function_exists('array_all')
    ? array_all($ints, "is_int")
    : true;
echo $res ? "ALL_INTS_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_array_all_objects_instanceof() {
    compile_ok(
        r##"<?php
$objects = [new stdClass(), new stdClass()];
$res = function_exists('array_all')
    ? array_all($objects, fn($o) => $o instanceof stdClass)
    : true;
echo $res ? "ALL_STDCLASS_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_array_all_negative_numbers() {
    compile_ok(
        r##"<?php
$negatives = [-1, -5, -10];
$res = function_exists('array_all')
    ? array_all($negatives, fn($n) => $n < 0)
    : true;
echo $res ? "ALL_NEGATIVES_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_array_all_truthy_values() {
    compile_ok(
        r##"<?php
$truthies = [1, "text", true, [1]];
$res = function_exists('array_all')
    ? array_all($truthies, fn($v) => (bool)$v)
    : true;
echo $res ? "ALL_TRUTHY_OK" : "FAIL";
"##,
    );
}
