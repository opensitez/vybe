use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Array Functions: array_diff, array_intersect, assoc & key variations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_array_diff_values_only() {
    let out = run_prints(
        r##"<?php
$a1 = ["a" => "green", "red", "blue", "red"];
$a2 = ["b" => "green", "yellow", "red"];
$diff = array_diff($a1, $a2);
echo implode(",", $diff);
"##,
    );
    assert_eq!(out, vec!["blue"]);
}

#[test]
fn test_php_array_diff_assoc_keys_and_values() {
    let out = run_prints(
        r##"<?php
$a1 = ["a" => "green", "b" => "brown", "c" => "blue", "red"];
$a2 = ["a" => "green", "yellow", "red"];
$diff = array_diff_assoc($a1, $a2);
echo implode(",", array_keys($diff));
"##,
    );
    assert_eq!(out, vec!["b,c,0"]);
}

#[test]
fn test_php_array_intersect_values_only() {
    let out = run_prints(
        r##"<?php
$a1 = ["a" => "green", "red", "blue"];
$a2 = ["b" => "green", "yellow", "red"];
$intersect = array_intersect($a1, $a2);
echo implode(",", $intersect);
"##,
    );
    assert_eq!(out, vec!["green,red"]);
}

#[test]
fn test_php_array_intersect_key_matching_keys() {
    let out = run_prints(
        r##"<?php
$a1 = ["blue" => 1, "red" => 2, "green" => 3];
$a2 = ["green" => 4, "blue" => 5, "yellow" => 6];
$inter = array_intersect_key($a1, $a2);
echo implode(",", array_keys($inter));
"##,
    );
    assert_eq!(out, vec!["blue,green"]);
}

#[test]
fn test_php_array_diff_ukey_callback() {
    let out = run_prints(
        r##"<?php
$a1 = ["blue" => 1, "red" => 2, "green" => 3];
$a2 = ["blue" => 5, "yellow" => 7];
$diff = array_diff_ukey($a1, $a2, fn($k1, $k2) => strcasecmp($k1, $k2));
echo implode(",", array_keys($diff));
"##,
    );
    assert_eq!(out, vec!["red,green"]);
}

#[test]
fn test_php_array_intersect_assoc_keys_values() {
    compile_ok(
        r##"<?php
$a1 = ["a" => "green", "b" => "brown", "c" => "blue"];
$a2 = ["a" => "green", "b" => "yellow", "e" => "blue"];
$inter = array_intersect_assoc($a1, $a2);
echo count($inter) === 1 && isset($inter["a"]) ? "INTER_ASSOC_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_array_diff_key_compares_keys() {
    compile_ok(
        r##"<?php
$a1 = [10 => "val1", 20 => "val2", 30 => "val3"];
$a2 = [10 => "different", 40 => "val4"];
$diff = array_diff_key($a1, $a2);
echo count($diff) === 2 && isset($diff[20]) && isset($diff[30]) ? "DIFF_KEY_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_array_udiff_value_callback() {
    compile_ok(
        r##"<?php
$a1 = [1.5, 2.5, 3.5];
$a2 = [1.0, 2.0];
$diff = array_udiff($a1, $a2, fn($a, $b) => (int)$a <=> (int)$b);
echo count($diff) === 1 && $diff[2] == 3.5 ? "UDIFF_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_array_uintersect_value_callback() {
    compile_ok(
        r##"<?php
$a1 = ["Apple", "banana"];
$a2 = ["apple", "BANANA"];
$inter = array_uintersect($a1, $a2, fn($a, $b) => strcasecmp($a, $b));
echo count($inter) === 2 ? "UINTERSECT_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_array_diff_uassoc_callback() {
    compile_ok(
        r##"<?php
$a1 = ["a" => "green", "b" => "brown"];
$a2 = ["A" => "green", "b" => "yellow"];
$diff = array_diff_uassoc($a1, $a2, fn($a, $b) => strcasecmp($a, $b));
echo count($diff) === 1 && isset($diff["b"]) ? "DIFF_UASSOC_OK" : "FAIL";
"##,
    );
}
