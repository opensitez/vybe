use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Array Functions: array_chunk, array_slice, array_splice, array_combine
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_array_chunk_splits_array_evenly() {
    let out = run_prints(
        r##"<?php
$arr = [1, 2, 3, 4, 5];
$chunks = array_chunk($arr, 2);
echo count($chunks) . " " . count($chunks[2]);
"##,
    );
    assert_eq!(out, vec!["3 1"]);
}

#[test]
fn test_php_array_chunk_preserve_keys() {
    let out = run_prints(
        r##"<?php
$arr = ["a" => 10, "b" => 20, "c" => 30];
$chunks = array_chunk($arr, 2, true);
echo isset($chunks[0]["b"]) ? "PRESERVED" : "LOST";
"##,
    );
    assert_eq!(out, vec!["PRESERVED"]);
}

#[test]
fn test_php_array_slice_positive_offset_length() {
    let out = run_prints(
        r##"<?php
$arr = ["a", "b", "c", "d", "e"];
$sliced = array_slice($arr, 1, 3);
echo implode(",", $sliced);
"##,
    );
    assert_eq!(out, vec!["b,c,d"]);
}

#[test]
fn test_php_array_slice_negative_offset() {
    let out = run_prints(
        r##"<?php
$arr = [10, 20, 30, 40, 50];
$sliced = array_slice($arr, -2);
echo implode("-", $sliced);
"##,
    );
    assert_eq!(out, vec!["40-50"]);
}

#[test]
fn test_php_array_splice_remove_and_replace() {
    let out = run_prints(
        r##"<?php
$input = ["red", "green", "blue", "yellow"];
$removed = array_splice($input, 1, 2, ["orange"]);
echo "Input: " . implode(",", $input) . " | Removed: " . implode(",", $removed);
"##,
    );
    assert_eq!(out, vec!["Input: red,orange,yellow | Removed: green,blue"]);
}

#[test]
fn test_php_array_combine_keys_and_values() {
    let out = run_prints(
        r##"<?php
$keys = ["x", "y", "z"];
$values = [1, 2, 3];
$combined = array_combine($keys, $values);
echo $combined["y"];
"##,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_php_array_slice_preserve_keys_flag() {
    compile_ok(
        r##"<?php
$arr = [10 => "ten", 20 => "twenty", 30 => "thirty"];
$sliced = array_slice($arr, 1, 2, true);
echo isset($sliced[20]) ? "KEY_PRESERVED" : "FAIL";
"##,
    );
}

#[test]
fn test_php_array_splice_insert_without_deleting() {
    compile_ok(
        r##"<?php
$a = ["first", "last"];
array_splice($a, 1, 0, ["middle"]);
echo implode(",", $a);
"##,
    );
}

#[test]
fn test_php_array_chunk_empty_array() {
    compile_ok(
        r##"<?php
$c = array_chunk([], 3);
echo count($c) === 0 ? "EMPTY_CHUNK_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_array_combine_mismatched_lengths_throws() {
    compile_ok(
        r##"<?php
try {
    @array_combine(["a"], [1, 2]);
} catch (ValueError $e) {
    echo "Mismatched length caught";
}
"##,
    );
}

#[test]
fn test_php_array_splice_returns_reindexed_removed() {
    let out = run_prints(
        r##"<?php
$input = ["a", "b", "c", "d"];
$removed = array_splice($input, 1, 2);
echo implode("-", $removed) . "|" . count($removed) . "|" . $removed[0];
"##,
    );
    assert_eq!(out, vec!["b-c|2|b"]);
}

#[test]
fn test_php_array_slice_negative_length() {
    let out = run_prints(
        r##"<?php
$arr = [1, 2, 3, 4, 5];
$s = array_slice($arr, 1, -1);
echo implode(",", $s);
"##,
    );
    assert_eq!(out, vec!["2,3,4"]);
}

#[test]
fn test_php_array_chunk_non_divisible_chunk_size() {
    let out = run_prints(
        r##"<?php
$chunks = array_chunk([1,2,3,4,5], 2, false);
echo count($chunks) . "|" . implode("|", array_map("count", $chunks));
"##,
    );
    assert_eq!(out, vec!["3|2|2|1"]);
}

#[test]
fn test_php_array_slice_preserve_keys_false_reindex() {
    let out = run_prints(
        r##"<?php
$arr = ["x" => 10, "y" => 20, "z" => 30];
$s = array_slice($arr, 1, 2, false);
echo implode(",", array_keys($s)) . "|" . implode(",", $s);
"##,
    );
    assert_eq!(out, vec!["0,1|20,30"]);
}
