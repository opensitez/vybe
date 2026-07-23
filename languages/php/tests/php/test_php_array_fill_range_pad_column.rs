use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Array Functions: array_fill, range, array_pad, array_column
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_array_fill_creates_array_with_fixed_value() {
    let out = run_prints(
        r##"<?php
$a = array_fill(5, 3, "banana");
echo implode(",", array_keys($a)) . " = " . $a[5];
"##,
    );
    assert_eq!(out, vec!["5,6,7 = banana"]);
}

#[test]
fn test_php_range_numeric_step() {
    let out = run_prints(
        r##"<?php
$r = range(0, 10, 2);
echo implode(",", $r);
"##,
    );
    assert_eq!(out, vec!["0,2,4,6,8,10"]);
}

#[test]
fn test_php_range_character_sequence() {
    let out = run_prints(
        r##"<?php
$letters = range("a", "d");
echo implode("", $letters);
"##,
    );
    assert_eq!(out, vec!["abcd"]);
}

#[test]
fn test_php_array_pad_positive_negative_length() {
    let out = run_prints(
        r##"<?php
$input = [12, 10, 9];
$padded_right = array_pad($input, 5, 0);
$padded_left = array_pad($input, -5, -1);
echo implode(",", $padded_right) . " | " . implode(",", $padded_left);
"##,
    );
    assert_eq!(out, vec!["12,10,9,0,0 | -1,-1,12,10,9"]);
}

#[test]
fn test_php_array_column_associative_index_key() {
    let out = run_prints(
        r##"<?php
$records = [
    ["id" => 2135, "first_name" => "John", "last_name" => "Doe"],
    ["id" => 3245, "first_name" => "Sally", "last_name" => "Smith"],
];
$last_names = array_column($records, "last_name", "id");
echo "2135:" . $last_names[2135] . " 3245:" . $last_names[3245];
"##,
    );
    assert_eq!(out, vec!["2135:Doe 3245:Smith"]);
}

#[test]
fn test_php_array_fill_keys_custom_keys() {
    compile_ok(
        r##"<?php
$keys = ["foo", 5, 10, "bar"];
$a = array_fill_keys($keys, "default");
echo $a["foo"] === "default" && $a[5] === "default" ? "FILL_KEYS_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_range_descending_order() {
    compile_ok(
        r##"<?php
$desc = range(10, 1);
echo count($desc) === 10 && $desc[0] === 10 ? "DESC_RANGE_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_array_column_object_properties() {
    compile_ok(
        r##"<?php
class Person {
    public function __construct(public int $id, public string $name) {}
}
$people = [new Person(1, "Alice"), new Person(2, "Bob")];
$names = array_column($people, "name");
echo implode(",", $names);
"##,
    );
}

#[test]
fn test_php_array_pad_smaller_than_input_no_change() {
    compile_ok(
        r##"<?php
$a = [1, 2, 3];
$padded = array_pad($a, 2, "x");
echo count($padded) === 3 ? "NO_PAD_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_array_fill_zero_num_returns_empty() {
    compile_ok(
        r##"<?php
$res = array_fill(0, 0, "val");
echo count($res) === 0 ? "ZERO_NUM_OK" : "FAIL";
"##,
    );
}
