use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Array Functions: array_key_exists, in_array, array_search, array_keys, array_values
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_array_key_exists_null_value() {
    let out = run_prints(
        r##"<?php
$a = ["key" => null];
echo array_key_exists("key", $a) ? "EXISTS" : "MISSING";
"##,
    );
    assert_eq!(out, vec!["EXISTS"]);
}

#[test]
fn test_php_in_array_strict_comparison_types() {
    let out = run_prints(
        r##"<?php
$arr = [1, "2", 3];
$loose = in_array(2, $arr);
$strict = in_array(2, $arr, true);
echo ($loose ? "L1" : "L0") . " " . ($strict ? "S1" : "S0");
"##,
    );
    assert_eq!(out, vec!["L1 S0"]);
}

#[test]
fn test_php_array_search_returns_first_matching_key() {
    let out = run_prints(
        r##"<?php
$arr = ["blue" => 10, "red" => 20, "green" => 10];
$key = array_search(10, $arr);
echo "Key: $key";
"##,
    );
    assert_eq!(out, vec!["Key: blue"]);
}

#[test]
fn test_php_array_keys_filter_by_search_value() {
    let out = run_prints(
        r##"<?php
$arr = ["a" => "apple", "b" => "banana", "c" => "apple"];
$keys = array_keys($arr, "apple", true);
echo implode(",", $keys);
"##,
    );
    assert_eq!(out, vec!["a,c"]);
}

#[test]
fn test_php_array_values_reindexes_keys() {
    let out = run_prints(
        r##"<?php
$arr = [10 => "x", 20 => "y", 30 => "z"];
$reindexed = array_values($arr);
echo implode(":", array_keys($reindexed)) . " = " . implode(",", $reindexed);
"##,
    );
    assert_eq!(out, vec!["0:1:2 = x,y,z"]);
}

#[test]
fn test_php_array_key_exists_object_property_access() {
    compile_ok(
        r##"<?php
class Container implements ArrayAccess {
    private array $container = ["foo" => "bar"];
    public function offsetExists(mixed $offset): bool { return isset($this->container[$offset]); }
    public function offsetGet(mixed $offset): mixed { return $this->container[$offset]; }
    public function offsetSet(mixed $offset, mixed $value): void { $this->container[$offset] = $value; }
    public function offsetUnset(mixed $offset): void { unset($this->container[$offset]); }
}
$c = new Container();
echo isset($c["foo"]) ? "OFFSET_EXISTS_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_in_array_multidimensional_array_search() {
    compile_ok(
        r##"<?php
$user1 = ["id" => 1, "name" => "Alice"];
$user2 = ["id" => 2, "name" => "Bob"];
$list = [$user1, $user2];
echo in_array(["id" => 1, "name" => "Alice"], $list, true) ? "STRICT_STRUCT_FOUND" : "FAIL";
"##,
    );
}

#[test]
fn test_php_array_search_not_found_returns_false() {
    compile_ok(
        r##"<?php
$arr = [1, 2, 3];
$res = array_search(99, $arr, true);
echo $res === false ? "NOT_FOUND_FALSE" : "FAIL";
"##,
    );
}

#[test]
fn test_php_array_keys_empty_array_returns_empty() {
    compile_ok(
        r##"<?php
$k = array_keys([]);
echo count($k) === 0 ? "EMPTY_KEYS_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_array_key_first_and_last_php73() {
    compile_ok(
        r##"<?php
$a = ["first" => 1, "mid" => 2, "last" => 3];
echo "First=" . array_key_first($a) . " Last=" . array_key_last($a);
"##,
    );
}

#[test]
fn test_php_array_key_exists_with_int_and_numeric_string_keys() {
    let out = run_prints(
        r##"<?php
$a = ["0" => "zero", 1 => "one", 1.2 => "float1", true => "bool1"];
echo (array_key_exists(0, $a) ? "0ok" : "0no") . "|";
echo (array_key_exists("1", $a) ? "1ok" : "1no") . "|";
echo (array_key_exists(1, $a) ? "1intok" : "1noint");
"##,
    );
    assert_eq!(out, vec!["0ok|1ok|1intok"]);
}

#[test]
fn test_php_in_array_loose_and_strict_on_null_false_zero() {
    let out = run_prints(
        r##"<?php
$a = [0, "0", false, null];
echo (in_array("", $a) ? "empty-loose" : "empty-loose-no") . "|";
echo (in_array("", $a, true) ? "empty-strict" : "empty-strict-no") . "|";
echo (in_array(0, $a, true) ? "zero-strict" : "zero-strict-no");
"##,
    );
    assert_eq!(out, vec!["empty-loose-no|empty-strict-no|zero-strict-no"]);
}

#[test]
fn test_php_array_search_with_strict_matching_key_type() {
    let out = run_prints(
        r##"<?php
$a = ["1" => "one", 2 => "two", "3" => "three"];
echo array_search("2", $a) . "|";
echo (array_search("2", $a, true) === false ? "strict-no" : "strict-yes");
"##,
    );
    assert_eq!(out, vec!["1|strict-no"]);
}

#[test]
fn test_php_array_values_empty_array_is_empty() {
    let out = run_prints(
        r##"<?php
$out = array_values([]);
echo count($out) . "|" . json_encode($out);
"##,
    );
    assert_eq!(out, vec!["0|[]"]);
}

#[test]
fn test_php_array_keys_with_object_values() {
    let out = run_prints(
        r##"<?php
$o1 = (object)["id" => 1];
$o2 = (object)["id" => 2];
$a = ["a" => $o1, "b" => $o2];
$keys = array_keys($a);
echo implode(",", $keys);
"##,
    );
    assert_eq!(out, vec!["a,b"]);
}

#[test]
fn test_php_in_array_with_type_juggling_false_negative() {
    let out = run_prints(
        r##"<?php
$a = ["0", 0, false];
echo (in_array("false", $a) ? "has-false" : "no-false") . "|";
echo (in_array(false, $a) ? "has-bool-false" : "no-bool-false");
"##,
    );
    assert_eq!(out, vec!["no-false|has-bool-false"]);
}
