use super::helpers::compile_ok;

// ── shuffle ──────────────────────────────────────────────────────
#[test]
fn shuffle_randomize_array_order() {
    compile_ok(r#"<?php
$a = [1, 2, 3, 4, 5, 6];
shuffle($a);
echo count($a);
echo is_array($a) ? "array" : "not";
"#);
}

// ── array_rand single key ────────────────────────────────────────
#[test]
fn array_rand_single_random_key() {
    compile_ok(r#"<?php
$a = ["alpha" => 1, "beta" => 2, "gamma" => 3];
$key = array_rand($a);
echo array_key_exists($key, $a) ? "found" : "missing";
"#);
}

// ── array_rand with count ────────────────────────────────────────
#[test]
fn array_rand_multiple_random_keys() {
    compile_ok(r#"<?php
$a = ["a" => 1, "b" => 2, "c" => 3, "d" => 4, "e" => 5];
$keys = array_rand($a, 3);
echo count($keys);
echo is_array($keys) ? "array" : "not";
"#);
}

// ── array_change_key_case CASE_UPPER ─────────────────────────────
#[test]
fn array_change_key_case_upper() {
    compile_ok(r#"<?php
$a = ["first" => 1, "Second" => 2, "THIRD" => 3];
$upper = array_change_key_case($a, CASE_UPPER);
echo array_key_exists("FIRST", $upper) ? "yes" : "no";
echo array_key_exists("SECOND", $upper) ? "yes" : "no";
echo array_key_exists("THIRD", $upper) ? "yes" : "no";
"#);
}

// ── array_change_key_case CASE_LOWER ─────────────────────────────
#[test]
fn array_change_key_case_lower() {
    compile_ok(r#"<?php
$a = ["FOO" => 10, "Bar" => 20, "baz" => 30];
$lower = array_change_key_case($a, CASE_LOWER);
echo array_key_exists("foo", $lower) ? "yes" : "no";
echo array_key_exists("bar", $lower) ? "yes" : "no";
echo $lower["baz"];
"#);
}

// ── array_udiff ──────────────────────────────────────────────────
#[test]
fn array_udiff_user_value_comparison() {
    compile_ok(r#"<?php
$a = [1, 2, 3, 4, 5];
$b = [3, 4, 5, 6, 7];
$diff = array_udiff($a, $b, function($x, $y) { return $x - $y; });
echo implode(",", $diff);
"#);
}

// ── array_udiff_assoc ────────────────────────────────────────────
#[test]
fn array_udiff_assoc_user_value_comparison() {
    compile_ok(r#"<?php
$a = ["x" => 1, "y" => 2, "z" => 3];
$b = ["x" => 1, "y" => 9, "w" => 3];
$diff = array_udiff_assoc($a, $b, function($v1, $v2) { return $v1 - $v2; });
echo implode(",", array_keys($diff));
"#);
}

// ── array_uintersect ─────────────────────────────────────────────
#[test]
fn array_uintersect_user_value_comparison() {
    compile_ok(r#"<?php
$a = [1, 2, 3, 4];
$b = [2, 3, 5, 6];
$common = array_uintersect($a, $b, function($x, $y) { return $x - $y; });
echo implode(",", $common);
"#);
}

// ── array_uintersect_assoc ───────────────────────────────────────
#[test]
fn array_uintersect_assoc_user_comparison() {
    compile_ok(r#"<?php
$a = ["a" => 1, "b" => 2, "c" => 3];
$b = ["a" => 1, "b" => 9, "c" => 3];
$result = array_uintersect_assoc($a, $b, function($v1, $v2) { return $v1 - $v2; });
echo implode(",", array_keys($result));
"#);
}

// ── array_multisort ──────────────────────────────────────────────
#[test]
fn array_multisort_multiple_arrays() {
    compile_ok(r#"<?php
$data = [3, 1, 4, 1, 5, 9, 2, 6];
$keys = ["c", "a", "d", "a2", "e", "i", "b", "f"];
array_multisort($data, SORT_ASC, $keys);
echo count($data);
echo is_array($keys) ? "array" : "not";
"#);
}

// ── preg_grep ────────────────────────────────────────────────────
#[test]
fn preg_grep_matching_entries() {
    compile_ok(r#"<?php
$input = ["foo1", "bar", "foo2", "baz", "foo3"];
$matches = preg_grep('/^foo/', $input);
echo count($matches);
echo is_array($matches) ? "array" : "not";
"#);
}

// ── array_key_exists with integer key 0 ─────────────────────────
#[test]
fn array_key_exists_integer_key_zero() {
    compile_ok(r#"<?php
$a = [0 => "first", 1 => "second", 2 => "third"];
echo array_key_exists(0, $a) ? "exists" : "missing";
echo array_key_exists(3, $a) ? "exists" : "missing";
$empty = [];
echo array_key_exists(0, $empty) ? "exists" : "missing";
"#);
}

// ── in_array loose comparison to null ───────────────────────────
#[test]
fn in_array_loose_comparison_to_null() {
    compile_ok(r#"<?php
$a = [null, false, 0, ""];
echo in_array(null, $a) ? "found" : "not";
echo in_array(null, $a, true) ? "strict-found" : "strict-not";
echo in_array(false, $a, true) ? "found" : "not";
"#);
}

// ── array_splice returning removed elements ──────────────────────
#[test]
fn array_splice_returns_removed_elements() {
    compile_ok(r#"<?php
$a = ["a", "b", "c", "d", "e"];
$removed = array_splice($a, 1, 2);
echo count($removed);
echo implode(",", $removed);
echo count($a);
"#);
}

// ── array_push returning new count ──────────────────────────────
#[test]
fn array_push_returns_new_count() {
    compile_ok(r#"<?php
$a = [1, 2, 3];
$count = array_push($a, 4, 5, 6);
echo $count;
echo count($a);
echo $a[5];
"#);
}
