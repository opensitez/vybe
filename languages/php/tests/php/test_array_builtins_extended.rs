use super::helpers::{compile_ok, run_prints};

// ── array_column ─────────────────────────────────────────────────
#[test]
fn array_column_extract_field() {
    compile_ok(
        r#"<?php
$rows = [
    ["id" => 1, "name" => "Alice", "dept" => "Eng"],
    ["id" => 2, "name" => "Bob",   "dept" => "Mkt"],
    ["id" => 3, "name" => "Carol", "dept" => "Eng"],
];
$names = array_column($rows, "name");
echo implode(",", $names);
$byId = array_column($rows, "dept", "id");
echo $byId[2];
"#,
    );
}

// ── array_combine ────────────────────────────────────────────────
#[test]
fn array_combine_keys_values() {
    compile_ok(
        r#"<?php
$keys   = ["a", "b", "c", "d"];
$values = [10, 20, 30, 40];
$map = array_combine($keys, $values);
echo $map["b"];
echo count($map);
"#,
    );
}

// ── array_diff ───────────────────────────────────────────────────
#[test]
fn array_diff_missing_values() {
    compile_ok(
        r#"<?php
$a = ["apple", "banana", "cherry", "date"];
$b = ["banana", "date", "elderberry"];
$diff = array_diff($a, $b);
echo implode(",", $diff);
"#,
    );
}

// ── array_diff_key ───────────────────────────────────────────────
#[test]
fn array_diff_key_missing_keys() {
    compile_ok(
        r#"<?php
$a = ["x" => 1, "y" => 2, "z" => 3];
$b = ["x" => 99, "z" => 100];
$result = array_diff_key($a, $b);
echo implode(",", array_keys($result));
"#,
    );
}

// ── array_diff_assoc ─────────────────────────────────────────────
#[test]
fn array_diff_assoc_key_value_pairs() {
    compile_ok(
        r#"<?php
$a = ["color" => "red",  "size" => "M",  "weight" => 10];
$b = ["color" => "red",  "size" => "L",  "weight" => 10];
$diff = array_diff_assoc($a, $b);
echo implode(",", array_keys($diff));
"#,
    );
}

// ── array_intersect ──────────────────────────────────────────────
#[test]
fn array_intersect_common_values() {
    compile_ok(
        r#"<?php
$a = [1, 2, 3, 4, 5];
$b = [3, 4, 5, 6, 7];
$c = [4, 5, 8, 9];
$common = array_intersect($a, $b, $c);
echo implode(",", $common);
"#,
    );
}

// ── array_intersect_key ──────────────────────────────────────────
#[test]
fn array_intersect_key_shared_keys() {
    compile_ok(
        r#"<?php
$a = ["foo" => 1, "bar" => 2, "baz" => 3];
$b = ["foo" => 99, "baz" => 88];
$result = array_intersect_key($a, $b);
echo implode(",", array_keys($result));
echo implode(",", $result);
"#,
    );
}

// ── array_flip ───────────────────────────────────────────────────
#[test]
fn array_flip_swap_keys_values() {
    compile_ok(
        r#"<?php
$a = ["one" => 1, "two" => 2, "three" => 3];
$flipped = array_flip($a);
echo $flipped[1];
echo $flipped[2];
echo array_key_exists("one", $flipped) ? "bad" : "ok";
"#,
    );
}

// ── array_fill ───────────────────────────────────────────────────
#[test]
fn array_fill_with_start_index() {
    compile_ok(
        r#"<?php
$a = array_fill(5, 4, "x");
echo count($a);
echo $a[5];
echo $a[8];
echo array_key_exists(4, $a) ? "bad" : "ok";
"#,
    );
}

// ── array_fill_keys ──────────────────────────────────────────────
#[test]
fn array_fill_keys_from_array() {
    compile_ok(
        r#"<?php
$keys = ["alpha", "beta", "gamma"];
$a = array_fill_keys($keys, null);
echo count($a);
echo array_key_exists("beta", $a) ? "yes" : "no";
echo array_key_exists("delta", $a) ? "yes" : "no";
"#,
    );
}

// ── array_pad ────────────────────────────────────────────────────
#[test]
fn array_pad_right_and_left() {
    compile_ok(
        r#"<?php
$a = [1, 2, 3];
$right = array_pad($a, 6, 0);
echo count($right);
echo $right[5];
$left = array_pad($a, -6, 9);
echo $left[0];
"#,
    );
}

// ── array_unique ─────────────────────────────────────────────────
#[test]
fn array_unique_deduplicate() {
    compile_ok(
        r#"<?php
$a = ["a", "b", "a", "c", "b", "d", "d"];
$u = array_unique($a);
echo count($u);
echo implode(",", $u);
"#,
    );
}

// ── array_product ────────────────────────────────────────────────
#[test]
fn array_product_multiply_all() {
    compile_ok(
        r#"<?php
echo array_product([1, 2, 3, 4, 5]);
echo array_product([7]);
echo array_product([]);
"#,
    );
}

// ── arsort ───────────────────────────────────────────────────────
#[test]
fn arsort_descending_preserve_keys() {
    compile_ok(
        r#"<?php
$a = ["b" => 2, "d" => 4, "a" => 1, "c" => 3];
arsort($a);
echo implode(",", array_keys($a));
echo implode(",", $a);
"#,
    );
}

// ── ksort ────────────────────────────────────────────────────────
#[test]
fn ksort_ascending_by_key() {
    compile_ok(
        r#"<?php
$a = ["banana" => 2, "apple" => 1, "cherry" => 3];
ksort($a);
echo implode(",", array_keys($a));
"#,
    );
}

// ── krsort ───────────────────────────────────────────────────────
#[test]
fn krsort_descending_by_key() {
    compile_ok(
        r#"<?php
$a = ["alpha" => 10, "gamma" => 30, "beta" => 20];
krsort($a);
echo implode(",", array_keys($a));
"#,
    );
}

// ── uasort ───────────────────────────────────────────────────────
#[test]
fn uasort_custom_callback_preserve_keys() {
    compile_ok(
        r#"<?php
$a = ["p" => 30, "q" => 10, "r" => 20];
uasort($a, function($x, $y) { return $x - $y; });
echo implode(",", array_keys($a));
echo implode(",", $a);
"#,
    );
}

// ── uksort ───────────────────────────────────────────────────────
#[test]
fn uksort_custom_key_comparator() {
    compile_ok(
        r#"<?php
$a = ["cc" => 3, "aaa" => 1, "b" => 2];
uksort($a, fn($x, $y) => strlen($x) - strlen($y));
echo implode(",", array_keys($a));
"#,
    );
}

// ── natsort ──────────────────────────────────────────────────────
#[test]
fn natsort_natural_string_order() {
    compile_ok(
        r#"<?php
$files = ["file10.txt", "file2.txt", "file1.txt", "file20.txt"];
natsort($files);
echo implode(",", $files);
"#,
    );
}

// ── array_splice ─────────────────────────────────────────────────
#[test]
fn array_splice_remove_and_replace() {
    compile_ok(
        r#"<?php
$a = ["a", "b", "c", "d", "e"];
$removed = array_splice($a, 1, 2, ["x", "y", "z"]);
echo implode(",", $a);
echo implode(",", $removed);
"#,
    );
}

// ── array_chunk ──────────────────────────────────────────────────
#[test]
fn array_chunk_split_into_groups() {
    compile_ok(
        r#"<?php
$a = range(1, 7);
$chunks = array_chunk($a, 3);
echo count($chunks);
echo count($chunks[0]);
echo count($chunks[2]);
"#,
    );
}

// ── array_count_values ───────────────────────────────────────────
#[test]
fn array_count_values_frequency_map() {
    compile_ok(
        r#"<?php
$a = ["red", "blue", "red", "green", "blue", "red"];
$freq = array_count_values($a);
echo $freq["red"];
echo $freq["blue"];
echo $freq["green"];
"#,
    );
}

// ── array_replace ────────────────────────────────────────────────
#[test]
fn array_replace_overlay_values() {
    compile_ok(
        r#"<?php
$base    = ["a" => 1, "b" => 2, "c" => 3];
$overlay = ["b" => 20, "d" => 40];
$result  = array_replace($base, $overlay);
echo $result["a"];
echo $result["b"];
echo $result["d"];
echo array_key_exists("c", $result) ? "yes" : "no";
"#,
    );
}

// ── array_walk_recursive ─────────────────────────────────────────
#[test]
fn array_walk_recursive_leaf_callback() {
    compile_ok(
        r#"<?php
$nested = [1, [2, 3], [[4], 5]];
$collected = [];
array_walk_recursive($nested, function($val) use (&$collected) {
    $collected[] = $val * 2;
});
echo implode(",", $collected);
"#,
    );
}

// ── extract ──────────────────────────────────────────────────────
#[test]
fn extract_variables_from_array() {
    compile_ok(
        r#"<?php
$data = ["username" => "alice", "role" => "admin", "level" => 5];
extract($data);
echo $username;
echo $role;
echo $level;
"#,
    );
}

#[test]
fn array_search_finds_last_duplicate_key_preference() {
    let out = run_prints(
        r#"<?php
$a = [0 => "first", 1 => "dup", 2 => "dup"];
echo array_search("dup", $a);
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn array_combine_mismatched_counts_throws() {
    let out = run_prints(
        r#"<?php
try {
    array_combine([1, 2], [1]);
    echo "no-error";
} catch (ValueError $e) {
    echo "value-error";
}
"#,
    );
    assert_eq!(out, vec!["value-error"]);
}

#[test]
fn array_diff_assoc_keeps_associative_only() {
    let out = run_prints(
        r#"<?php
$a = ["x" => 1, "y" => 2, "z" => 3];
$b = ["x" => 9, "y" => 2];
$d = array_diff_assoc($a, $b);
ksort($d);
echo json_encode(array_keys($d)) . "|" . $d["z"];
"#,
    );
    assert_eq!(out, vec!["[\"x\",\"z\"]|3"]);
}

#[test]
fn sort_key_order_stable_on_equal_values() {
    let out = run_prints(
        r#"<?php
$a = [2 => "same", 1 => "same", 0 => "low"];
asort($a);
echo implode(',', array_keys($a));
"#,
    );
    assert_eq!(out, vec!["2,1,0"]);
}

#[test]
fn array_search_loose_zero_matches_false_or_null_string() {
    let out = run_prints(
        r#"<?php
$a = [false, null, 0, "0"];
echo in_array("", $a) ? "yes" : "no";
echo "|";
echo in_array(0, $a) ? "yes0" : "no0";
"#,
    );
    assert_eq!(out, vec!["no|yes0"]);
}

#[test]
fn array_search_find_missing_with_strict() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ["10", 20, 30];
$first = array_search(10, $a, true);
echo $first === false ? "f" : $first;
echo array_search("10", $a, true);
"#,
        ),
        &["f0"]
    );
}

#[test]
fn array_filter_use_value_and_key_modes() {
    assert_eq!(
        run_prints(
            r#"<?php
$data = ["a" => 1, "b" => 0, "c" => 3, "d" => "0"];
$vals = array_filter($data, fn($v) => $v, 0);
$keys = array_filter($data, fn($k) => $k === "a" || $k === "d", ARRAY_FILTER_USE_KEY);
echo count($vals) . "|" . count($keys);
"#,
        ),
        &["2|2"]
    );
}

#[test]
fn array_fill_keys_collision_keeps_last() {
    assert_eq!(
        run_prints(
            r#"<?php
$keys = ["x", "y", "x"];
$a = array_fill_keys($keys, 0);
$a["x"] = 9;
echo $a["x"] . "|" . count($a);
"#,
        ),
        &["9|2"]
    );
}

#[test]
fn array_push_pop_shift_unshift_sequence() {
    assert_eq!(
        run_prints(
            r#"<?php
$q = [];
array_push($q, 1, 2);
array_unshift($q, 0);
$last = array_pop($q);
$first = array_shift($q);
echo $last . "|" . $first . "|" . implode("-", $q);
"#,
        ),
        &["2|0|1"]
    );
}

#[test]
fn array_merge_union_and_numeric_index_behavior() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [0 => "a", 1 => "b", "x" => 1];
$b = [0 => "z", 2 => "c", "y" => 2];
$m = array_merge($a, $b);
$u = $a + $b;
echo $m[0] . $m[1] . $m[2] . $m[3];
echo "|";
echo $u[0] . $u[1] . $u["x"] . (array_key_exists("y", $u) ? "y" : "n");
"#,
        ),
        &["abzc|ab1y"]
    );
}
