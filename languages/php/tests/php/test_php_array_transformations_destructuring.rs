use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Array Transformations & Destructuring — list(), [$a, $b], array_map, array_filter, array_reduce, array_column, array_combine
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_array_short_destructuring_positional() {
    let out = run_prints(
        r#"<?php
$data = [10, 20, 30];
[$a, $b, $c] = $data;
echo "$a-$b-$c";
"#,
    );
    assert_eq!(out, vec!["10-20-30"]);
}

#[test]
fn test_php_array_keyed_destructuring() {
    let out = run_prints(
        r#"<?php
$user = ["name" => "Alice", "role" => "admin", "age" => 30];
["name" => $name, "role" => $role] = $user;
echo "$name is $role";
"#,
    );
    assert_eq!(out, vec!["Alice is admin"]);
}

#[test]
fn test_php_array_nested_destructuring() {
    let out = run_prints(
        r#"<?php
$point = [1, [2, 3]];
[$x, [$y, $z]] = $point;
echo "$x, $y, $z";
"#,
    );
    assert_eq!(out, vec!["1, 2, 3"]);
}

#[test]
fn test_php_array_map_transformation() {
    let out = run_prints(
        r#"<?php
$nums = [1, 2, 3, 4];
$squared = array_map(fn($n) => $n * $n, $nums);
echo implode(", ", $squared);
"#,
    );
    assert_eq!(out, vec!["1, 4, 9, 16"]);
}

#[test]
fn test_php_array_filter_with_use_both_flag() {
    let out = run_prints(
        r#"<?php
$arr = ["a" => 1, "b" => 2, "c" => 3, "d" => 4];
$filtered = array_filter($arr, fn($val, $key) => $val > 2 && $key !== "d", ARRAY_FILTER_USE_BOTH);
echo implode(", ", array_keys($filtered));
"#,
    );
    assert_eq!(out, vec!["c"]);
}

#[test]
fn test_php_array_reduce_accumulator() {
    let out = run_prints(
        r#"<?php
$nums = [10, 20, 30];
$sum = array_reduce($nums, fn($acc, $item) => $acc + $item, 100);
echo $sum;
"#,
    );
    assert_eq!(out, vec!["160"]);
}

#[test]
fn test_php_array_column_extraction() {
    let out = run_prints(
        r#"<?php
$records = [
    ["id" => 101, "name" => "Alice"],
    ["id" => 102, "name" => "Bob"],
    ["id" => 103, "name" => "Charlie"],
];
$names = array_column($records, "name", "id");
echo $names[102];
"#,
    );
    assert_eq!(out, vec!["Bob"]);
}

#[test]
fn test_php_array_combine_keys_and_values() {
    let out = run_prints(
        r#"<?php
$keys = ["x", "y", "z"];
$vals = [10, 20, 30];
$combined = array_combine($keys, $vals);
echo $combined["y"];
"#,
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn test_php_array_chunk_splitting() {
    let out = run_prints(
        r#"<?php
$items = [1, 2, 3, 4, 5];
$chunks = array_chunk($items, 2);
echo count($chunks) . ":" . count($chunks[2]);
"#,
    );
    assert_eq!(out, vec!["3:1"]);
}

#[test]
fn test_php_array_walk_by_reference_mutation() {
    let out = run_prints(
        r#"<?php
$fruits = ["apple", "banana"];
array_walk($fruits, function(&$val, $key) {
    $val = strtoupper($val);
});
echo implode("-", $fruits);
"#,
    );
    assert_eq!(out, vec!["APPLE-BANANA"]);
}

#[test]
fn test_php_array_slice_and_splice_compilation() {
    compile_ok(
        r#"<?php
$input = ["red", "green", "blue", "yellow"];
$output = array_slice($input, 2);
$removed = array_splice($input, 1, 2, ["orange"]);
echo count($output) + count($removed);
"#,
    );
}

#[test]
fn test_php_array_merge_recursive_semantics() {
    compile_ok(
        r#"<?php
$ar1 = ["color" => ["favorite" => "red"], 5];
$ar2 = [10, "color" => ["favorite" => "green", "blue"]];
$result = array_merge_recursive($ar1, $ar2);
print_r($result);
"#,
    );
}

#[test]
fn test_php_array_replace_recursive() {
    compile_ok(
        r#"<?php
$base = ["citrus" => ["orange"], "berries" => ["blackberry"]];
$replacement = ["citrus" => ["pineapple"], "berries" => ["strawberry"]];
$basket = array_replace_recursive($base, $replacement);
print_r($basket);
"#,
    );
}

#[test]
fn test_php_array_fill_keys() {
    compile_ok(
        r#"<?php
$keys = ["foo", 5, 10, "bar"];
$a = array_fill_keys($keys, "default");
print_r($a);
"#,
    );
}

#[test]
fn test_php_array_intersect_key_and_diff_key() {
    compile_ok(
        r#"<?php
$array1 = ['blue' => 1, 'red' => 2, 'green' => 3];
$array2 = ['green' => 5, 'yellow' => 7, 'cyan' => 8];
$intersect = array_intersect_key($array1, $array2);
$diff = array_diff_key($array1, $array2);
echo count($intersect) . count($diff);
"#,
    );
}
