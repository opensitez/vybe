use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Array Sorting, Multisort & Callbacks — sort, asort, ksort, usort, uasort, uksort, array_multisort, shuffle
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_usort_spaceship_operator() {
    let out = run_prints(
        r#"<?php
$users = [
    ["name" => "Bob", "age" => 30],
    ["name" => "Alice", "age" => 25],
    ["name" => "Charlie", "age" => 25],
];

usort($users, fn($a, $b) => $a["age"] <=> $b["age"] ?: $a["name"] <=> $b["name"]);
$sortedNames = array_column($users, "name");
echo implode(", ", $sortedNames);
"#,
    );
    assert_eq!(out, vec!["Alice, Charlie, Bob"]);
}

#[test]
fn test_php_asort_and_ksort_association_preservation() {
    let out = run_prints(
        r#"<?php
$fruit = ["d" => "lemon", "a" => "orange", "b" => "banana", "c" => "apple"];
asort($fruit);
echo implode(",", array_keys($fruit)) . " | " . implode(",", $fruit);
"#,
    );
    assert_eq!(out, vec!["b,d,a,c | banana,lemon,orange,apple"]);
}

#[test]
fn test_php_array_multisort_multiple_columns() {
    let out = run_prints(
        r#"<?php
$volume = [67, 86, 85, 98, 86, 67];
$edition = [2, 1, 6, 2, 1, 6];

array_multisort($volume, SORT_DESC, $edition, SORT_ASC);
echo "v0={$volume[0]} e0={$edition[0]} | v1={$volume[1]} e1={$edition[1]}";
"#,
    );
    assert_eq!(out, vec!["v0=98 e0=2 | v1=86 e1=1"]);
}

#[test]
fn test_php_uksort_key_comparator_callback() {
    let out = run_prints(
        r#"<?php
$data = ["2024-05-01" => "a", "2024-01-01" => "b", "2024-12-01" => "c"];
uksort($data, fn($k1, $k2) => strtotime($k1) <=> strtotime($k2));
echo implode(",", array_keys($data));
"#,
    );
    assert_eq!(out, vec!["2024-01-01,2024-05-01,2024-12-01"]);
}

#[test]
fn test_php_natcasesort_natural_order_sorting() {
    compile_ok(
        r#"<?php
$files = ["img12.png", "img10.png", "img2.png", "img1.png"];
natcasesort($files);
echo implode(",", $files);
"#,
    );
}

#[test]
fn test_php_array_rand_pick_keys() {
    compile_ok(
        r#"<?php
$input = ["Neo", "Morpheus", "Trinity", "Cypher", "Tank"];
$randKey = array_rand($input);
echo $input[$randKey];
"#,
    );
}

#[test]
fn test_php_shuffle_randomize_array() {
    compile_ok(
        r#"<?php
$numbers = range(1, 10);
shuffle($numbers);
echo count($numbers);
"#,
    );
}

#[test]
fn test_php_arsort_reverse_association() {
    compile_ok(
        r#"<?php
$scores = ["Alice" => 90, "Bob" => 95, "Charlie" => 85];
arsort($scores);
$top = array_key_first($scores);
echo "$top=" . $scores[$top];
"#,
    );
}

#[test]
fn test_php_krsort_key_reverse_order() {
    compile_ok(
        r#"<?php
$arr = [1 => "a", 3 => "c", 2 => "b"];
krsort($arr);
echo implode(",", array_keys($arr));
"#,
    );
}

#[test]
fn test_php_uasort_custom_association_sort() {
    compile_ok(
        r#"<?php
$data = ["a" => 4, "b" => 2, "c" => 8];
uasort($data, fn($v1, $v2) => $v1 <=> $v2);
echo array_key_first($data);
"#,
    );
}
