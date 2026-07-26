use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Array Advanced Operations — array_count_values, array_product, array_sum, array_diff_assoc, array_intersect_assoc, array_udiff, array_uintersect
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_array_count_values_frequency_map() {
    let out = run_prints(
        r#"<?php
$values = ["apple", "banana", "apple", "cherry", "banana", "apple"];
$counts = array_count_values($values);
echo "apple=" . $counts["apple"] . " banana=" . $counts["banana"];
"#,
    );
    assert_eq!(out, vec!["apple=3 banana=2"]);
}

#[test]
fn test_php_array_sum_and_product_math() {
    let out = run_prints(
        r#"<?php
$nums = [2, 3, 4];
$sum = array_sum($nums);
$product = array_product($nums);
echo "Sum=$sum Product=$product";
"#,
    );
    assert_eq!(out, vec!["Sum=9 Product=24"]);
}

#[test]
fn test_php_array_diff_assoc_and_intersect_assoc() {
    let out = run_prints(
        r#"<?php
$a1 = ["a" => "green", "b" => "brown", "c" => "blue", "red"];
$a2 = ["a" => "green", "yellow", "red"];

$diff = array_diff_assoc($a1, $a2);
$intersect = array_intersect_assoc($a1, $a2);

echo "DiffCount=" . count($diff) . " IntersectCount=" . count($intersect);
"#,
    );
    assert_eq!(out, vec!["DiffCount=3 IntersectCount=1"]);
}

#[test]
fn test_php_array_udiff_custom_comparator_callback() {
    let out = run_prints(
        r#"<?php
$a1 = [10, 20, 30, 40];
$a2 = [15, 20, 35, 40];

$diff = array_udiff($a1, $a2, fn($x, $y) => $x <=> $y);
echo implode(",", $diff);
"#,
    );
    assert_eq!(out, vec!["10,30"]);
}

#[test]
fn test_php_array_uintersect_custom_comparator() {
    compile_ok(
        r#"<?php
$a1 = ["a" => 1, "b" => 2, "c" => 3];
$a2 = ["x" => 2, "y" => 3, "z" => 4];

$intersect = array_uintersect($a1, $a2, fn($v1, $v2) => $v1 <=> $v2);
echo implode(",", $intersect);
"#,
    );
}

#[test]
fn test_php_array_diff_uassoc_key_and_value_callback() {
    compile_ok(
        r#"<?php
$a1 = ["a" => 1, "b" => 2];
$a2 = ["A" => 1, "B" => 2];

$diff = array_diff_uassoc($a1, $a2, fn($k1, $k2) => strcasecmp($k1, $k2));
echo count($diff); // empty because keys match case-insensitively
"#,
    );
}

#[test]
fn test_php_array_intersect_uassoc_callback() {
    compile_ok(
        r#"<?php
$a1 = ["a" => 1, "b" => 2];
$a2 = ["A" => 1, "B" => 3];

$intersect = array_intersect_uassoc($a1, $a2, fn($k1, $k2) => strcasecmp($k1, $k2));
echo count($intersect); // matches "a" => 1
"#,
    );
}

#[test]
fn test_php_array_pad_expansion() {
    compile_ok(
        r#"<?php
$input = [12, 10, 9];
$result = array_pad($input, 5, 0);
echo implode(",", $result);
"#,
    );
}

#[test]
fn test_php_array_fill_negative_start_index() {
    compile_ok(
        r#"<?php
$a = array_fill(-2, 3, "val");
echo implode(",", array_keys($a)); // -2, 0, 1
"#,
    );
}

#[test]
fn test_php_array_reduce_string_concatenation() {
    compile_ok(
        r#"<?php
$words = ["PHP", "Is", "Great"];
$sentence = array_reduce($words, fn($carry, $w) => $carry === "" ? $w : "$carry $w", "");
echo $sentence;
"#,
    );
}

#[test]
fn test_php_array_count_values_string_number_coercion() {
    let out = run_prints(
        r#"<?php
$values = [1, "1", 2, "2", "2", 1];
$counts = array_count_values($values);
echo "1=" . $counts["1"] . " ";
echo "2=" . $counts["2"];
"#,
    );
    assert_eq!(out, vec!["1=3 2=3"]);
}

#[test]
fn test_php_array_fill_negative_count_throws() {
    let out = run_prints(
        r#"<?php
try {
    array_fill(0, -2, "x");
    echo "no-error";
} catch (ValueError $e) {
    echo "value-error";
}
"#,
    );
    assert_eq!(out, vec!["value-error"]);
}

#[test]
fn test_php_array_udiff_with_comparator_value_and_key() {
    let out = run_prints(
        r#"<?php
$a1 = ["a" => 1, "b" => 2, "c" => 3];
$a2 = ["a" => 1, "B" => 2, "c" => 4];
$diff = array_udiff_assoc(
    $a1,
    $a2,
    fn($v1, $v2) => $v1 <=> $v2
);
echo count($diff) . "|" . implode(",", $diff);
"#,
    );
    assert_eq!(out, vec!["2|2,3"]);
}

#[test]
fn test_php_array_intersect_assoc_preserves_key_order() {
    let out = run_prints(
        r#"<?php
$a1 = ["x" => 1, "y" => 2, "z" => 3];
$a2 = ["y" => 2, "x" => 1];
$i = array_intersect_assoc($a1, $a2);
echo implode(",", array_keys($i)) . ":" . implode(",", $i);
"#,
    );
    assert_eq!(out, vec!["x,y:1,2"]);
}
