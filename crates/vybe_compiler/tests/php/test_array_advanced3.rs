use super::helpers::run_prints;

// ── Nested array comprehension patterns ──────────────────────

#[test]
fn matrix_transpose() {
    assert_eq!(
        run_prints(
            r#"<?php
$matrix = [[1,2,3],[4,5,6],[7,8,9]];
$transposed = array_map(null, ...$matrix);
echo $transposed[1][0] . ',' . $transposed[0][1];
"#
        ),
        vec!["2,4"]
    );
}
#[test]
fn array_zip_with_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$names = ['Alice','Bob','Charlie'];
$scores = [85,92,78];
$zipped = array_combine($names, $scores);
arsort($zipped);
echo array_key_first($zipped) . ':' . $zipped[array_key_first($zipped)];
"#
        ),
        vec!["Bob:92"]
    );
}
#[test]
fn group_by_pattern() {
    assert_eq!(
        run_prints(
            r#"<?php
$items = [['type'=>'a','v'=>1],['type'=>'b','v'=>2],['type'=>'a','v'=>3],['type'=>'b','v'=>4]];
$grouped = [];
foreach ($items as $item) $grouped[$item['type']][] = $item['v'];
echo implode(',', $grouped['a']) . ':' . implode(',', $grouped['b']);
"#
        ),
        vec!["1,3:2,4"]
    );
}
#[test]
fn array_first_and_last() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [10,20,30,40,50];
echo array_key_first($a) . ':' . $a[array_key_first($a)];
echo ',';
echo array_key_last($a) . ':' . $a[array_key_last($a)];
"#
        ),
        vec!["0:10,4:50"]
    );
}

// ── array_map with keys ───────────────────────────────────────

#[test]
fn array_map_preserves_assoc_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$prices = ['apple'=>1.0,'banana'=>0.5,'cherry'=>2.0];
$discounted = array_map(fn($p) => $p * 0.9, $prices);
echo round($discounted['apple'], 1) . ',' . round($discounted['cherry'], 1);
"#
        ),
        vec!["0.9,1.8"]
    );
}
#[test]
fn array_keys_values_reindex() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [5=>'a', 2=>'b', 9=>'c'];
$vals = array_values($a);
$keys = array_keys($a);
echo implode(',', $keys) . '|' . implode(',', $vals);
"#
        ),
        vec!["5,2,9|a,b,c"]
    );
}

// ── Searching in arrays ───────────────────────────────────────

#[test]
fn array_search_strict() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1, '2', 3, '4'];
var_export(array_search('2', $a, true));
echo ',';
var_export(array_search(2, $a, true));
"#
        ),
        vec!["1,false"]
    );
}
#[test]
fn in_array_strict() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1, 2, 3];
echo in_array('1', $a, true) ? 'yes' : 'no';
echo in_array(1, $a, true) ? 'yes' : 'no';
"#
        ),
        vec!["noyes"]
    );
}
#[test]
fn array_key_exists_vs_isset() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['key' => null];
echo array_key_exists('key', $a) ? 'key_exists' : 'no_key';
echo isset($a['key']) ? 'isset' : 'not_isset';
"#
        ),
        vec!["key_existsnot_isset"]
    );
}

// ── Merging and slicing ───────────────────────────────────────

#[test]
fn array_merge_reindexes_numeric() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1=>10, 2=>20];
$b = [1=>30, 2=>40];
$merged = array_merge($a, $b);
echo implode(',', $merged);
"#
        ),
        vec!["10,20,30,40"]
    );
}
#[test]
fn union_operator_preserves_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1=>'a', 2=>'b'];
$b = [2=>'x', 3=>'c'];
$result = $a + $b;
echo implode(',', $result);
"#
        ),
        vec!["a,b,c"]
    );
}
#[test]
fn array_slice_preserve_keys_true() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['a'=>1,'b'=>2,'c'=>3,'d'=>4];
$slice = array_slice($a, 1, 2, true);
echo implode(',', array_keys($slice));
"#
        ),
        vec!["b,c"]
    );
}
#[test]
fn array_slice_preserve_keys_false() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['a'=>1,'b'=>2,'c'=>3];
$slice = array_slice($a, 1, 2, false);
echo implode(',', array_keys($slice));
"#
        ),
        vec!["0,1"]
    );
}

// ── Comparison and set operations ────────────────────────────

#[test]
fn array_diff_assoc() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['a'=>1,'b'=>2,'c'=>3];
$b = ['a'=>1,'b'=>5];
$diff = array_diff_assoc($a, $b);
echo implode(',', array_keys($diff));
"#
        ),
        vec!["b,c"]
    );
}
#[test]
fn array_intersect_assoc() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['a'=>1,'b'=>2,'c'=>3];
$b = ['a'=>1,'b'=>5,'c'=>3];
$inter = array_intersect_assoc($a, $b);
echo implode(',', array_keys($inter));
"#
        ),
        vec!["a,c"]
    );
}
#[test]
fn array_udiff_custom_comparison() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1,2,3,4,5];
$b = [2,4];
$diff = array_udiff($a, $b, fn($x,$y) => $x <=> $y);
echo implode(',', $diff);
"#
        ),
        vec!["1,3,5"]
    );
}

// ── Recursion in arrays ───────────────────────────────────────

#[test]
fn array_walk_recursive_deep() {
    assert_eq!(
        run_prints(
            r#"<?php
$tree = ['a' => [1, 2, ['b' => [3, 4]]], 'c' => 5];
$sum = 0;
array_walk_recursive($tree, function($v) use (&$sum) { $sum += $v; });
echo $sum;
"#
        ),
        vec!["15"]
    );
}
#[test]
fn array_map_nested_transform() {
    assert_eq!(
        run_prints(
            r#"<?php
$grid = [[1,2,3],[4,5,6],[7,8,9]];
$doubled = array_map(fn($row) => array_map(fn($v) => $v * 2, $row), $grid);
echo $doubled[1][1];
"#
        ),
        vec!["10"]
    );
}
