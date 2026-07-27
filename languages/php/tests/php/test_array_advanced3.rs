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
        vec!["b,c"]
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

#[test]
fn array_slice_negative_offset_and_length() {
    assert_eq!(
        run_prints(
            r#"<?php
$items = [10,20,30,40,50,60];
$tail = array_slice($items, -4, 2);
echo implode('-', $tail);
"#,
        ),
        vec!["30-40"]
    );
}

#[test]
fn array_search_returns_first_match_in_order() {
    assert_eq!(
        run_prints(
            r#"<?php
$data = [5, 7, 5, 9];
$idx = array_search(5, $data, false);
echo $idx;
"#,
        ),
        vec!["0"]
    );
}

#[test]
fn array_replace_preserves_non_overwritten_strings() {
    assert_eq!(
        run_prints(
            r#"<?php
$base = ['a' => 'keep', 'b' => 'replace'];
$patch = ['b' => 'new'];
$result = array_replace($base, $patch);
echo $result['a'] . '|' . $result['b'];
"#,
        ),
        vec!["keep|new"]
    );
}

#[test]
fn array_fill_with_implicit_string_key_values() {
    assert_eq!(
        run_prints(
            r#"<?php
$values = array_fill(0, 3, 'x');
echo $values[0] . '|' . $values[2] . '|' . count($values);
"#,
        ),
        vec!["x|x|3"]
    );
}

#[test]
fn array_combine_null_and_zero_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$m = array_combine([0, '0', 1], ['zero', 'string_zero', 'one']);
echo $m[0];
echo '|';
echo $m['0'];
echo '|';
echo $m[1];
"#,
        ),
        vec!["string_zero|string_zero|one"]
    );
}

#[test]
fn array_multisort_keeps_payload_alignment_when_ties() {
    assert_eq!(
        run_prints(
            r#"<?php
$scores = [9, 9, 9];
$names = ['eve', 'adam', 'zoe'];
array_multisort($scores, SORT_DESC, $names, SORT_ASC, SORT_STRING);
echo implode('|', $names);
"#,
        ),
        vec!["adam|eve|zoe"]
    );
}

#[test]
fn array_push_pop_sequence_preserves_lifo() {
    assert_eq!(
        run_prints(
            r#"<?php
$stack = [];
array_push($stack, 1, 2);
array_push($stack, 3);
$end = array_pop($stack);
echo $end;
array_push($stack, 4);
echo $stack[0];
echo $stack[1];
echo $stack[2];
"#
        ),
        vec!["314"]
    );
}

#[test]
fn array_unshift_adds_to_front_and_returns_count() {
    assert_eq!(
        run_prints(
            r#"<?php
$items = [2, 3];
$count = array_unshift($items, 0, 1);
echo $count;
echo '|';
echo implode(',', $items);
"#
        ),
        vec!["4|0,1,2,3"]
    );
}

#[test]
fn array_shift_on_empty_array_is_null() {
    assert_eq!(
        run_prints(
            r#"<?php
$items = [];
$first = array_shift($items);
echo is_null($first) ? 'null' : 'not-null';
echo '|';
echo count($items);
"#
        ),
        vec!["null|0"]
    );
}

#[test]
fn array_pop_on_empty_array_is_null() {
    assert_eq!(
        run_prints(
            r#"<?php
$items = [];
$last = array_pop($items);
echo is_null($last) ? 'null' : 'not-null';
echo '|';
echo count($items);
"#
        ),
        vec!["null|0"]
    );
}

#[test]
fn array_reverse_reindex_variants() {
    assert_eq!(
        run_prints(
            r#"<?php
$source = ['a' => 1, 2, 'x' => 3];
$a = array_reverse($source);
$b = array_reverse($source, true);
echo implode(',', array_keys($a));
echo '|';
echo implode(',', $a);
echo '|';
echo implode(',', array_keys($b));
echo '|';
echo implode(',', $b);
"#
        ),
        vec!["1,0,2|3,2,1|x,0,a|3,2,1"]
    );
}

#[test]
fn array_splice_with_empty_replacement() {
    assert_eq!(
        run_prints(
            r#"<?php
$data = ['a' => 1, 'b' => 2, 'c' => 3, 'd' => 4];
$removed = array_splice($data, 1, 2, []);
echo count($removed);
echo '|';
echo implode(',', $data);
"#
        ),
        vec!["2|1,4"]
    );
}

#[test]
fn array_slice_offset_beyond_length_returns_empty() {
    assert_eq!(
        run_prints(
            r#"<?php
$data = [1, 2, 3];
$tail = array_slice($data, 99, 2);
echo count($tail);
echo '|';
echo json_encode($tail);
"#
        ),
        vec!["0|[]"]
    );
}

#[test]
fn array_pad_zero_length_no_change() {
    assert_eq!(
        run_prints(
            r#"<?php
$input = [1, 2, 3];
$out = array_pad($input, 0, 0);
echo count($out);
echo '|';
echo implode(',', $out);
"#
        ),
        vec!["3|1,2,3"]
    );
}

#[test]
fn array_reduce_empty_without_initial_is_null() {
    assert_eq!(
        run_prints(
            r#"<?php
$acc = array_reduce([], fn($c, $i) => $c + $i);
echo is_null($acc) ? 'null' : 'value';
"#
        ),
        vec!["null"]
    );
}

#[test]
fn array_shift_reduces_keys_preserving_first_value() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['x' => 'left', 0 => 'center', 1 => 'right'];
$first = array_shift($a);
echo $first;
echo '|';
echo count($a);
echo '|';
echo isset($a['x']) ? 'has_x' : 'no_x';
"#,
        ),
        vec!["left|2|no_x"]
    );
}

#[test]
fn array_pop_reduces_indexed_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['a' => 1, 1 => 2, 4 => 3];
$last = array_pop($a);
echo $last;
echo '|';
echo json_encode(array_keys($a));
echo '|';
echo isset($a[4]) ? 'has4' : 'no4';
"#,
        ),
        vec!["3|[\"a\",1]|no4"]
    );
}

#[test]
fn array_pop_reduces_length_for_numeric_and_assoc_mixed() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['a' => 1, 1 => 2, 2 => 3];
$last = array_pop($a);
echo $last;
echo '|';
echo count($a);
echo '|';
echo json_encode(array_keys($a));
"#,
        ),
        vec!["3|2|[\"a\",1]"]
    );
}

#[test]
fn array_flip_duplicate_values_keeps_last_key() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['a' => 'x', 'b' => 'y', 'c' => 'x', 3 => 'z'];
$r = array_flip($a);
echo $r['x'];
echo '|';
echo $r['y'];
echo '|';
echo $r['z'];
"#,
        ),
        vec!["c|b|3"]
    );
}

#[test]
fn array_count_values_counts_numeric_string_distinct() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1, '1', 1.0, '01', 1];
$counts = array_count_values($a);
ksort($counts);
echo $counts["1"] . '|' . $counts["01"];
"#,
        ),
        vec!["3|1"]
    );
}

#[test]
fn array_filter_with_key_and_value() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['x' => 1, 'y' => 2, 'z' => 3];
$b = array_filter($a, fn($v, $k) => $v > 1 && $k !== 'z', ARRAY_FILTER_USE_BOTH);
echo implode(',', array_keys($b));
"#,
        ),
        vec!["y"]
    );
}

#[test]
fn array_intersect_with_literal_string_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['01a' => 'a', '1x' => 'b', 1 => 'c'];
$b = ['1x' => 'x', 1 => 'y'];
$c = array_intersect_key($a, $b);
echo implode('|', array_keys($c));
echo '|';
echo implode(',', $c);
"#,
        ),
        vec!["1x|1|b,c"]
    );
}

#[test]
fn array_unique_keeps_first_by_default() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['x', 'x', 'y', 'x'];
$u = array_unique($a);
echo implode(',', $u);
"#,
        ),
        vec!["x,y"]
    );
}
