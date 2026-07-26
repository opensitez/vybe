//! `array_map`, `array_filter`, `array_reduce`, `array_walk`, and callback patterns.

crate::php_cases! {
    array_map_doubles_each_element => {
        r#"<?php
echo implode(',', array_map(fn($n) => $n * 2, [1, 2, 3]));
"#,
        ["2,4,6"]
    };

    array_map_with_multiple_arrays_zips => {
        r#"<?php
echo implode(',', array_map(fn($a, $b) => $a + $b, [1, 2], [10, 20]));
"#,
        ["11,22"]
    };

    array_filter_keeps_truthy_values => {
        r#"<?php
echo implode(',', array_filter([0, 1, '', 'a', null]));
"#,
        ["1,a"]
    };

    array_filter_with_callback_keeps_evens => {
        r#"<?php
echo implode(',', array_filter([1, 2, 3, 4], fn($n) => $n % 2 === 0));
"#,
        ["2,4"]
    };

    array_filter_use_both_value_and_key => {
        r#"<?php
echo implode(',', array_filter(['a' => 1, 'b' => 0, 'c' => 3], fn($v, $k) => $k !== 'b'));
"#,
        ["1,3"]
    };

    array_reduce_sums_with_initial_zero => {
        r#"<?php
echo array_reduce([1, 2, 3], fn($c, $n) => $c + $n, 0);
"#,
        ["6"]
    };

    array_reduce_without_initial_uses_first => {
        r#"<?php
echo array_reduce([2, 3, 4], fn($c, $n) => $c * $n);
"#,
        ["0"]
    };

    array_walk_mutates_by_reference => {
        r#"<?php
$a = [1, 2, 3];
array_walk($a, function (&$v) { $v *= 10; });
echo implode(',', $a);
"#,
        ["10,20,30"]
    };

    array_walk_recursive_nested_multiply => {
        r#"<?php
$a = ['x' => [1, 2], 'y' => 3];
array_walk_recursive($a, function (&$v) { if (is_int($v)) $v++; });
echo $a['x'][0] . ':' . $a['y'];
"#,
        ["2:4"]
    };

    array_column_extracts_field => {
        r#"<?php
$rows = [['id' => 1, 'n' => 'a'], ['id' => 2, 'n' => 'b']];
echo implode(',', array_column($rows, 'n'));
"#,
        ["a,b"]
    };

    array_column_with_index_key => {
        r#"<?php
$rows = [['id' => 10, 'v' => 'x'], ['id' => 20, 'v' => 'y']];
$map = array_column($rows, 'v', 'id');
echo $map[10] . $map[20];
"#,
        ["xy"]
    };

    array_combine_builds_map => {
        r#"<?php
$m = array_combine(['a', 'b'], [1, 2]);
echo $m['a'] . ':' . $m['b'];
"#,
        ["1:2"]
    };

    array_flip_swaps_keys_values => {
        r#"<?php
$f = array_flip(['a' => 1, 'b' => 2]);
echo $f[1] . $f[2];
"#,
        ["ab"]
    };

    array_unique_preserves_first_key => {
        r#"<?php
$u = array_unique([1, 1, 2, 2, 3]);
echo implode(',', $u);
"#,
        ["1,2,3"]
    };

    array_values_reindexes => {
        r#"<?php
echo implode(',', array_values(['a' => 1, 'b' => 2]));
"#,
        ["1,2"]
    };

    array_keys_lists_indexes => {
        r#"<?php
echo implode(',', array_keys(['x' => 1, 'y' => 2]));
"#,
        ["x,y"]
    };

    array_merge_appends_lists => {
        r#"<?php
echo implode(',', array_merge([1, 2], [3]));
"#,
        ["1,2,3"]
    };

    array_replace_overwrites_by_key => {
        r#"<?php
echo json_encode(array_replace(['a' => 1, 'b' => 2], ['b' => 9]));
"#,
        ["{\"a\":1,\"b\":9}"]
    };

    array_chunk_splits_with_keys_preserved => {
        r#"<?php
$chunks = array_chunk(['a' => 1, 'b' => 2, 'c' => 3], 2, true);
echo count($chunks) . ':' . $chunks[1]['c'];
"#,
        ["2:3"]
    };

    array_pad_appends_to_length => {
        r#"<?php
echo implode(',', array_pad([1], 4, 0));
"#,
        ["1,0,0,0"]
    };

    array_slice_extracts_middle => {
        r#"<?php
echo implode(',', array_slice([0, 1, 2, 3, 4], 1, 3));
"#,
        ["1,2,3"]
    };

    array_diff_compares_values => {
        r#"<?php
echo implode(',', array_diff([1, 2, 3], [2]));
"#,
        ["1,3"]
    };

    array_intersect_keeps_common_values => {
        r#"<?php
echo implode(',', array_intersect([1, 2, 3], [2, 3, 4]));
"#,
        ["2,3"]
    };

    array_key_exists_vs_isset_null => {
        r#"<?php
$a = ['k' => null];
echo (array_key_exists('k', $a) ? '1' : '0') . (isset($a['k']) ? '1' : '0');
"#,
        ["10"]
    };

    array_search_finds_value_key => {
        r#"<?php
echo array_search(2, ['a' => 1, 'b' => 2]);
"#,
        ["b"]
    };

    array_map_captures_outer_sum => {
        r#"<?php
$numbers = [1, 2, 3];
$base = 10;
$mapped = array_map(fn($n) => $n + $base, $numbers);
echo implode(',', $mapped);
"#,
        ["11,12,13"]
    };

    array_filter_use_key_mode => {
        r#"<?php
$filtered = array_filter(['a' => 1, 'bb' => 2, 'ccc' => 3], fn($k) => strlen($k) > 1, ARRAY_FILTER_USE_KEY);
echo implode(',', array_keys($filtered));
"#,
        ["bb,ccc"]
    };

    array_filter_use_both_key_and_value => {
        r#"<?php
$filtered = array_filter(
    ['a' => 1, 'b' => 0, 'cc' => 3],
    fn($v, $k) => $v === 1 || $k === 'cc',
    ARRAY_FILTER_USE_BOTH
);
echo implode(',', array_keys($filtered));
"#,
        ["a,cc"]
    };

    array_map_with_null_callback_zip => {
        r#"<?php
$letters = ['a', 'b', 'c'];
$numbers = [1, 2, 3];
$zipped = array_map(null, $letters, $numbers);
echo json_encode($zipped);
"#,
        ["[[\"a\",1],[\"b\",2],[\"c\",3]]"]
    };

    array_walk_with_key_and_data => {
        r#"<?php
$items = ['x' => 1, 'y' => 2];
$factor = 5;
array_walk($items, function(&$v, $k, $factor) {
    $v *= $factor;
    if ($k === 'x') {
        $v += 1;
    }
}, $factor);
echo implode('|', $items);
"#,
        ["6|10"]
    };

    array_reduce_initial_non_trivial_seed => {
        r#"<?php
$result = array_reduce([1, 2, 3], fn($c, $n) => $c . ':' . $n, 'seed');
echo $result;
"#,
        ["seed:1:2:3"]
    };

    array_map_with_empty_input_is_empty => {
        r#"<?php
echo json_encode(array_map(fn($n) => $n * 2, []));
"#,
        ["[]"]
    };

    array_map_truncates_to_shortest_input => {
        r#"<?php
$a = [1, 2, 3];
$b = [10, 20];
$mapped = array_map(fn($x, $y) => $x + $y, $a, $b);
echo json_encode($mapped);
"#,
        ["[11,22]"]
    };

    array_filter_default_keeps_only_truthy => {
        r#"<?php
echo implode(',', array_filter([0, 1, false, 'php', '', 5, null, '0']));
"#,
        ["1,php,5"]
    };

    array_filter_key_callback_skips_numeric_keys => {
        r#"<?php
$items = ['0' => 1, 'a' => 2, 3 => 3, 'b' => 4];
$filtered = array_filter($items, fn($k) => ctype_alpha((string)$k), ARRAY_FILTER_USE_KEY);
echo implode(',', array_keys($filtered));
"#,
        ["a,b"]
    };

    array_filter_with_empty_callback_and_both_modes => {
        r#"<?php
$items = ['x' => 0, 'y' => 1, 'z' => 2];
$filtered = array_filter(
    $items,
    fn($v, $k) => $v > 0 || $k === 'x',
    ARRAY_FILTER_USE_BOTH
);
echo implode(',', array_keys($filtered));
"#,
        ["x,z"]
    };

    array_reduce_accumulates_objects_by_reference => {
        r#"<?php
$items = [[1], [2], [3]];
$result = array_reduce($items, function($carry, $item) {
    $carry[] = $item[0] * 2;
    return $carry;
}, []);
echo implode(',', $result);
"#,
        ["2,4,6"]
    };

    array_walk_recursive_preserves_nested_keys => {
        r#"<?php
$tree = ['a' => ['count' => 1], 'b' => ['count' => 2]];
array_walk_recursive($tree, function(&$value, $key) {
    if (is_numeric($value)) {
        $value += 1;
    }
});
echo $tree['a']['count'];
echo $tree['b']['count'];
"#,
        ["23"]
    };

    array_map_with_null_preserves_shape => {
        r#"<?php
$out = array_map(null, [1, 2], ['a', 'b', 'c']);
echo json_encode($out);
"#,
        ["[[1,\"a\"],[2,\"b\"]]"]
    };

    array_walk_with_data_argument => {
        r#"<?php
$items = ['x' => 1, 'y' => 2];
array_walk($items, function(&$v, $k, $scale) {
    $v = $v * $scale;
}, 3);
echo implode('|', $items);
"#,
        ["3|6"]
    };
}
