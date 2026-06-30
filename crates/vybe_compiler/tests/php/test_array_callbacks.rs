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
        ["24"]
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
}
