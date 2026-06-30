//! Core array builtins — `array_map`, `array_column`, `array_merge`, `in_array`, etc.

crate::php_cases! {
    array_map_scales_each_integer => {
        r#"<?php
echo implode(',', array_map(fn(int $n): int => $n * 10, [1, 2, 3]));
"#,
        ["10,20,30"]
    };

    array_filter_keeps_rows_matching_predicate => {
        r#"<?php
$rows = [['n' => 1], ['n' => 0], ['n' => 3]];
echo count(array_values(array_filter($rows, fn(array $r): bool => $r['n'] > 0)));
"#,
        ["2"]
    };

    array_column_extracts_single_field => {
        r#"<?php
$users = [['id' => 1, 'name' => 'ada'], ['id' => 2, 'name' => 'bob']];
echo implode('-', array_column($users, 'name'));
"#,
        ["ada-bob"]
    };

    array_column_indexes_rows_by_id => {
        r#"<?php
$users = [['id' => 10, 'name' => 'a'], ['id' => 20, 'name' => 'b']];
$byId = array_column($users, null, 'id');
echo $byId[20]['name'];
"#,
        ["b"]
    };

    array_intersect_key_keeps_only_listed_keys => {
        r#"<?php
$data = ['a' => 1, 'b' => 2, 'c' => 3];
$only = array_intersect_key($data, array_flip(['a', 'c']));
echo implode(',', array_keys($only));
"#,
        ["a,c"]
    };

    array_diff_key_removes_listed_keys => {
        r#"<?php
$data = ['a' => 1, 'b' => 2, 'c' => 3];
$except = array_diff_key($data, array_flip(['b']));
echo implode(',', array_keys($except));
"#,
        ["a,c"]
    };

    array_merge_later_keys_override_defaults => {
        r#"<?php
$defaults = ['limit' => 10, 'sort' => 'date'];
$args = ['limit' => 5];
$merged = array_merge($defaults, $args);
echo $merged['limit'] . ':' . $merged['sort'];
"#,
        ["5:date"]
    };

    in_array_strict_finds_string_member => {
        r#"<?php
$roles = ['admin', 'editor'];
echo in_array('admin', $roles, true) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    in_array_strict_rejects_loose_zero_match => {
        r#"<?php
$roles = ['admin'];
echo in_array(0, $roles, true) ? 'yes' : 'no';
"#,
        ["no"]
    };

    array_values_reindexes_hash_map => {
        r#"<?php
echo implode(',', array_values(['a' => 1, 'b' => 2]));
"#,
        ["1,2"]
    };

    array_keys_returns_index_list => {
        r#"<?php
echo implode(',', array_keys(['x' => 1, 'y' => 2]));
"#,
        ["x,y"]
    };

    array_flip_swaps_keys_and_values => {
        r#"<?php
$flipped = array_flip(['a', 'b', 'c']);
echo $flipped['b'];
"#,
        ["1"]
    };

    array_unique_preserves_first_occurrence => {
        r#"<?php
echo implode(',', array_unique([1, 2, 2, 3, 1]));
"#,
        ["1,2,3"]
    };

    array_reverse_reverses_order => {
        r#"<?php
echo implode(',', array_reverse([1, 2, 3]));
"#,
        ["3,2,1"]
    };

    array_chunk_splits_into_batches => {
        r#"<?php
$chunks = array_chunk([1, 2, 3, 4, 5], 2);
echo count($chunks) . ':' . count($chunks[1]);
"#,
        ["3:2"]
    };

    array_pad_extends_with_fill_value => {
        r#"<?php
echo implode(',', array_pad([1], 3, 0));
"#,
        ["1,0,0"]
    };

    array_sum_of_numeric_list => {
        r#"<?php
echo array_sum([1, 2, 3, 4]);
"#,
        ["10"]
    };

    array_product_multiplies_elements => {
        r#"<?php
echo array_product([2, 3, 4]);
"#,
        ["24"]
    };

    array_count_values_counts_occurrences => {
        r#"<?php
$counts = array_count_values(['a', 'b', 'a']);
echo $counts['a'];
"#,
        ["2"]
    };

    array_search_finds_key_by_value => {
        r#"<?php
echo array_search('b', ['a', 'b', 'c']);
"#,
        ["1"]
    };

    array_key_exists_distinguishes_null_value => {
        r#"<?php
$a = ['k' => null];
echo array_key_exists('k', $a) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    array_is_list_detects_sequential_keys => {
        r#"<?php
echo array_is_list([1, 2, 3]) ? 'list' : 'map';
"#,
        ["list"]
    };

    array_is_list_false_for_associative => {
        r#"<?php
echo array_is_list(['a' => 1]) ? 'list' : 'map';
"#,
        ["map"]
    };

    array_combine_pairs_keys_with_values => {
        r#"<?php
$m = array_combine(['a', 'b'], [1, 2]);
echo $m['b'];
"#,
        ["2"]
    };

    array_replace_recursive_merges_nested => {
        r#"<?php
$base = ['cfg' => ['a' => 1, 'b' => 2]];
$over = ['cfg' => ['b' => 9]];
$r = array_replace_recursive($base, $over);
echo $r['cfg']['a'] . ':' . $r['cfg']['b'];
"#,
        ["1:9"]
    };

    array_walk_mutates_by_reference => {
        r#"<?php
$a = [1, 2, 3];
array_walk($a, function (&$v) { $v *= 2; });
echo implode(',', $a);
"#,
        ["2,4,6"]
    };

    array_reduce_folds_to_single_value => {
        r#"<?php
echo array_reduce([1, 2, 3, 4], fn(int $c, int $n): int => $c + $n, 0);
"#,
        ["10"]
    };

    array_splice_replaces_middle_segment => {
        r#"<?php
$a = [1, 2, 3, 4];
array_splice($a, 1, 2, [9]);
echo implode(',', $a);
"#,
        ["1,9,4"]
    };

    array_slice_extracts_subrange => {
        r#"<?php
echo implode(',', array_slice([0, 1, 2, 3, 4], 1, 3));
"#,
        ["1,2,3"]
    };

    array_merge_recursive_combines_nested_lists => {
        r#"<?php
$r = array_merge_recursive(['k' => [1]], ['k' => [2]]);
echo implode(',', $r['k']);
"#,
        ["1,2"]
    };
}
