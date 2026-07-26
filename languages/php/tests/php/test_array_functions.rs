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

    array_fill_and_index_base => {
        r#"<?php
$a = array_fill(3, 3, 7);
echo implode(',', $a);
"#,
        ["7,7,7"]
    };

    array_flip_preserves_type_after_flip => {
        r#"<?php
$idx = array_flip(['x' => '1', 'y' => '2']);
echo $idx['1'];
echo ':';
echo $idx['2'];
"#,
        ["x:y"]
    };

    array_fill_keys_supports_multiple_keys => {
        r#"<?php
$a = array_fill_keys(['a', 'b'], 4);
echo $a['a'] . '|' . $a['b'] . '|';
echo count($a);
"#,
        ["4|4|2"]
    };

    array_slice_preserves_keys_with_preserve_key_true => {
        r#"<?php
$a = ['a' => 1, 'b' => 2, 'c' => 3, 'd' => 4];
$b = array_slice($a, 1, 2, true);
echo array_key_exists('b', $b) ? 'b' : 'nb';
echo ':';
echo array_key_exists('c', $b) ? 'c' : 'nc';
echo ':';
echo count($b);
"#,
        ["b:c:2"]
    };

    array_filter_with_flag_preserve_keys_runtime => {
        r#"<?php
$a = [0 => 'keep', 1 => '', 2 => 'x'];
$b = array_filter($a, fn($v) => $v !== '');
echo implode('|', $b);
echo ':';
echo array_key_last($b);
"#,
        ["keep|x:2"]
    };

    array_filter_with_flag_key => {
        r#"<?php
$m = ['alpha' => 1, 'beta' => 0, 'gamma' => 3];
$b = array_filter($m, fn($k) => $k !== 'beta', ARRAY_FILTER_USE_KEY);
echo implode(',', array_keys($b));
"#,
        ["alpha,gamma"]
    };

    array_filter_with_flag_both => {
        r#"<?php
$m = ['alpha' => 1, 'beta' => 0, 'gamma' => 2];
$b = array_filter($m, fn($v, $k) => $v === 0 || $k === 'gamma', ARRAY_FILTER_USE_BOTH);
echo implode(',', array_keys($b)) . '|' . implode(',', $b);
"#,
        ["beta,gamma|0,2"]
    };

    array_search_not_found_without_strict => {
        r#"<?php
$a = ['a' => 1, 'b' => '1'];
echo array_search('1', $a, false) . ':' . (array_search('2', $a, false) === false ? 'nf' : 'found');
"#,
        ["a:nf"]
    };

    array_search_not_found_with_strict => {
        r#"<?php
$a = [1, '1', 2];
echo array_search('1', $a, true) === false ? 'nf' : 'found';
"#,
        ["nf"]
    };

    array_fill_mixed_value_types => {
        r#"<?php
$a = array_fill(1, 4, ['x' => 1]);
$a[1]['x'] = 2;
echo $a[1]['x'] . '|' . $a[2]['x'];
"#,
        ["2|1"]
    };

    array_pad_truncates_only_negative_fill => {
        r#"<?php
$a = [1, 2, 3];
$b = array_pad($a, 2, 0);
echo implode(',', $b);
"#,
        ["1,2,3"]
    };

    array_replace_uses_late_argument => {
        r#"<?php
$a = ['k' => ['v' => 1], 'm' => 2];
$b = ['k' => ['other' => 9], 'n' => 3];
$c = ['k' => ['v' => 7]];
$d = array_replace($a, $b, $c);
echo json_encode($d['k']) . '|' . $d['n'];
"#,
        ["{\"v\":7,\"other\":9}|3"]
    };

    array_replace_with_multiple_sources => {
        r#"<?php
$a = ['a' => 1, 'b' => 2];
$b = ['b' => 20, 'c' => 30];
$c = ['a' => 3, 'd' => 40];
$m = array_replace($a, $b, $c);
ksort($m);
echo implode(',', array_keys($m)) . '|' . $m['a'] . '|' . $m['c'];
"#,
        ["a,b,c,d|3|30"]
    };

    array_merge_recursive_with_scalar_and_array_values => {
        r#"<?php
$x = ['k' => 1];
$y = ['k' => [2, 3]];
$r = array_merge_recursive($x, $y);
echo is_array($r['k']) ? 'array' : 'scalar';
echo '|' . $r['k'][0];
"#,
        ["array|1"]
    };

    array_diff_assoc_missing_and_extra => {
        r#"<?php
$a = ['a' => 1, 'b' => 2, 'c' => 2];
$b = ['b' => 2, 'd' => 4];
$d = array_diff_assoc($a, $b);
ksort($d);
echo implode(',', array_keys($d)) . '|' . $d['a'];
"#,
        ["a,c|1"]
    };

    array_key_exists_for_non_existent => {
        r#"<?php
echo array_key_exists('missing', ['x' => 1]) ? 'yes' : 'no';
"#,
        ["no"]
    };

    array_rand_two_keys_sorted => {
        r#"<?php
$r = array_rand(['a' => 1, 'b' => 2, 'c' => 3, 'd' => 4], 3);
sort($r);
echo implode(',', $r);
"#,
        ["a,b,c"]
    };

    array_sum_handles_float_integers => {
        r#"<?php
echo array_sum([1, 2.5, '3', 'bad']) . '|' . array_sum([true, false, null]);
"#,
        ["6.5|1"]
    };

    array_sum_empty_returns_0 => {
        r#"<?php
echo array_sum([]);
"#,
        ["0"]
    };

    array_product_empty_returns_1 => {
        r#"<?php
echo array_product([]);
"#,
        ["1"]
    };

    array_fill_with_zero_count_returns_empty => {
        r#"<?php
$a = array_fill(0, 0, 9);
echo is_array($a) ? 'array' : 'na';
echo '|';
echo count($a);
"#,
        ["array|0"]
    };

    array_fill_negative_start_throws => {
        r#"<?php
try {
    array_fill(-1, 2, 'x');
    echo 'no-error';
} catch (Throwable $e) {
    echo 'error';
}
"#,
        ["error"]
    };

    array_splice_length_beyond_end_keeps_tail => {
        r#"<?php
$a = ['a', 'b', 'c'];
array_splice($a, 1, 99, ['x', 'y']);
echo implode(',', $a);
"#,
        ["a,x,y"]
    };

    array_diff_recursive_with_nested_array => {
        r#"<?php
$a = ['a' => ['x' => 1], 'b' => ['y' => 2]];
$b = ['a' => ['x' => 1]];
$d = array_diff($a, $b, SORT_REGULAR);
echo json_encode($d);
"#,
        ["{\"b\":{\"y\":2}}"]
    };

    array_rand_singleton_returns_scalar => {
        r#"<?php
$k = array_rand(['a' => 1, 'b' => 2, 'c' => 3], 1);
echo is_string($k) ? 'scalar' : 'array';
"#,
        ["scalar"]
    };
}
