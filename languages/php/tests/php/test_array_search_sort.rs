//! Array search, column, unique, flip — distinct from `test_array_functions.rs`.

crate::php_cases! {
    array_column_simple_list => {
        r#"<?php
$rows = [['id' => 1, 'n' => 'a'], ['id' => 2, 'n' => 'b']];
echo implode(',', array_column($rows, 'n'));
"#,
        ["a,b"]
    };

    array_column_with_index_key => {
        r#"<?php
$rows = [['id' => 1, 'n' => 'a'], ['id' => 2, 'n' => 'b']];
$c = array_column($rows, 'n', 'id');
echo $c[2];
"#,
        ["b"]
    };

    array_unique_preserves_first_key => {
        r#"<?php
$a = array_unique([1, 1, 2, 2]);
echo implode(',', array_values($a));
"#,
        ["1,2"]
    };

    array_unique_string_sort => {
        r#"<?php
$a = array_unique(['b', 'a', 'b'], SORT_STRING);
echo count($a);
"#,
        ["2"]
    };

    array_flip_swaps_keys_values => {
        r#"<?php
$f = array_flip(['a' => 1, 'b' => 2]);
echo $f[1];
"#,
        ["a"]
    };

    array_keys_default => {
        r#"<?php
echo implode(',', array_keys(['x' => 1, 'y' => 2]));
"#,
        ["x,y"]
    };

    array_keys_with_value_filter => {
        r#"<?php
echo implode(',', array_keys(['a' => 1, 'b' => 0], 1));
"#,
        ["a"]
    };

    array_values_reindexes => {
        r#"<?php
echo implode(',', array_values(['a' => 1, 'b' => 2]));
"#,
        ["1,2"]
    };

    array_search_strict => {
        r#"<?php
echo array_search(1, ['a', 1, 2], true);
"#,
        ["1"]
    };

    array_search_not_found => {
        r#"<?php
echo array_search(9, [1, 2, 3]) === false ? 'no' : 'yes';
"#,
        ["no"]
    };

    array_key_exists_true => {
        r#"<?php
echo array_key_exists('k', ['k' => null]) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    array_key_exists_numeric_string => {
        r#"<?php
echo array_key_exists(0, ['0' => 'x']) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    in_array_strict_false_for_string_one => {
        r#"<?php
echo in_array('1', [1, 2, 3], true) ? 'yes' : 'no';
"#,
        ["no"]
    };

    in_array_loose_true => {
        r#"<?php
echo in_array('1', [1, 2, 3]) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    array_key_first => {
        r#"<?php
echo array_key_first(['a' => 1, 'b' => 2]);
"#,
        ["a"]
    };

    array_key_last => {
        r#"<?php
echo array_key_last(['a' => 1, 'b' => 2]);
"#,
        ["b"]
    };

    array_chunk_two_columns => {
        r#"<?php
$c = array_chunk([1, 2, 3, 4], 2);
echo count($c);
"#,
        ["2"]
    };

    array_chunk_preserve_keys => {
        r#"<?php
$c = array_chunk(['a' => 1, 'b' => 2, 'c' => 3], 2, true);
echo array_key_first($c[0]);
"#,
        ["a"]
    };

    array_pad_positive => {
        r#"<?php
echo implode(',', array_pad([1], 3, 0));
"#,
        ["1,0,0"]
    };

    array_pad_negative => {
        r#"<?php
$p = array_pad([1], -3, 0);
echo $p[0];
"#,
        ["0"]
    };

    array_fill_keys => {
        r#"<?php
$a = array_fill_keys(['a', 'b'], 1);
echo $a['b'];
"#,
        ["1"]
    };

    array_fill_range => {
        r#"<?php
echo implode(',', array_fill(0, 3, 'x'));
"#,
        ["x,x,x"]
    };

    array_replace_recursive => {
        r#"<?php
$a = ['k' => ['a' => 1]];
$b = ['k' => ['b' => 2]];
$r = array_replace_recursive($a, $b);
echo $r['k']['a'] . $r['k']['b'];
"#,
        ["12"]
    };

    array_merge_recursive => {
        r#"<?php
$r = array_merge_recursive(['k' => [1]], ['k' => [2]]);
echo implode(',', $r['k']);
"#,
        ["1,2"]
    };

    array_intersect_assoc => {
        r#"<?php
$a = ['a' => 1, 'b' => 2];
$b = ['a' => 1, 'b' => 9];
echo count(array_intersect_assoc($a, $b));
"#,
        ["1"]
    };

    array_diff_assoc => {
        r#"<?php
$a = ['a' => 1, 'b' => 2];
$b = ['a' => 1, 'b' => 9];
echo implode(',', array_keys(array_diff_assoc($a, $b)));
"#,
        ["b"]
    };

    array_udiff_user_callback => {
        r#"<?php
$cmp = fn($a, $b) => $a <=> $b;
echo count(array_udiff([1, 2], [2], $cmp));
"#,
        ["1"]
    };

    array_uintersect_user_callback => {
        r#"<?php
$cmp = fn($a, $b) => $a <=> $b;
echo count(array_uintersect([1, 2], [2, 3], $cmp));
"#,
        ["1"]
    };

    array_walk_recursive_sum => {
        r#"<?php
$a = ['x' => [1, 2], 'y' => [3]];
$s = 0;
array_walk_recursive($a, function ($v) use (&$s) { $s += $v; });
echo $s;
"#,
        ["6"]
    };

    array_change_key_case_lower => {
        r#"<?php
$a = array_change_key_case(['A' => 1]);
echo array_key_first($a);
"#,
        ["a"]
    };

    array_reverse_preserve_keys => {
        r#"<?php
$a = array_reverse(['a' => 1, 'b' => 2], true);
echo array_key_first($a);
"#,
        ["b"]
    };

    array_slice_negative_offset => {
        r#"<?php
echo implode(',', array_slice([1, 2, 3, 4], -2));
"#,
        ["3,4"]
    };

    array_splice_replace_with_list => {
        r#"<?php
$a = [1, 2, 3, 4];
array_splice($a, 1, 2, [9, 9]);
echo implode(',', $a);
"#,
        ["1,9,9,4"]
    };

    array_combine_keys_values => {
        r#"<?php
$c = array_combine(['a', 'b'], [1, 2]);
echo $c['a'];
"#,
        ["1"]
    };

    array_count_values_frequency => {
        r#"<?php
$c = array_count_values(['a', 'a', 'b']);
echo $c['a'];
"#,
        ["2"]
    };

    array_product_of_numbers => {
        r#"<?php
echo array_product([2, 3, 4]);
"#,
        ["24"]
    };

    array_sum_floats => {
        r#"<?php
echo array_sum([1.5, 2.5]);
"#,
        ["4"]
    };

    array_search_loose_prefers_first_match => {
        r#"<?php
echo array_search('2', [1, '2', 2, '2'], false);
"#,
        ["1"]
    };

    array_search_strict_with_type_mismatch => {
        r#"<?php
echo array_search('2', [1, '2', 2, '2'], true) === false ? 'no' : 'yes';
"#,
        ["no"]
    };

    in_array_with_default_mode_skips_type => {
        r#"<?php
echo in_array(0, [false, '0', null], false) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    array_key_exists_false_for_missing_key => {
        r#"<?php
echo array_key_exists('missing', ['a' => 1, 'b' => 2]) ? 'yes' : 'no';
"#,
        ["no"]
    };

    array_rand_multiple_keys => {
        r#"<?php
$r = array_rand(['a' => 1, 'b' => 2, 'c' => 3], 2);
sort($r);
echo count($r) . '|' . implode(',', $r);
"#,
        ["2|a,b"]
    };

    array_unique_preserves_first_numeric_index => {
        r#"<?php
$u = array_unique([10 => 'a', 11 => 'b', 12 => 'a']);
echo array_key_first($u) . ':' . implode(',', array_values($u));
"#,
        ["10|a,b"]
    };

    array_slice_preserve_keys_false_reindex => {
        r#"<?php
$s = ['a' => 1, 'b' => 2, 'c' => 3];
$t = array_slice($s, 1, 2, false);
echo implode(',', array_keys($t)) . ':' . implode(',', $t);
"#,
        ["0,1|2,3"]
    };

    array_rand_one_key_returns_scalar => {
        r#"<?php
$r = array_rand(['only' => 1], 1);
echo is_array($r) ? 'array' : 'scalar';
if (!is_array($r)) echo '|' . $r;
"#,
        ["scalar|only"]
    };

    array_key_first_empty => {
        r#"<?php
$first = array_key_first([]);
echo is_null($first) ? 'null' : 'value';
"#,
        ["null"]
    };

    array_key_last_empty => {
        r#"<?php
$last = array_key_last([]);
echo is_null($last) ? 'null' : 'value';
"#,
        ["null"]
    };

    array_search_prefers_last_matching_key_with_loose_and_strict_mix => {
        r#"<?php
$r = array_search('2', [1, '2', 2, '2']);
echo is_int($r) ? $r : gettype($r);
"#,
        ["1"]
    };

    array_fill_with_zero_count => {
        r#"<?php
echo count(array_fill(0, 0, 9)) . '|' . json_encode(array_fill(1, 0, 'x'));
"#,
        ["0|[]"]
    };

    array_search_offset_starts_after_index => {
        r#"<?php
echo array_search('2', ['2', 3, '2', 4], false, 2);
"#,
        ["2"]
    };

    array_search_null_needle_strict => {
        r#"<?php
$a = [null, 0, ''];
echo array_search(null, $a, true);
"#,
        ["0"]
    };

    in_array_on_nested_array_loose => {
        r#"<?php
echo in_array([1, 2], [[1, 2], [3, 4]]) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    in_array_on_nested_array_strict => {
        r#"<?php
echo in_array([1, '2'], [[1, 2], [3, 4]], true) ? 'yes' : 'no';
"#,
        ["no"]
    };

    array_key_exists_on_numeric_string_key_collisions => {
        r#"<?php
$a = ['0' => 'zero', 0 => 'int0', '01' => 'leading'];
echo array_key_exists(0, $a) . '|' . array_key_exists('0', $a) . '|' . array_key_exists('01', $a);
"#,
        ["1|1|1"]
    };

    array_flip_duplicate_values_last_value_keeps_last_key => {
        r#"<?php
$f = array_flip(['a', 'b', 'a', 'c']);
echo $f['a'];
"#,
        ["3"]
    };

    array_merge_recursive_with_scalar_and_array => {
        r#"<?php
$x = ['a' => [1], 'b' => 2];
$y = ['a' => 3];
$m = array_merge_recursive($x, $y);
echo is_array($m['a']) ? implode('|', $m['a']) : 'not-array';
"#,
        ["1|3"]
    };

    array_values_empty_input => {
        r#"<?php
$v = array_values([]);
echo json_encode($v);
"#,
        ["[]"]
    };

    array_keys_search_value_strict => {
        r#"<?php
$a = ['a' => '1', 'b' => 1, 'c' => '1'];
echo implode('|', array_keys($a, '1', true));
"#,
        ["b"]
    };
}
