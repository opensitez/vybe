use super::helpers::run_prints;

// ── array_fill / array_fill_keys ──────────────────────────────

#[test]
fn array_fill_basic() {
    assert_eq!(
        run_prints(r#"<?php echo implode(',', array_fill(0, 5, 'x')); "#),
        vec!["x,x,x,x,x"]
    );
}
#[test]
fn array_fill_nonzero_start() {
    assert_eq!(
        run_prints(r#"<?php $a = array_fill(5, 3, 0); echo implode(',', array_keys($a)); "#),
        vec!["5,6,7"]
    );
}
#[test]
fn array_fill_keys_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = array_fill_keys(['a','b','c'], 0);
echo $a['a'] . ',' . $a['b'] . ',' . $a['c'];
"#
        ),
        vec!["0,0,0"]
    );
}
#[test]
fn array_fill_keys_with_range() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = array_fill_keys(range(1, 3), null);
echo count($a) . ':' . implode(',', array_keys($a));
"#
        ),
        vec!["3:1,2,3"]
    );
}

// ── array_pad ─────────────────────────────────────────────────

#[test]
fn array_pad_right() {
    assert_eq!(
        run_prints(r#"<?php echo implode(',', array_pad([1,2,3], 5, 0)); "#),
        vec!["1,2,3,0,0"]
    );
}
#[test]
fn array_pad_left() {
    assert_eq!(
        run_prints(r#"<?php echo implode(',', array_pad([1,2,3], -5, 0)); "#),
        vec!["0,0,1,2,3"]
    );
}
#[test]
fn array_pad_no_change_when_longer() {
    assert_eq!(
        run_prints(r#"<?php echo count(array_pad([1,2,3,4,5], 3, 0)); "#),
        vec!["5"]
    );
}

// ── array_flip ────────────────────────────────────────────────

#[test]
fn array_flip_keys_values() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = array_flip(['a'=>1,'b'=>2,'c'=>3]);
echo $a[1] . ',' . $a[2] . ',' . $a[3];
"#
        ),
        vec!["a,b,c"]
    );
}
#[test]
fn array_flip_indexed() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = array_flip(['x','y','z']);
echo $a['x'] . ',' . $a['y'] . ',' . $a['z'];
"#
        ),
        vec!["0,1,2"]
    );
}
#[test]
fn array_flip_duplicate_values_last_wins() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = array_flip(['a','b','a']);
echo $a['a'];
"#
        ),
        vec!["2"]
    );
}

// ── array_unique ──────────────────────────────────────────────

#[test]
fn array_unique_removes_duplicates() {
    assert_eq!(
        run_prints(r#"<?php echo implode(',', array_unique([1,2,2,3,3,3])); "#),
        vec!["1,2,3"]
    );
}
#[test]
fn array_unique_preserves_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = array_unique([3=>1, 5=>2, 7=>1]);
echo implode(',', array_keys($a));
"#
        ),
        vec!["3,5"]
    );
}
#[test]
fn array_unique_type_coercion() {
    assert_eq!(
        run_prints(r#"<?php echo implode(',', array_unique([1,'1',true,'true'])); "#),
        vec!["1,true"]
    );
}

// ── array_combine ─────────────────────────────────────────────

#[test]
fn array_combine_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = array_combine(['a','b','c'], [1,2,3]);
echo $a['a'] . ',' . $a['b'] . ',' . $a['c'];
"#
        ),
        vec!["1,2,3"]
    );
}

// ── array_count_values ────────────────────────────────────────

#[test]
fn array_count_values_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$c = array_count_values(['a','b','a','c','b','a']);
echo $c['a'] . ',' . $c['b'] . ',' . $c['c'];
"#
        ),
        vec!["3,2,1"]
    );
}

// ── array_sum / array_product ─────────────────────────────────

#[test]
fn array_sum_mixed_types() {
    assert_eq!(
        run_prints(r#"<?php echo array_sum([1, '2.5', true, null]); "#),
        vec!["4.5"]
    );
}
#[test]
fn array_product_integers() {
    assert_eq!(
        run_prints(r#"<?php echo array_product([1,2,3,4,5]); "#),
        vec!["120"]
    );
}
#[test]
fn array_product_with_zero() {
    assert_eq!(
        run_prints(r#"<?php echo array_product([1,2,0,4]); "#),
        vec!["0"]
    );
}

// ── range ─────────────────────────────────────────────────────

#[test]
fn range_integers() {
    assert_eq!(
        run_prints(r#"<?php echo implode(',', range(1, 5)); "#),
        vec!["1,2,3,4,5"]
    );
}
#[test]
fn range_with_step() {
    assert_eq!(
        run_prints(r#"<?php echo implode(',', range(0, 10, 2)); "#),
        vec!["0,2,4,6,8,10"]
    );
}
#[test]
fn range_descending() {
    assert_eq!(
        run_prints(r#"<?php echo implode(',', range(5, 1)); "#),
        vec!["5,4,3,2,1"]
    );
}
#[test]
fn range_chars() {
    assert_eq!(
        run_prints(r#"<?php echo implode(',', range('a', 'e')); "#),
        vec!["a,b,c,d,e"]
    );
}

// ── array_diff / array_intersect ──────────────────────────────

#[test]
fn array_diff_basic() {
    assert_eq!(
        run_prints(r#"<?php echo implode(',', array_diff([1,2,3,4,5], [2,4])); "#),
        vec!["1,3,5"]
    );
}
#[test]
fn array_intersect_basic() {
    assert_eq!(
        run_prints(r#"<?php echo implode(',', array_intersect([1,2,3,4], [2,4,6])); "#),
        vec!["2,4"]
    );
}
#[test]
fn array_diff_key_based() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['a'=>1,'b'=>2,'c'=>3];
$b = ['a'=>99,'c'=>99];
echo implode(',', array_keys(array_diff_key($a, $b)));
"#
        ),
        vec!["b"]
    );
}

#[test]
fn array_merge_reindexes_numeric_and_preserves_strings() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [0 => 'a', 3 => 'b', 'k' => 'v'];
$b = ['c', 'd'];
$m = array_merge($a, $b);
echo implode(',', $m);
echo '|';
echo $m['k'];
"#,
        ),
        vec!["a,b,v,c,d|v"]
    );
}

#[test]
fn array_intersect_assoc_strict_match() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['x' => 1, 'y' => '2', 'z' => 2];
$b = ['x' => 1, 'y' => 2, 'z' => '2'];
$r = array_intersect_assoc($a, $b);
echo implode('|', array_keys($r));
"#,
        ),
        vec!["x"]
    );
}

#[test]
fn array_udiff_with_callback() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1,2,3];
$b = ['2'];
$r = array_udiff($a, $b, fn($x, $y) => $x <=> (int)$y);
echo implode(',', $r);
"#,
        ),
        vec!["1,3"]
    );
}

#[test]
fn array_pop_from_empty_is_null() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [];
$v = array_pop($a);
echo is_null($v) ? 'null' : 'notnull';
echo '|';
echo is_array($a) ? 'is_array' : 'no';
echo '|';
echo count($a);
"#,
        ),
        vec!["null|is_array|0"]
    );
}

#[test]
fn array_shift_from_empty_is_null() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [];
$v = array_shift($a);
echo is_null($v) ? 'null' : 'notnull';
echo '|';
echo count($a);
"#,
        ),
        vec!["null|0"]
    );
}

#[test]
fn array_push_and_unshift_return_counts() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1];
$pushed = array_push($a, 2, 3);
$unshifted = array_unshift($a, 0);
echo $pushed;
echo '|';
echo $unshifted;
echo '|';
echo implode(',', $a);
"#,
        ),
        vec!["3|4|0,1,2,3"]
    );
}

#[test]
fn array_slice_length_zero_or_negative() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1, 2, 3];
$zero = array_slice($a, 1, 0);
$neg = array_slice($a, 0, -1);
echo count($zero);
echo '|';
echo implode(',', $neg);
"#,
        ),
        vec!["0|1,2"]
    );
}

#[test]
fn array_splice_replace_empty_removal_is_noop_when_count_zero() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1, 2, 3, 4];
$removed = array_splice($a, 1, 0, [9, 9]);
echo count($removed);
echo '|';
echo implode(',', $a);
"#,
        ),
        vec!["0|1,9,9,2,3,4"]
    );
}

#[test]
fn array_splice_zero_offset_negative_length_takes_tail() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1, 2, 3, 4];
$removed = array_splice($a, 0, -1, [9]);
echo count($removed);
echo '|';
echo implode(',', $a);
"#,
        ),
        vec!["3|9,4"]
    );
}

#[test]
fn array_search_with_offset_starts_after_index() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['a', 'b', 'a', 'c', 'a'];
echo array_search('a', $a, true, 2);
"#,
        ),
        vec!["4"]
    );
}

#[test]
fn array_key_exists_with_missing_and_existing_string_false_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['' => 'empty', '0' => 'zero', false => 'false-key'];
echo array_key_exists('', $a) ? 'yes' : 'no';
echo '|';
echo array_key_exists(0, $a) ? 'zero' : 'nozero';
echo '|';
echo array_key_exists('0', $a) ? 'zero2' : 'nozero2';
"#,
        ),
        vec!["yes|zero|zero2"]
    );
}

#[test]
fn array_map_with_key_mode_not_supported_keeps_single_column() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1, 2, 3];
$mapped = array_map(fn($value) => $value * 2, $a);
echo implode(',', $mapped);
"#,
        ),
        vec!["2,4,6"]
    );
}
