use super::helpers::run_prints;

// ── array_fill / array_fill_keys ──────────────────────────────

#[test] fn array_fill_basic() {
    assert_eq!(run_prints(r#"<?php echo implode(',', array_fill(0, 5, 'x')); "#), vec!["x,x,x,x,x"]);
}
#[test] fn array_fill_nonzero_start() {
    assert_eq!(run_prints(r#"<?php $a = array_fill(5, 3, 0); echo implode(',', array_keys($a)); "#), vec!["5,6,7"]);
}
#[test] fn array_fill_keys_basic() {
    assert_eq!(run_prints(r#"<?php
$a = array_fill_keys(['a','b','c'], 0);
echo $a['a'] . ',' . $a['b'] . ',' . $a['c'];
"#), vec!["0,0,0"]);
}
#[test] fn array_fill_keys_with_range() {
    assert_eq!(run_prints(r#"<?php
$a = array_fill_keys(range(1, 3), null);
echo count($a) . ':' . implode(',', array_keys($a));
"#), vec!["3:1,2,3"]);
}

// ── array_pad ─────────────────────────────────────────────────

#[test] fn array_pad_right() {
    assert_eq!(run_prints(r#"<?php echo implode(',', array_pad([1,2,3], 5, 0)); "#), vec!["1,2,3,0,0"]);
}
#[test] fn array_pad_left() {
    assert_eq!(run_prints(r#"<?php echo implode(',', array_pad([1,2,3], -5, 0)); "#), vec!["0,0,1,2,3"]);
}
#[test] fn array_pad_no_change_when_longer() {
    assert_eq!(run_prints(r#"<?php echo count(array_pad([1,2,3,4,5], 3, 0)); "#), vec!["5"]);
}

// ── array_flip ────────────────────────────────────────────────

#[test] fn array_flip_keys_values() {
    assert_eq!(run_prints(r#"<?php
$a = array_flip(['a'=>1,'b'=>2,'c'=>3]);
echo $a[1] . ',' . $a[2] . ',' . $a[3];
"#), vec!["a,b,c"]);
}
#[test] fn array_flip_indexed() {
    assert_eq!(run_prints(r#"<?php
$a = array_flip(['x','y','z']);
echo $a['x'] . ',' . $a['y'] . ',' . $a['z'];
"#), vec!["0,1,2"]);
}
#[test] fn array_flip_duplicate_values_last_wins() {
    assert_eq!(run_prints(r#"<?php
$a = array_flip(['a','b','a']);
echo $a['a'];
"#), vec!["2"]);
}

// ── array_unique ──────────────────────────────────────────────

#[test] fn array_unique_removes_duplicates() {
    assert_eq!(run_prints(r#"<?php echo implode(',', array_unique([1,2,2,3,3,3])); "#), vec!["1,2,3"]);
}
#[test] fn array_unique_preserves_keys() {
    assert_eq!(run_prints(r#"<?php
$a = array_unique([3=>1, 5=>2, 7=>1]);
echo implode(',', array_keys($a));
"#), vec!["3,5"]);
}
#[test] fn array_unique_type_coercion() {
    assert_eq!(run_prints(r#"<?php echo implode(',', array_unique([1,'1',true,'true'])); "#), vec!["1,true"]);
}

// ── array_combine ─────────────────────────────────────────────

#[test] fn array_combine_basic() {
    assert_eq!(run_prints(r#"<?php
$a = array_combine(['a','b','c'], [1,2,3]);
echo $a['a'] . ',' . $a['b'] . ',' . $a['c'];
"#), vec!["1,2,3"]);
}

// ── array_count_values ────────────────────────────────────────

#[test] fn array_count_values_basic() {
    assert_eq!(run_prints(r#"<?php
$c = array_count_values(['a','b','a','c','b','a']);
echo $c['a'] . ',' . $c['b'] . ',' . $c['c'];
"#), vec!["3,2,1"]);
}

// ── array_sum / array_product ─────────────────────────────────

#[test] fn array_sum_mixed_types() {
    assert_eq!(run_prints(r#"<?php echo array_sum([1, '2.5', true, null]); "#), vec!["4.5"]);
}
#[test] fn array_product_integers() {
    assert_eq!(run_prints(r#"<?php echo array_product([1,2,3,4,5]); "#), vec!["120"]);
}
#[test] fn array_product_with_zero() {
    assert_eq!(run_prints(r#"<?php echo array_product([1,2,0,4]); "#), vec!["0"]);
}

// ── range ─────────────────────────────────────────────────────

#[test] fn range_integers() {
    assert_eq!(run_prints(r#"<?php echo implode(',', range(1, 5)); "#), vec!["1,2,3,4,5"]);
}
#[test] fn range_with_step() {
    assert_eq!(run_prints(r#"<?php echo implode(',', range(0, 10, 2)); "#), vec!["0,2,4,6,8,10"]);
}
#[test] fn range_descending() {
    assert_eq!(run_prints(r#"<?php echo implode(',', range(5, 1)); "#), vec!["5,4,3,2,1"]);
}
#[test] fn range_chars() {
    assert_eq!(run_prints(r#"<?php echo implode(',', range('a', 'e')); "#), vec!["a,b,c,d,e"]);
}

// ── array_diff / array_intersect ──────────────────────────────

#[test] fn array_diff_basic() {
    assert_eq!(run_prints(r#"<?php echo implode(',', array_diff([1,2,3,4,5], [2,4])); "#), vec!["1,3,5"]);
}
#[test] fn array_intersect_basic() {
    assert_eq!(run_prints(r#"<?php echo implode(',', array_intersect([1,2,3,4], [2,4,6])); "#), vec!["2,4"]);
}
#[test] fn array_diff_key_based() {
    assert_eq!(run_prints(r#"<?php
$a = ['a'=>1,'b'=>2,'c'=>3];
$b = ['a'=>99,'c'=>99];
echo implode(',', array_keys(array_diff_key($a, $b)));
"#), vec!["b"]);
}
