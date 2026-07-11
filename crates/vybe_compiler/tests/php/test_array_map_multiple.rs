use super::helpers::run_prints;

// ── array_map with single array ───────────────────────────────

#[test]
fn array_map_squares_each_element() {
    assert_eq!(
        run_prints(r#"<?php echo implode(',', array_map(fn($n) => $n**2, [1,2,3,4])); "#),
        vec!["1,4,9,16"]
    );
}
#[test]
fn array_map_preserves_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$r = array_map(fn($v) => $v * 10, ['a' => 1, 'b' => 2]);
echo $r['a'] . ',' . $r['b'];
"#
        ),
        vec!["10,20"]
    );
}
#[test]
fn array_map_with_named_function() {
    assert_eq!(
        run_prints(r#"<?php echo implode(',', array_map('strtoupper', ['a','b','c'])); "#),
        vec!["A,B,C"]
    );
}

// ── array_map with multiple arrays ───────────────────────────

#[test]
fn array_map_two_arrays_zips() {
    assert_eq!(
        run_prints(
            r#"<?php
$sums = array_map(fn($a,$b) => $a+$b, [1,2,3], [10,20,30]);
echo implode(',', $sums);
"#
        ),
        vec!["11,22,33"]
    );
}
#[test]
fn array_map_three_arrays() {
    assert_eq!(
        run_prints(
            r#"<?php
$r = array_map(fn($a,$b,$c) => $a+$b+$c, [1,2], [3,4], [5,6]);
echo implode(',', $r);
"#
        ),
        vec!["9,12"]
    );
}
#[test]
fn array_map_unequal_length_pads_null() {
    assert_eq!(
        run_prints(
            r#"<?php
$r = array_map(fn($a,$b) => "$a:$b", [1,2,3], [10,20]);
echo implode(',', $r);
"#
        ),
        vec!["1:10,2:20,3:"]
    );
}

// ── array_map with null callback — zip ────────────────────────

#[test]
fn array_map_null_callback_zips_arrays() {
    assert_eq!(
        run_prints(
            r#"<?php
$zipped = array_map(null, [1,2,3], ['a','b','c']);
echo count($zipped) . ',' . implode(',', $zipped[1]);
"#
        ),
        vec!["3,2,b"]
    );
}
#[test]
fn array_map_null_single_array_wraps_in_array() {
    assert_eq!(
        run_prints(
            r#"<?php
$r = array_map(null, [1,2,3]);
echo implode(',', array_map(fn($x) => $x[0], $r));
"#
        ),
        vec![",,"]
    );
}

// ── array_filter ───────────────────────────────────────────────

#[test]
fn array_filter_removes_falsy_by_default() {
    assert_eq!(
        run_prints(
            r#"<?php
$r = array_filter([0, 1, '', 'a', null, false, true]);
echo implode(',', $r);
"#
        ),
        vec!["1,a,1"]
    );
}
#[test]
fn array_filter_with_callback() {
    assert_eq!(
        run_prints(
            r#"<?php
$evens = array_filter([1,2,3,4,5,6], fn($n) => $n % 2 === 0);
echo implode(',', $evens);
"#
        ),
        vec!["2,4,6"]
    );
}
#[test]
fn array_filter_preserves_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$r = array_filter([10,20,30,40], fn($v) => $v > 15);
echo implode(',', array_keys($r));
"#
        ),
        vec!["1,2,3"]
    );
}
#[test]
fn array_filter_mode_key() {
    assert_eq!(
        run_prints(
            r#"<?php
$arr = ['a' => 1, 'b' => 2, 'c' => 3];
$r = array_filter($arr, fn($k) => $k !== 'b', ARRAY_FILTER_USE_KEY);
echo implode(',', array_keys($r));
"#
        ),
        vec!["a,c"]
    );
}
#[test]
fn array_filter_mode_both() {
    assert_eq!(
        run_prints(
            r#"<?php
$r = array_filter(['x' => 1, 'y' => 2], fn($v,$k) => $k === 'x' || $v > 1, ARRAY_FILTER_USE_BOTH);
echo implode(',', array_keys($r));
"#
        ),
        vec!["x,y"]
    );
}

// ── array_reduce ───────────────────────────────────────────────

#[test]
fn array_reduce_sum() {
    assert_eq!(
        run_prints(r#"<?php echo array_reduce([1,2,3,4,5], fn($carry,$v) => $carry+$v, 0); "#),
        vec!["15"]
    );
}
#[test]
fn array_reduce_product() {
    assert_eq!(
        run_prints(r#"<?php echo array_reduce([1,2,3,4], fn($c,$v) => $c*$v, 1); "#),
        vec!["24"]
    );
}
#[test]
fn array_reduce_string_concat() {
    assert_eq!(
        run_prints(r#"<?php echo array_reduce(['a','b','c'], fn($c,$v) => $c.$v, ''); "#),
        vec!["abc"]
    );
}
#[test]
fn array_reduce_build_array() {
    assert_eq!(
        run_prints(
            r#"<?php
$r = array_reduce([1,2,3], fn($c,$v) => array_merge($c, [$v*2]), []);
echo implode(',', $r);
"#
        ),
        vec!["2,4,6"]
    );
}
#[test]
fn array_reduce_initial_null() {
    assert_eq!(
        run_prints(
            r#"<?php
$r = array_reduce([], fn($c,$v) => $c + $v, 42);
echo $r;
"#
        ),
        vec!["42"]
    );
}

// ── array_walk ────────────────────────────────────────────────

#[test]
fn array_walk_modifies_in_place() {
    assert_eq!(
        run_prints(
            r#"<?php
$arr = [1,2,3];
array_walk($arr, function(&$v, $k) { $v = $v * 10; });
echo implode(',', $arr);
"#
        ),
        vec!["10,20,30"]
    );
}
#[test]
fn array_walk_receives_extra_param() {
    assert_eq!(
        run_prints(
            r#"<?php
$arr = ['a','b','c'];
array_walk($arr, function(&$v, $k, $prefix) { $v = $prefix . $v; }, 'X');
echo implode(',', $arr);
"#
        ),
        vec!["Xa,Xb,Xc"]
    );
}
#[test]
fn array_walk_recursive_nested() {
    assert_eq!(
        run_prints(
            r#"<?php
$arr = [1, [2, 3], [4, [5]]];
$sum = 0;
array_walk_recursive($arr, function($v) use (&$sum) { $sum += $v; });
echo $sum;
"#
        ),
        vec!["15"]
    );
}
