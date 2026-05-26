use super::helpers::run_prints;

// ── var_export ────────────────────────────────────────────────

#[test] fn var_export_int() {
    assert_eq!(run_prints(r#"<?php var_export(42); "#), vec!["42"]);
}
#[test] fn var_export_string() {
    assert_eq!(run_prints(r#"<?php var_export('hello'); "#), vec!["'hello'"]);
}
#[test] fn var_export_array() {
    assert_eq!(run_prints(r#"<?php var_export([1,2,3]); "#), vec!["array (\n  0 => 1,\n  1 => 2,\n  2 => 3,\n)"]);
}
#[test] fn var_export_bool() {
    assert_eq!(run_prints(r#"<?php var_export(true); echo ','; var_export(false); "#), vec!["true,false"]);
}
#[test] fn var_export_null() {
    assert_eq!(run_prints(r#"<?php var_export(null); "#), vec!["NULL"]);
}
#[test] fn var_export_return_true() {
    assert_eq!(run_prints(r#"<?php $s = var_export(42, true); echo gettype($s) . ':' . $s; "#), vec!["string:42"]);
}

// ── print_r ───────────────────────────────────────────────────

#[test] fn print_r_simple() {
    assert_eq!(run_prints(r#"<?php print_r('hello'); "#), vec!["hello"]);
}
#[test] fn print_r_array_return() {
    assert_eq!(run_prints(r#"<?php $s = print_r([1,2], true); echo str_contains($s, 'Array') ? 'ok' : 'fail'; "#), vec!["ok"]);
}

// ── isset / empty / unset ────────────────────────────────────

#[test] fn isset_multiple_vars() {
    assert_eq!(run_prints(r#"<?php $a = 1; $b = 2; echo isset($a, $b) ? 'yes' : 'no'; echo isset($a, $c) ? 'yes' : 'no'; "#), vec!["yesno"]);
}
#[test] fn empty_various() {
    assert_eq!(run_prints(r#"<?php
echo empty('') ? '1' : '0';
echo empty(0) ? '1' : '0';
echo empty([]) ? '1' : '0';
echo empty('hello') ? '1' : '0';
echo empty(1) ? '1' : '0';
"#), vec!["11100"]);
}
#[test] fn unset_removes_var() {
    assert_eq!(run_prints(r#"<?php $a = 42; unset($a); echo isset($a) ? 'yes' : 'no'; "#), vec!["no"]);
}
#[test] fn unset_array_key() {
    assert_eq!(run_prints(r#"<?php $a = [1,2,3]; unset($a[1]); echo implode(',', $a); "#), vec!["1,3"]);
}

// ── list() / [] assignment edge cases ────────────────────────

#[test] fn list_from_string_split() {
    assert_eq!(run_prints(r#"<?php [$a,$b,$c] = explode(',', 'x,y,z'); echo $a.$b.$c; "#), vec!["xyz"]);
}
#[test] fn list_with_extra_elements() {
    assert_eq!(run_prints(r#"<?php [$a,$b] = [1,2,3,4,5]; echo $a . ',' . $b; "#), vec!["1,2"]);
}

// ── PHP type juggling with comparisons ───────────────────────

#[test] fn spaceship_with_strings_alphabetical() {
    assert_eq!(run_prints(r#"<?php echo ('b' <=> 'a') . ',' . ('a' <=> 'b') . ',' . ('a' <=> 'a'); "#), vec!["1,-1,0"]);
}
#[test] fn comparison_with_null() {
    assert_eq!(run_prints(r#"<?php echo (null < 0) ? 'lt' : 'gte'; echo (null > 0) ? 'gt' : 'lte'; "#), vec!["ltlte"]);
}

// ── Date functions ────────────────────────────────────────────

#[test] fn date_format_year() {
    assert_eq!(run_prints(r#"<?php echo date('Y', mktime(0,0,0,1,1,2024)); "#), vec!["2024"]);
}
#[test] fn date_format_month_day() {
    assert_eq!(run_prints(r#"<?php echo date('m-d', mktime(0,0,0,7,15,2024)); "#), vec!["07-15"]);
}
#[test] fn mktime_returns_timestamp() {
    assert_eq!(run_prints(r#"<?php echo mktime(0,0,0,1,1,1970) >= 0 ? 'pos_or_zero' : 'neg'; "#), vec!["pos_or_zero"]);
}
#[test] fn time_returns_positive_int() {
    assert_eq!(run_prints(r#"<?php echo time() > 0 ? 'ok' : 'fail'; "#), vec!["ok"]);
}
#[test] fn microtime_float() {
    assert_eq!(run_prints(r#"<?php echo microtime(true) > 0 ? 'ok' : 'fail'; "#), vec!["ok"]);
}

// ── Array creation shortcuts ──────────────────────────────────

#[test] fn range_with_float_step() {
    assert_eq!(run_prints(r#"<?php echo implode(',', range(0, 1, 0.25)); "#), vec!["0,0.25,0.5,0.75,1"]);
}
#[test] fn array_fill_float_values() {
    assert_eq!(run_prints(r#"<?php echo implode(',', array_fill(0, 3, 0.0)); "#), vec!["0,0,0"]);
}

// ── String to array and back ──────────────────────────────────

#[test] fn implode_with_empty_sep() {
    assert_eq!(run_prints(r#"<?php echo implode('', ['a','b','c','d']); "#), vec!["abcd"]);
}
#[test] fn join_alias() {
    assert_eq!(run_prints(r#"<?php echo join('-', ['x','y','z']); "#), vec!["x-y-z"]);
}
#[test] fn explode_with_limit() {
    assert_eq!(run_prints(r#"<?php echo implode('|', explode(',', 'a,b,c,d', 3)); "#), vec!["a|b|c,d"]);
}
#[test] fn explode_negative_limit() {
    assert_eq!(run_prints(r#"<?php echo count(explode(',', 'a,b,c,d', -1)); "#), vec!["3"]);
}

// ── Misc utility functions ────────────────────────────────────

#[test] fn array_key_first_last() {
    assert_eq!(run_prints(r#"<?php $a = ['x'=>1,'y'=>2,'z'=>3]; echo array_key_first($a) . ',' . array_key_last($a); "#), vec!["x,z"]);
}
#[test] fn compact_and_count() {
    assert_eq!(run_prints(r#"<?php $a=1;$b=2;$c=3; echo count(compact('a','b','c')); "#), vec!["3"]);
}
#[test] fn in_array_with_objects() {
    assert_eq!(run_prints(r#"<?php
class Tag {}
$t = new Tag;
$arr = [$t, new Tag];
echo in_array($t, $arr, true) ? 'yes' : 'no';
"#), vec!["yes"]);
}
#[test] fn array_reverse_basic() {
    assert_eq!(run_prints(r#"<?php echo implode(',', array_reverse([1,2,3,4,5])); "#), vec!["5,4,3,2,1"]);
}
#[test] fn array_reverse_preserve_keys() {
    assert_eq!(run_prints(r#"<?php $r = array_reverse(['a'=>1,'b'=>2,'c'=>3], true); echo implode(',', array_keys($r)); "#), vec!["c,b,a"]);
}
