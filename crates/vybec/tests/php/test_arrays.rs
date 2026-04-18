use super::helpers;
use helpers::compile_ok;

// ── Array creation ──────────────────────────────────────────
#[test] fn array_indexed() { compile_ok("<?php $a = [1, 2, 3]; echo $a[0];"); }
#[test] fn array_assoc() { compile_ok("<?php $a = ['name' => 'John', 'age' => 30]; echo $a['name'];"); }
#[test] fn array_empty() { compile_ok("<?php $a = []; echo count($a);"); }
#[test] fn array_nested() { compile_ok("<?php $a = [[1,2],[3,4]]; echo $a[0][1];"); }
#[test] fn array_mixed_keys() { compile_ok("<?php $a = [0 => 'a', 'key' => 'b']; echo $a[0]; echo $a['key'];"); }
#[test] fn array_function() { compile_ok("<?php $a = array(1, 2, 3); echo count($a);"); }
#[test] fn array_trailing_comma() { compile_ok("<?php $a = [1, 2, 3,]; echo count($a);"); }

// ── Array modification ──────────────────────────────────────
#[test] fn array_assign_index() { compile_ok("<?php $a = [1, 2]; $a[0] = 99; echo $a[0];"); }
#[test] fn array_assign_key() { compile_ok("<?php $a = []; $a['x'] = 'hello'; echo $a['x'];"); }
#[test] fn array_push() { compile_ok("<?php $a = [1]; array_push($a, 2); echo count($a);"); }
#[test] fn array_pop() { compile_ok("<?php $a = [1, 2, 3]; $v = array_pop($a); echo $v;"); }
#[test] fn array_shift() { compile_ok("<?php $a = [1, 2, 3]; $v = array_shift($a); echo $v;"); }

// ── Array query ─────────────────────────────────────────────
#[test] fn in_array() { compile_ok("<?php echo in_array(2, [1, 2, 3]);"); }
#[test] fn array_search() { compile_ok("<?php echo array_search('b', ['a', 'b', 'c']);"); }
#[test] fn array_key_exists() { compile_ok("<?php echo array_key_exists('name', ['name' => 'John']);"); }
#[test] fn count_arr() { compile_ok("<?php echo count([1, 2, 3]);"); }

// ── Array transform ─────────────────────────────────────────
#[test] fn array_merge() { compile_ok("<?php $x = array_merge([1, 2], [3, 4]); echo count($x);"); }
#[test] fn array_slice() { compile_ok("<?php $x = array_slice([1, 2, 3, 4, 5], 1, 3); echo count($x);"); }
#[test] fn array_reverse() { compile_ok("<?php $x = array_reverse([1, 2, 3]);"); }
#[test] fn array_keys() { compile_ok("<?php $x = array_keys(['a' => 1, 'b' => 2]);"); }
#[test] fn array_values() { compile_ok("<?php $x = array_values(['a' => 1, 'b' => 2]);"); }
#[test] fn sort_array() { compile_ok("<?php $a = [3, 1, 2]; sort($a);"); }
#[test] fn range_func() { compile_ok("<?php $x = range(1, 10);"); }
#[test] fn array_sum() { compile_ok("<?php echo array_sum([1, 2, 3, 4]);"); }

// ── Callback array ops ──────────────────────────────────────
#[test] fn array_map() { compile_ok("<?php $x = array_map(fn($n) => $n * 2, [1, 2, 3]);"); }
#[test] fn array_filter() { compile_ok("<?php $x = array_filter([0, 1, '', 'a', null], fn($v) => $v);"); }
#[test] fn array_filter_no_cb() { compile_ok("<?php $x = array_filter([0, 1, 2, null, 3]);"); }
#[test] fn array_reduce() { compile_ok("<?php $sum = array_reduce([1,2,3], fn($carry, $item) => $carry + $item, 0);"); }
#[test] fn array_walk() { compile_ok("<?php $a = [1,2,3]; array_walk($a, fn($v, $k) => $v);"); }
#[test] fn usort_array() { compile_ok("<?php $a = [3,1,2]; usort($a);"); }

// ── Array destructuring ─────────────────────────────────────
#[test] fn list_basic() { compile_ok("<?php list($a, $b) = [1, 2]; echo $a;"); }
#[test] fn short_list() { compile_ok("<?php [$x, $y, $z] = [10, 20, 30]; echo $y;"); }
#[test] fn compact_extract() { compile_ok("<?php $a = 1; $b = 2; $arr = compact('a', 'b');"); }

// ── Foreach iteration ───────────────────────────────────────
#[test] fn foreach_indexed() { compile_ok("<?php foreach ([1,2,3] as $v) { echo $v; }"); }
#[test] fn foreach_assoc() { compile_ok("<?php foreach (['a'=>1, 'b'=>2] as $k => $v) { echo $k . ': ' . $v; }"); }
#[test] fn foreach_nested() { compile_ok("<?php foreach ([[1,2],[3,4]] as $row) { foreach ($row as $v) { echo $v; } }"); }

// ── Spread operator ─────────────────────────────────────────
#[test] fn spread_in_array() { compile_ok("<?php $a = [1, 2]; $b = [...$a, 3, 4];"); }
#[test] fn spread_in_call() { compile_ok("<?php function sum(...$nums) { return 0; } sum(...[1,2,3]);"); }
