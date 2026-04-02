mod helpers;
use helpers::compile_ok;

// ── Basic closures ──────────────────────────────────────────
#[test] fn closure_assign() { compile_ok("<?php $fn = function($x) { return $x * 2; }; echo $fn(5);"); }
#[test] fn closure_no_args() { compile_ok("<?php $fn = function() { return 42; }; echo $fn();"); }
#[test] fn closure_multi_args() { compile_ok("<?php $fn = function($a, $b, $c) { return $a + $b + $c; }; echo $fn(1,2,3);"); }
#[test] fn closure_with_body() { compile_ok("<?php $fn = function($n) { $result = 1; for ($i = 1; $i <= $n; $i++) { $result *= $i; } return $result; }; echo $fn(5);"); }

// ── Closure use captures ────────────────────────────────────
#[test] fn use_single() { compile_ok("<?php $x = 10; $fn = function() use ($x) { return $x; }; echo $fn();"); }
#[test] fn use_multiple() { compile_ok("<?php $a = 1; $b = 2; $fn = function() use ($a, $b) { return $a + $b; }; echo $fn();"); }
#[test] fn use_with_params() { compile_ok("<?php $factor = 3; $fn = function($x) use ($factor) { return $x * $factor; }; echo $fn(5);"); }
#[test] fn use_does_not_mutate_outer() { compile_ok("<?php $x = 'original'; $fn = function() use ($x) { $x = 'modified'; return $x; }; echo $fn(); echo $x;"); }
#[test] fn use_trailing_comma() { compile_ok("<?php $a = 1; $fn = function() use ($a,) { return $a; };"); }

// ── Arrow functions ─────────────────────────────────────────
#[test] fn arrow_basic() { compile_ok("<?php $fn = fn($x) => $x * 2; echo $fn(5);"); }
#[test] fn arrow_no_params() { compile_ok("<?php $fn = fn() => 42; echo $fn();"); }
#[test] fn arrow_multi_params() { compile_ok("<?php $fn = fn($a, $b) => $a + $b; echo $fn(3, 4);"); }
#[test] fn arrow_auto_capture() { compile_ok("<?php $x = 10; $fn = fn($y) => $x + $y; echo $fn(5);"); }
#[test] fn arrow_nested_capture() { compile_ok("<?php $a = 1; $outer = fn($b) => fn($c) => $a + $b + $c;"); }
#[test] fn arrow_as_callback() { compile_ok("<?php $doubled = array_map(fn($n) => $n * 2, [1, 2, 3]);"); }

// ── Closures as callbacks ───────────────────────────────────
#[test] fn closure_in_map() { compile_ok("<?php array_map(function($x) { return $x * $x; }, [1,2,3]);"); }
#[test] fn closure_in_filter() { compile_ok("<?php array_filter([1,2,3,4,5], function($x) { return $x > 2; });"); }
#[test] fn closure_in_reduce() { compile_ok("<?php array_reduce([1,2,3], function($carry, $item) { return $carry + $item; }, 0);"); }
#[test] fn closure_in_usort() { compile_ok("<?php $a = [3,1,2]; usort($a);"); }
#[test] fn closure_passed_to_function() { compile_ok("<?php function apply($fn, $val) { return $fn($val); } echo apply(fn($x) => $x + 1, 41);"); }

// ── IIFE ────────────────────────────────────────────────────
#[test] fn iife() { compile_ok("<?php $result = (function() { return 42; })(); echo $result;"); }
#[test] fn iife_with_args() { compile_ok("<?php $result = (function($a, $b) { return $a + $b; })(3, 4); echo $result;"); }

// ── Closure returning closure ───────────────────────────────
#[test] fn closure_factory() { compile_ok("<?php function multiplier($factor) { return function($x) use ($factor) { return $x * $factor; }; } $double = multiplier(2); echo $double(5);"); }
#[test] fn arrow_factory() { compile_ok("<?php function adder($n) { return fn($x) => $x + $n; } $add5 = adder(5); echo $add5(10);"); }
