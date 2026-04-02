mod helpers;
use helpers::compile_ok;

#[test] fn function_decl() { compile_ok("<?php function greet($name) { echo 'Hello ' . $name; } greet('World');"); }
#[test] fn default_params() { compile_ok("<?php function add($a, $b = 10) { return $a + $b; } echo add(5);"); }
#[test] fn return_value() { compile_ok("<?php function square($n) { return $n * $n; } $x = square(5);"); }
#[test] fn return_void() { compile_ok("<?php function noop() { return; } noop();"); }
#[test] fn recursive() { compile_ok("<?php function fib($n) { if ($n <= 1) return $n; return fib($n-1) + fib($n-2); } echo fib(10);"); }
#[test] fn multiple_params() { compile_ok("<?php function sum($a, $b, $c) { return $a + $b + $c; } echo sum(1,2,3);"); }
#[test] fn closure_basic() { compile_ok("<?php $fn = function($x) { return $x * 2; }; echo $fn(5);"); }
#[test] fn arrow_fn() { compile_ok("<?php $fn = fn($x) => $x * 2; echo $fn(5);"); }
#[test] fn closure_as_arg() { compile_ok("<?php function apply($fn, $val) { return $fn($val); } echo apply(fn($x) => $x + 1, 41);"); }
#[test] fn global_stmt() { compile_ok("<?php $g = 10; function foo() { global $g; echo $g; } foo();"); }
#[test] fn nested_functions() { compile_ok("<?php function outer() { function inner() { return 42; } return inner(); } echo outer();"); }
