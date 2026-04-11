use super::helpers::compile_ok;

// ── Empty / minimal programs ────────────────────────────────
#[test] fn empty_program() { compile_ok("<?php"); }
#[test] fn just_semicolons() { compile_ok("<?php ;;;"); }
#[test] fn only_echo() { compile_ok("<?php echo 'hi';"); }

// ── Nested control flow ─────────────────────────────────────
#[test] fn deeply_nested_if() { compile_ok("<?php if (true) { if (true) { if (true) { echo 'deep'; } } }"); }
#[test] fn nested_loops() { compile_ok("<?php for ($i=0;$i<3;$i++) { for ($j=0;$j<3;$j++) { for ($k=0;$k<3;$k++) { echo $i+$j+$k; } } }"); }
#[test] fn break_nested() { compile_ok("<?php for ($i=0;$i<10;$i++) { for ($j=0;$j<10;$j++) { if ($j==5) break; } }"); }
#[test] fn continue_nested() { compile_ok("<?php for ($i=0;$i<5;$i++) { if ($i==2) continue; for ($j=0;$j<5;$j++) { if ($j==3) continue; } }"); }

// ── Edge case expressions ───────────────────────────────────
#[test] fn chained_ternary() { compile_ok("<?php $x = $a ? 'a' : ($b ? 'b' : 'c');"); }
#[test] fn chained_coalesce() { compile_ok("<?php $x = $a ?? $b ?? $c ?? 'default';"); }
#[test] fn negative_numbers() { compile_ok("<?php $x = -42; $y = -3.14; $z = -$x;"); }
#[test] fn large_numbers() { compile_ok("<?php $x = 999999999; $y = 0.000001;"); }
#[test] fn empty_string() { compile_ok("<?php $x = ''; $y = \"\"; echo strlen($x);"); }
#[test] fn multiline_string() { compile_ok("<?php $x = 'line1\nline2\nline3'; echo $x;"); }
#[test] fn special_chars_in_string() { compile_ok(r#"<?php $x = "tab\there\nnewline\r\n";"#); }

// ── Complex assignments ─────────────────────────────────────
#[test] fn chain_assign() { compile_ok("<?php $a = $b = $c = 0;"); }
#[test] fn nested_array_assign() { compile_ok("<?php $a = []; $a['x'] = []; $a['x']['y'] = 42;"); }
#[test] fn assign_in_condition() { compile_ok("<?php if ($x = 5) { echo $x; }"); }
#[test] fn assign_in_while() { compile_ok("<?php $i = 0; while (($i = $i + 1) < 10) { echo $i; }"); }

// ── Mixed expressions ───────────────────────────────────────
#[test] fn expr_as_statement() { compile_ok("<?php 1 + 2; 'hello'; true;"); }
#[test] fn complex_math() { compile_ok("<?php $x = (1 + 2) * (3 - 4) / 5 % 6 + pow(2, 10);"); }
#[test] fn string_in_condition() { compile_ok("<?php if ('hello') { echo 'truthy'; } if ('') { echo 'falsy'; }"); }
#[test] fn null_checks() { compile_ok("<?php $x = null; if ($x === null) {} if (is_null($x)) {} if (!isset($x)) {}"); }

// ── Function edge cases ────────────────────────────────────
#[test] fn variadic() { compile_ok("<?php function sum(...$nums) { return array_sum($nums); } echo sum(1,2,3,4,5);"); }
#[test] fn return_early() { compile_ok("<?php function check($x) { if ($x < 0) return 'negative'; if ($x == 0) return 'zero'; return 'positive'; } echo check(-1);"); }
#[test] fn recursive_mutual() { compile_ok("<?php function isEven($n) { if ($n == 0) return true; return isOdd($n - 1); } function isOdd($n) { if ($n == 0) return false; return isEven($n - 1); } echo isEven(4);"); }
#[test] fn default_null() { compile_ok("<?php function foo($x = null) { return $x ?? 'default'; } echo foo(); echo foo('val');"); }

// ── Class edge cases ────────────────────────────────────────
#[test] fn empty_class() { compile_ok("<?php class Empty {} $e = new Empty();"); }
#[test] fn self_referencing() { compile_ok("<?php class Node { public $next; } $a = new Node(); $b = new Node(); $a->next = $b; $b->next = $a;"); }
#[test] fn method_returns_new() { compile_ok("<?php class Factory { public function create() { return new Factory(); } } $f = new Factory(); $f2 = $f->create();"); }
#[test] fn property_method_same_name() { compile_ok("<?php class A { public $name = 'prop'; public function name() { return 'method'; } } $a = new A(); echo $a->name; echo $a->name();"); }
#[test] fn multiple_constructors_via_static() { compile_ok(r#"<?php
class Color {
    public $r; public $g; public $b;
    public function __construct($r, $g, $b) { $this->r = $r; $this->g = $g; $this->b = $b; }
    public static function red() { return new Color(255, 0, 0); }
    public static function fromHex($hex) { return new Color(0, 0, 0); }
}
$red = Color::red();
echo $red->r;
"#); }

// ── Scope edge cases ────────────────────────────────────────
#[test] fn redefine_var_in_loop() { compile_ok("<?php for ($i=0;$i<3;$i++) { $x = $i; } echo $x;"); }
#[test] fn same_name_diff_scope() { compile_ok("<?php $x = 'global'; function foo() { $x = 'local'; return $x; } echo foo(); echo $x;"); }
#[test] fn closure_modifies_local() { compile_ok("<?php $arr = []; $fn = function($v) use ($arr) { array_push($arr, $v); return $arr; }; $fn(1);"); }

// ── PHP-specific quirks ─────────────────────────────────────
#[test] fn loose_comparison_quirks() { compile_ok("<?php echo 0 == '0'; echo 0 == ''; echo '' == null; echo 0 == null;"); }
#[test] fn string_numeric_ops() { compile_ok("<?php $x = '5' + '3'; $y = '10' - '3'; $z = '2' * '4';"); }
#[test] fn array_as_bool() { compile_ok("<?php if ([]) { echo 'truthy'; } else { echo 'falsy'; } if ([1]) { echo 'truthy'; }"); }
#[test] fn null_coalesce_nested() { compile_ok("<?php $config = ['db' => ['host' => 'localhost']]; $host = $config['db']['host'] ?? 'default';"); }
#[test] fn match_no_break_needed() { compile_ok("<?php $x = match(2) { 1 => 'one', 2 => 'two', 3 => 'three', default => '?' };"); }
