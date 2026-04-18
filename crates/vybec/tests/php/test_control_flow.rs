use super::helpers;
use helpers::compile_ok;

// If / elseif / else
#[test] fn if_simple() { compile_ok("<?php if ($x > 0) { echo 'yes'; }"); }
#[test] fn if_else() { compile_ok("<?php if ($x > 0) { echo 'pos'; } else { echo 'neg'; }"); }
#[test] fn if_elseif_else() { compile_ok("<?php if ($x > 0) { echo 'pos'; } elseif ($x < 0) { echo 'neg'; } else { echo 'zero'; }"); }
#[test] fn nested_if() { compile_ok("<?php if ($a) { if ($b) { echo 'both'; } }"); }

// While
#[test] fn while_loop() { compile_ok("<?php $i = 0; while ($i < 10) { $i++; }"); }
#[test] fn do_while() { compile_ok("<?php $i = 0; do { $i++; } while ($i < 10);"); }

// For
#[test] fn for_loop() { compile_ok("<?php for ($i = 0; $i < 10; $i++) { echo $i; }"); }
#[test] fn for_no_init() { compile_ok("<?php $i = 0; for (; $i < 10; $i++) {}"); }
#[test] fn for_infinite() { compile_ok("<?php for (;;) { break; }"); }

// Foreach
#[test] fn foreach_value() { compile_ok("<?php foreach ([1,2,3] as $v) { echo $v; }"); }
#[test] fn foreach_key_value() { compile_ok("<?php foreach (['a'=>1] as $k => $v) { echo $k . $v; }"); }
#[test] fn foreach_nested() { compile_ok("<?php foreach ([[1,2],[3,4]] as $row) { foreach ($row as $v) { echo $v; } }"); }

// Switch
#[test] fn switch_basic() { compile_ok("<?php switch ($x) { case 1: echo 'one'; break; case 2: echo 'two'; break; default: echo 'other'; }"); }
#[test] fn switch_fallthrough() { compile_ok("<?php switch ($x) { case 1: case 2: echo 'one or two'; break; }"); }

// Break / Continue
#[test] fn break_in_loop() { compile_ok("<?php for ($i=0;$i<10;$i++) { if ($i==5) break; }"); }
#[test] fn continue_in_loop() { compile_ok("<?php for ($i=0;$i<10;$i++) { if ($i==3) continue; echo $i; }"); }

// Match (PHP 8)
#[test] fn match_expr() { compile_ok("<?php $x = match($v) { 1 => 'one', 2 => 'two', default => 'other' };"); }
