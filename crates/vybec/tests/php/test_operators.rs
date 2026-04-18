use super::helpers;
use helpers::compile_ok;

// Arithmetic
#[test] fn add_sub_mul_div() { compile_ok("<?php $x = 1 + 2 * 3 - 4 / 2;"); }
#[test] fn modulo() { compile_ok("<?php $x = 10 % 3;"); }
#[test] fn power() { compile_ok("<?php $x = 2 ** 10;"); }
#[test] fn unary_neg() { compile_ok("<?php $x = -$a;"); }
#[test] fn unary_not() { compile_ok("<?php $x = !$a;"); }
#[test] fn unary_bitnot() { compile_ok("<?php $x = ~$a;"); }

// String concat
#[test] fn concat_dot() { compile_ok("<?php $x = 'hello' . ' ' . 'world';"); }
#[test] fn concat_assign() { compile_ok("<?php $x = 'a'; $x .= 'b';"); }

// Comparison
#[test] fn loose_eq() { compile_ok("<?php $x = $a == $b;"); }
#[test] fn loose_ne() { compile_ok("<?php $x = $a != $b;"); }
#[test] fn strict_eq() { compile_ok("<?php $x = $a === $b;"); }
#[test] fn strict_ne() { compile_ok("<?php $x = $a !== $b;"); }
#[test] fn lt_gt_le_ge() { compile_ok("<?php $x = $a < $b; $y = $a > $b; $z = $a <= $b; $w = $a >= $b;"); }
#[test] fn spaceship() { compile_ok("<?php $x = 1 <=> 2;"); }

// Logical
#[test] fn and_or() { compile_ok("<?php $x = $a && $b || $c;"); }
#[test] fn short_circuit_and() { compile_ok("<?php $x = false && expensive();"); }
#[test] fn short_circuit_or() { compile_ok("<?php $x = true || expensive();"); }

// Bitwise
#[test] fn bitwise_ops() { compile_ok("<?php $x = $a & $b | $c ^ $d; $y = $a << 2; $z = $b >> 1;"); }

// Ternary / null coalesce
#[test] fn ternary() { compile_ok("<?php $x = $a ? 'yes' : 'no';"); }
#[test] fn short_ternary() { compile_ok("<?php $x = $a ?: 'default';"); }
#[test] fn null_coalesce() { compile_ok("<?php $x = $a ?? 'default';"); }

// Increment / Decrement
#[test] fn pre_inc() { compile_ok("<?php $x = 0; ++$x;"); }
#[test] fn post_inc() { compile_ok("<?php $x = 0; $x++;"); }
#[test] fn pre_dec() { compile_ok("<?php $x = 0; --$x;"); }
#[test] fn post_dec() { compile_ok("<?php $x = 0; $x--;"); }

// Assignment
#[test] fn assign() { compile_ok("<?php $x = 5;"); }
#[test] fn add_assign() { compile_ok("<?php $x = 0; $x += 5;"); }
#[test] fn sub_assign() { compile_ok("<?php $x = 10; $x -= 3;"); }
#[test] fn mul_assign() { compile_ok("<?php $x = 2; $x *= 4;"); }
#[test] fn div_assign() { compile_ok("<?php $x = 10; $x /= 2;"); }
#[test] fn mod_assign() { compile_ok("<?php $x = 10; $x %= 3;"); }
// **= not yet supported by lexer (no StarStarEq token)
// #[test] fn pow_assign() { compile_ok("<?php $x = 2; $x **= 8;"); }
#[test] fn array_access_assign() { compile_ok("<?php $a = [1,2]; $a[0] = 99;"); }
#[test] fn assoc_access_assign() { compile_ok("<?php $a = []; $a['key'] = 'value';"); }
#[test] fn property_assign() { compile_ok("<?php $obj->name = 'test';"); }
