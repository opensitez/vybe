use super::helpers::{compile_ok, run_prints};

fn assert_php_output(src: &str, expected: &[&str]) {
    assert_eq!(
        run_prints(src),
        expected
            .iter()
            .map(|value| (*value).to_string())
            .collect::<Vec<_>>()
    );
}

// Arithmetic
#[test]
fn add_sub_mul_div() {
    compile_ok("<?php $x = 1 + 2 * 3 - 4 / 2;");
}
#[test]
fn modulo() {
    compile_ok("<?php $x = 10 % 3;");
}
#[test]
fn power() {
    compile_ok("<?php $x = 2 ** 10;");
}
#[test]
fn unary_neg() {
    compile_ok("<?php $x = -$a;");
}
#[test]
fn unary_not() {
    compile_ok("<?php $x = !$a;");
}
#[test]
fn unary_bitnot() {
    compile_ok("<?php $x = ~$a;");
}

// String concat
#[test]
fn concat_dot() {
    compile_ok("<?php $x = 'hello' . ' ' . 'world';");
}
#[test]
fn concat_assign() {
    compile_ok("<?php $x = 'a'; $x .= 'b';");
}

// Comparison
#[test]
fn loose_eq() {
    compile_ok("<?php $x = $a == $b;");
}
#[test]
fn loose_ne() {
    compile_ok("<?php $x = $a != $b;");
}
#[test]
fn strict_eq() {
    compile_ok("<?php $x = $a === $b;");
}
#[test]
fn strict_ne() {
    compile_ok("<?php $x = $a !== $b;");
}
#[test]
fn lt_gt_le_ge() {
    compile_ok("<?php $x = $a < $b; $y = $a > $b; $z = $a <= $b; $w = $a >= $b;");
}
#[test]
fn spaceship() {
    compile_ok("<?php $x = 1 <=> 2;");
}

// Logical
#[test]
fn and_or() {
    compile_ok("<?php $x = $a && $b || $c;");
}
#[test]
fn short_circuit_and() {
    compile_ok("<?php $x = false && expensive();");
}
#[test]
fn short_circuit_or() {
    compile_ok("<?php $x = true || expensive();");
}

// Bitwise
#[test]
fn bitwise_ops() {
    compile_ok("<?php $x = $a & $b | $c ^ $d; $y = $a << 2; $z = $b >> 1;");
}

// Ternary / null coalesce
#[test]
fn ternary() {
    compile_ok("<?php $x = $a ? 'yes' : 'no';");
}
#[test]
fn short_ternary() {
    compile_ok("<?php $x = $a ?: 'default';");
}
#[test]
fn null_coalesce() {
    compile_ok("<?php $x = $a ?? 'default';");
}

// Increment / Decrement
#[test]
fn pre_inc() {
    compile_ok("<?php $x = 0; ++$x;");
}
#[test]
fn post_inc() {
    compile_ok("<?php $x = 0; $x++;");
}
#[test]
fn pre_dec() {
    compile_ok("<?php $x = 0; --$x;");
}
#[test]
fn post_dec() {
    compile_ok("<?php $x = 0; $x--;");
}

// Assignment
#[test]
fn assign() {
    compile_ok("<?php $x = 5;");
}
#[test]
fn add_assign() {
    compile_ok("<?php $x = 0; $x += 5;");
}
#[test]
fn sub_assign() {
    compile_ok("<?php $x = 10; $x -= 3;");
}
#[test]
fn mul_assign() {
    compile_ok("<?php $x = 2; $x *= 4;");
}
#[test]
fn div_assign() {
    compile_ok("<?php $x = 10; $x /= 2;");
}
#[test]
fn mod_assign() {
    compile_ok("<?php $x = 10; $x %= 3;");
}
// **= not yet supported by lexer (no StarStarEq token)
// #[test] fn pow_assign() { compile_ok("<?php $x = 2; $x **= 8;"); }
#[test]
fn array_access_assign() {
    compile_ok("<?php $a = [1,2]; $a[0] = 99;");
}
#[test]
fn assoc_access_assign() {
    compile_ok("<?php $a = []; $a['key'] = 'value';");
}
#[test]
fn property_assign() {
    compile_ok("<?php $obj->name = 'test';");
}

#[test]
fn logical_operators_runtime_results() {
    assert_php_output(
        r#"<?php
$_SERVER = [];
if (!isset($defaultLang) && !empty($_SERVER['HTTP_ACCEPT_LANGUAGE'])) {
	echo 'and-bad';
} else {
	echo 'and-ok';
}

if (!empty($_SERVER['HTTP_ACCEPT_LANGUAGE']) || !isset($defaultLang)) {
	echo 'or-ok';
} else {
	echo 'or-bad';
}

if (true and false) {
	echo 'word-and-bad';
} else {
	echo 'word-and-ok';
}

if (false or true) {
	echo 'word-or-ok';
} else {
	echo 'word-or-bad';
}

if (true xor true) {
	echo 'word-xor-bad';
} else {
	echo 'word-xor-ok';
}
"#,
        &["and-okor-okword-and-okword-or-okword-xor-ok"],
    );
}

#[test]
fn arithmetic_comparison_and_control_operator_runtime_results() {
    assert_php_output(
        r#"<?php
echo 1 + 2;
echo 7 - 4;
echo 6 * 7;
echo 7 / 2;
echo 7 % 3;
echo 2 ** 3;
echo 'a' . 'b';
echo (-5) + 8;
echo (+5);
echo (!false) ? 't' : 'f';
echo (2 < 3) ? 't' : 'f';
echo (3 > 2) ? 't' : 'f';
echo (3 <= 3) ? 't' : 'f';
echo (4 >= 5) ? 't' : 'f';
echo (2 == '2') ? 't' : 'f';
echo (2 === '2') ? 't' : 'f';
echo (2 != 3) ? 't' : 'f';
echo (2 !== '2') ? 't' : 'f';
echo 1 <=> 2;
echo 2 <=> 2;
echo 3 <=> 2;
echo null ?? 'fallback';
echo 'value' ?? 'fallback';
echo false ? 'then' : 'else';
echo 0 ?: 'fallback';
"#,
        &["33423.518ab35ttttftftt-101fallbackvalueelsefallback"],
    );
}

#[test]
fn bitwise_and_shift_operator_runtime_results() {
    assert_php_output(
        r#"<?php
echo 6 & 3;
echo 6 | 3;
echo 6 ^ 3;
echo 1 << 3;
echo 8 >> 2;
echo ~1;
"#,
        &["27582-2"],
    );
}

#[test]
fn compound_assignment_operator_runtime_results() {
    assert_php_output(
        r#"<?php
$x = 5;
$x += 2;
echo $x;
$x -= 4;
echo $x;
$x *= 3;
echo $x;
$x /= 9;
echo $x;
$x %= 2;
echo $x;

$text = 'a';
$text .= 'b';
echo $text;

$bits = 6;
$bits &= 3;
echo $bits;
$bits |= 4;
echo $bits;
$bits ^= 1;
echo $bits;

$shift = 1;
$shift <<= 3;
echo $shift;
$shift >>= 2;
echo $shift;

$fallback = null;
$fallback ??= 'set';
echo $fallback;
$fallback ??= 'again';
echo $fallback;
"#,
        &["73911ab26782setset"],
    );
}
