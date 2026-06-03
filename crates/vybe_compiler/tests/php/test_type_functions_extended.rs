use super::helpers::compile_ok;

// ── In-place type mutation ───────────────────────────────────────
#[test]
fn settype_change_in_place() {
    compile_ok(
        r#"<?php
$v = '42';
settype($v, 'integer');
echo $v;
"#,
    );
}

// ── PHP 8 precise type name ──────────────────────────────────────
#[test]
fn get_debug_type_builtin() {
    compile_ok(
        r#"<?php
echo get_debug_type(42);
echo get_debug_type(3.14);
echo get_debug_type('hello');
echo get_debug_type(null);
"#,
    );
}

// ── intval with explicit base ────────────────────────────────────
#[test]
fn intval_with_hex_base() {
    compile_ok(
        r#"<?php
$n = intval('1F', 16);
echo $n;
"#,
    );
}

#[test]
fn intval_with_octal_base() {
    compile_ok(
        r#"<?php
$n = intval('777', 8);
echo $n;
"#,
    );
}

#[test]
fn intval_with_binary_base() {
    compile_ok(
        r#"<?php
$n = intval('1010', 2);
echo $n;
"#,
    );
}

// ── Float predicate functions ────────────────────────────────────
#[test]
fn is_finite_regular_float() {
    compile_ok(
        r#"<?php
echo is_finite(3.14) ? 'yes' : 'no';
echo is_finite(INF) ? 'yes' : 'no';
"#,
    );
}

#[test]
fn is_infinite_check() {
    compile_ok(
        r#"<?php
echo is_infinite(INF) ? 'yes' : 'no';
echo is_infinite(3.14) ? 'yes' : 'no';
"#,
    );
}

#[test]
fn is_nan_check() {
    compile_ok(
        r#"<?php
echo is_nan(NAN) ? 'yes' : 'no';
echo is_nan(0.0) ? 'yes' : 'no';
"#,
    );
}

// ── ctype functions ──────────────────────────────────────────────
#[test]
fn ctype_alpha_all_letters() {
    compile_ok(
        r#"<?php
echo ctype_alpha('Hello') ? 'yes' : 'no';
echo ctype_alpha('Hello1') ? 'yes' : 'no';
"#,
    );
}

#[test]
fn ctype_digit_all_digits() {
    compile_ok(
        r#"<?php
echo ctype_digit('12345') ? 'yes' : 'no';
echo ctype_digit('123a5') ? 'yes' : 'no';
"#,
    );
}

#[test]
fn ctype_alnum_alphanumeric() {
    compile_ok(
        r#"<?php
echo ctype_alnum('abc123') ? 'yes' : 'no';
echo ctype_alnum('abc!23') ? 'yes' : 'no';
"#,
    );
}

#[test]
fn ctype_space_whitespace() {
    compile_ok(
        r#"<?php
echo ctype_space("  \t\n") ? 'yes' : 'no';
echo ctype_space('  x  ') ? 'yes' : 'no';
"#,
    );
}

#[test]
fn ctype_upper_uppercase() {
    compile_ok(
        r#"<?php
echo ctype_upper('HELLO') ? 'yes' : 'no';
echo ctype_upper('Hello') ? 'yes' : 'no';
"#,
    );
}

#[test]
fn ctype_lower_lowercase() {
    compile_ok(
        r#"<?php
echo ctype_lower('hello') ? 'yes' : 'no';
echo ctype_lower('Hello') ? 'yes' : 'no';
"#,
    );
}

#[test]
fn ctype_punct_punctuation() {
    compile_ok(
        r#"<?php
echo ctype_punct('!@#') ? 'yes' : 'no';
echo ctype_punct('!a#') ? 'yes' : 'no';
"#,
    );
}

// ── Object / callable predicates ────────────────────────────────
#[test]
fn is_object_check() {
    compile_ok(
        r#"<?php
class Point { public int $x; public int $y; }
$p = new Point();
echo is_object($p) ? 'yes' : 'no';
echo is_object([]) ? 'yes' : 'no';
"#,
    );
}

#[test]
fn is_callable_check() {
    compile_ok(
        r#"<?php
$fn = function($x) { return $x * 2; };
echo is_callable($fn) ? 'yes' : 'no';
echo is_callable(42) ? 'yes' : 'no';
"#,
    );
}

#[test]
fn is_float_alias() {
    compile_ok(
        r#"<?php
echo is_float(3.14) ? 'yes' : 'no';
echo is_float(42) ? 'yes' : 'no';
"#,
    );
}

// ── PHP numeric constants ────────────────────────────────────────
#[test]
fn php_int_max_arithmetic() {
    compile_ok(
        r#"<?php
$big = PHP_INT_MAX;
$overflow = $big + 1;
echo is_float($overflow) ? 'float' : 'int';
"#,
    );
}

#[test]
fn php_float_epsilon_comparison() {
    compile_ok(
        r#"<?php
$a = 0.1 + 0.2;
$b = 0.3;
$close = abs($a - $b) < PHP_FLOAT_EPSILON;
echo $close ? 'close' : 'far';
"#,
    );
}
