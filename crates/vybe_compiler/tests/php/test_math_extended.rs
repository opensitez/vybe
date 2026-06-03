use super::helpers::compile_ok;

// ── Practical geometry ───────────────────────────────────────────
#[test]
fn hypot_right_triangle() {
    compile_ok(
        r#"<?php
$c = hypot(3, 4);
echo $c;
"#,
    );
}

// ── Angle conversion ─────────────────────────────────────────────
#[test]
fn deg2rad_conversion() {
    compile_ok(
        r#"<?php
$r = deg2rad(90);
echo $r;
"#,
    );
}

#[test]
fn rad2deg_conversion() {
    compile_ok(
        r#"<?php
$d = rad2deg(M_PI / 2);
echo $d;
"#,
    );
}

// ── Inverse trigonometry ─────────────────────────────────────────
#[test]
fn atan2_two_args() {
    compile_ok(
        r#"<?php
$a = atan2(1, 1);
echo $a;
"#,
    );
}

#[test]
fn asin_arc_sine() {
    compile_ok(
        r#"<?php
$a = asin(0.5);
echo $a;
"#,
    );
}

#[test]
fn acos_arc_cosine() {
    compile_ok(
        r#"<?php
$a = acos(0.5);
echo $a;
"#,
    );
}

// ── Hyperbolic functions ─────────────────────────────────────────
#[test]
fn sinh_hyperbolic_sine() {
    compile_ok(
        r#"<?php
$v = sinh(1);
echo $v;
"#,
    );
}

#[test]
fn cosh_hyperbolic_cosine() {
    compile_ok(
        r#"<?php
$v = cosh(1);
echo $v;
"#,
    );
}

#[test]
fn tanh_hyperbolic_tangent() {
    compile_ok(
        r#"<?php
$v = tanh(0.5);
echo $v;
"#,
    );
}

// ── Logarithms ───────────────────────────────────────────────────
#[test]
fn log2_base2() {
    compile_ok(
        r#"<?php
$v = log2(1024);
echo $v;
"#,
    );
}

#[test]
fn log10_base10() {
    compile_ok(
        r#"<?php
$v = log10(1000);
echo $v;
"#,
    );
}

// ── Base conversion ──────────────────────────────────────────────
#[test]
fn base_convert_arbitrary_bases() {
    compile_ok(
        r#"<?php
$dec = base_convert('ff', 16, 10);
$hex = base_convert('255', 10, 16);
echo $dec;
echo $hex;
"#,
    );
}

#[test]
fn bindec_binary_to_decimal() {
    compile_ok(
        r#"<?php
$n = bindec('11001');
echo $n;
"#,
    );
}

#[test]
fn octdec_octal_to_decimal() {
    compile_ok(
        r#"<?php
$n = octdec('77');
echo $n;
"#,
    );
}

#[test]
fn hexdec_hex_to_decimal() {
    compile_ok(
        r#"<?php
$n = hexdec('1f');
echo $n;
"#,
    );
}

#[test]
fn decbin_decimal_to_binary() {
    compile_ok(
        r#"<?php
$s = decbin(42);
echo $s;
"#,
    );
}

#[test]
fn decoct_decimal_to_octal() {
    compile_ok(
        r#"<?php
$s = decoct(255);
echo $s;
"#,
    );
}

#[test]
fn dechex_decimal_to_hex() {
    compile_ok(
        r#"<?php
$s = dechex(255);
echo $s;
"#,
    );
}

// ── Integer / float division ─────────────────────────────────────
#[test]
fn intdiv_integer_division() {
    compile_ok(
        r#"<?php
$q = intdiv(17, 5);
echo $q;
"#,
    );
}

#[test]
fn fmod_float_modulo() {
    compile_ok(
        r#"<?php
$r = fmod(10.5, 3.0);
echo $r;
"#,
    );
}
