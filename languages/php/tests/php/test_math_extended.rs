use super::helpers::{compile_ok, run_prints};

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

#[test]
fn math_abs_and_sign() {
    assert_eq!(
        run_prints(r#"<?php echo abs(-15); echo '|'; echo abs(15); "#),
        vec!["15|15"]
    );
}

#[test]
fn math_round_and_floor_ceil() {
    assert_eq!(
        run_prints(
            r#"<?php echo round(3.6); echo '|'; echo floor(3.6); echo '|'; echo ceil(3.1); "#
        ),
        vec!["4|3|4"]
    );
}

#[test]
fn math_pow_and_sqrt() {
    assert_eq!(
        run_prints(r#"<?php echo pow(3, 4); echo '|'; echo sqrt(81); "#),
        vec!["81|9"]
    );
}

#[test]
fn math_min_max() {
    assert_eq!(
        run_prints(r#"<?php echo min(9, 3, 7); echo '|'; echo max([4, 12, 6]); "#),
        vec!["3|12"]
    );
}

#[test]
fn math_intdiv_rounding_toward_zero() {
    assert_eq!(
        run_prints(r#"<?php echo intdiv(-7, 2); echo '|'; echo intdiv(7, 2); "#),
        vec!["-3|3"]
    );
}

#[test]
fn math_trig_hypotenuse() {
    assert_eq!(run_prints(r#"<?php echo hypot(5, 12); "#), vec!["13"]);
}

#[test]
fn math_logarithms_consistent() {
    assert_eq!(
        run_prints(r#"<?php echo log10(1000); echo '|'; echo log(1); "#),
        vec!["3|0"]
    );
}

#[test]
fn math_base_conversion_roundtrip() {
    assert_eq!(
        run_prints(
            r#"<?php echo base_convert('1010', 2, 10); echo '|'; echo base_convert('1e', 36, 10); "#
        ),
        vec!["10|50"]
    );
}

#[test]
fn math_bitwise_not_for_int() {
    assert_eq!(run_prints(r#"<?php $x = 0b1111; echo ~$x; "#), vec!["-16"]);
}
