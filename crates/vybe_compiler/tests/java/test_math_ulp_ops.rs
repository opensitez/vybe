use crate::helpers::run_main;

#[test]
fn strict_math_scalb_doubles_value() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.scalb(3.0, 2));"#);
    assert_eq!(out, vec!["12"]);
}

#[test]
fn strict_math_scalb_halves() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.scalb(8.0, -1));"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn strict_math_scalb_zero_exp() {
    let out = run_main(r#"System.out.println(java.lang.StrictMath.scalb(5.5, 0));"#);
    assert_eq!(out, vec!["5.5"]);
}

#[test]
fn strict_math_next_up_from_zero() {
    let out = run_main(r#"System.out.println(java.lang.StrictMath.nextUp(0.0) > 0.0);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn strict_math_next_up_from_negative() {
    let out = run_main(r#"System.out.println(java.lang.StrictMath.nextUp(-1.0) > -1.0);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn strict_math_next_up_from_positive() {
    let out = run_main(r#"System.out.println(java.lang.StrictMath.nextUp(1.0) > 1.0);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn strict_math_next_down_from_zero() {
    let out = run_main(r#"System.out.println(java.lang.StrictMath.nextDown(0.0) < 0.0);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn strict_math_next_down_from_positive() {
    let out = run_main(r#"System.out.println(java.lang.StrictMath.nextDown(1.0) < 1.0);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn strict_math_next_down_from_negative() {
    let out = run_main(r#"System.out.println(java.lang.StrictMath.nextDown(-1.0) < -1.0);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn strict_math_next_after_toward_positive() {
    let out = run_main(r#"System.out.println(java.lang.StrictMath.nextAfter(1.0, 2.0) > 1.0);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn strict_math_next_after_toward_negative() {
    let out = run_main(r#"System.out.println(java.lang.StrictMath.nextAfter(1.0, 0.0) < 1.0);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn strict_math_ulp_of_one() {
    let out = run_main(r#"System.out.println(java.lang.StrictMath.ulp(1.0) > 0.0);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn strict_math_ulp_of_zero() {
    let out = run_main(r#"System.out.println(java.lang.StrictMath.ulp(0.0) > 0.0);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn strict_math_get_exponent_of_eight() {
    let out = run_main(r#"System.out.println(java.lang.StrictMath.getExponent(8.0));"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn strict_math_get_exponent_of_one() {
    let out = run_main(r#"System.out.println(java.lang.StrictMath.getExponent(1.0));"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn strict_math_get_exponent_of_half() {
    let out = run_main(r#"System.out.println(java.lang.StrictMath.getExponent(0.5));"#);
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn strict_math_copy_sign_positive() {
    let out = run_main(r#"System.out.println(java.lang.StrictMath.copySign(5.0, 1.0));"#);
    assert_eq!(out, vec!["5.0"]);
}

#[test]
fn strict_math_copy_sign_negative() {
    let out = run_main(r#"System.out.println(java.lang.StrictMath.copySign(5.0, -1.0));"#);
    assert_eq!(out, vec!["-5.0"]);
}

#[test]
fn strict_math_copy_sign_negative_mag() {
    let out = run_main(r#"System.out.println(java.lang.StrictMath.copySign(-5.0, 1.0));"#);
    assert_eq!(out, vec!["5.0"]);
}

#[test]
fn strict_math_ieee_remainder_seven_four() {
    let out =
        run_main(r#"System.out.println((int) java.lang.StrictMath.IEEEremainder(7.0, 4.0));"#);
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn strict_math_ieee_remainder_eight_three() {
    let out =
        run_main(r#"System.out.println((int) java.lang.StrictMath.IEEEremainder(8.0, 3.0));"#);
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn strict_math_fma_basic() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.fma(2.0, 3.0, 4.0));"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn strict_math_fma_negative() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.fma(2.0, 3.0, -1.0));"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn strict_math_abs_negative() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.abs(-9));"#);
    assert_eq!(out, vec!["9"]);
}

#[test]
fn strict_math_abs_positive() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.abs(9));"#);
    assert_eq!(out, vec!["9"]);
}

#[test]
fn strict_math_max_int() {
    let out = run_main(r#"System.out.println(java.lang.StrictMath.max(3, 9));"#);
    assert_eq!(out, vec!["9"]);
}

#[test]
fn strict_math_min_int() {
    let out = run_main(r#"System.out.println(java.lang.StrictMath.min(3, 9));"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn strict_math_sqrt_four() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.sqrt(16.0));"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn strict_math_cbrt_eight() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.cbrt(27.0));"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn strict_math_hypot_three_four() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.hypot(3.0, 4.0));"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn strict_math_pow_two_cubed() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.pow(2.0, 3.0));"#);
    assert_eq!(out, vec!["8"]);
}

#[test]
fn strict_math_exp_zero() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.exp(0.0));"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn strict_math_log_e() {
    let out =
        run_main(r#"System.out.println((int) java.lang.StrictMath.log(java.lang.StrictMath.E));"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn strict_math_log10_thousand() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.log10(1000.0));"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn strict_math_sin_zero() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.sin(0.0));"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn strict_math_cos_zero() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.cos(0.0));"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn strict_math_tan_zero() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.tan(0.0));"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn strict_math_asin_zero() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.asin(0.0));"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn strict_math_acos_one() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.acos(1.0));"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn strict_math_atan_zero() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.atan(0.0));"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn strict_math_atan2_origin() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.atan2(0.0, 1.0));"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn strict_math_sinh_zero() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.sinh(0.0));"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn strict_math_cosh_zero() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.cosh(0.0));"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn strict_math_tanh_zero() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.tanh(0.0));"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn strict_math_expm1_zero() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.expm1(0.0));"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn strict_math_log1p_zero() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.log1p(0.0));"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn strict_math_signum_positive() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.signum(42.0));"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn strict_math_signum_negative() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.signum(-7.0));"#);
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn strict_math_signum_zero() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.signum(0.0));"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn strict_math_rint_whole() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.rint(3.5));"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn strict_math_ceil_fraction() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.ceil(2.1));"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn strict_math_floor_fraction() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.floor(2.9));"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn strict_math_scalb_negative_base() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.scalb(-2.0, 3));"#);
    assert_eq!(out, vec!["-16"]);
}

#[test]
fn strict_math_next_up_double_max_finite() {
    let out = run_main(
        r#"System.out.println(java.lang.StrictMath.nextUp(java.lang.Double.MAX_VALUE) > java.lang.Double.MAX_VALUE);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn strict_math_ulp_of_double_min() {
    let out = run_main(
        r#"System.out.println(java.lang.StrictMath.ulp(java.lang.Double.MIN_VALUE) > 0.0);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn strict_math_get_exponent_subnormal() {
    let out = run_main(
        r#"System.out.println(java.lang.StrictMath.getExponent(java.lang.Double.MIN_VALUE) == java.lang.Double.MIN_EXPONENT - 1);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn strict_math_copy_sign_zero() {
    let out = run_main(r#"System.out.println(java.lang.StrictMath.copySign(0.0, -1.0));"#);
    assert_eq!(out, vec!["-0.0"]);
}

#[test]
fn strict_math_fma_zero_addend() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.fma(3.0, 4.0, 0.0));"#);
    assert_eq!(out, vec!["12"]);
}

#[test]
fn strict_math_ieee_remainder_exact() {
    let out =
        run_main(r#"System.out.println((int) java.lang.StrictMath.IEEEremainder(10.0, 5.0));"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn strict_math_to_degrees_pi() {
    let out = run_main(
        r#"System.out.println((int) java.lang.StrictMath.toDegrees(java.lang.StrictMath.PI));"#,
    );
    assert_eq!(out, vec!["180"]);
}

#[test]
fn strict_math_to_radians_180() {
    let out = run_main(r#"System.out.println((int) java.lang.StrictMath.toRadians(180.0));"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn strict_math_next_after_equal_direction() {
    let out = run_main(r#"System.out.println(java.lang.StrictMath.nextAfter(0.0, 0.0) == 0.0);"#);
    assert_eq!(out, vec!["true"]);
}
