use super::helpers::run_js;

#[test]
fn test_math_sin_zero() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.sin(0));
"#
        ),
        vec!["0"]
    );
}

#[test]
fn test_math_cos_zero() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.cos(0));
"#
        ),
        vec!["1"]
    );
}

#[test]
fn test_math_tan_zero() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.tan(0));
"#
        ),
        vec!["0"]
    );
}

#[test]
fn test_math_asin_one() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.asin(1).toFixed(4));
"#
        ),
        vec!["1.5708"]
    );
}

#[test]
fn test_math_acos_one() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.acos(1));
"#
        ),
        vec!["0"]
    );
}

#[test]
fn test_math_atan_one() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.atan(1).toFixed(4));
"#
        ),
        vec!["0.7854"]
    );
}

#[test]
fn test_math_atan2_one_one() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.atan2(1, 1).toFixed(4));
"#
        ),
        vec!["0.7854"]
    );
}

#[test]
fn test_math_atan2_zero_neg_one() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.atan2(0, -1).toFixed(4));
"#
        ),
        vec!["3.1416"]
    );
}

#[test]
fn test_math_sinh_zero() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.sinh(0));
"#
        ),
        vec!["0"]
    );
}

#[test]
fn test_math_cosh_zero() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.cosh(0));
"#
        ),
        vec!["1"]
    );
}

#[test]
fn test_math_tanh_zero() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.tanh(0));
"#
        ),
        vec!["0"]
    );
}

#[test]
fn test_math_exp_one() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.exp(1).toFixed(4));
"#
        ),
        vec!["2.7183"]
    );
}

#[test]
fn test_math_exp_zero() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.exp(0));
"#
        ),
        vec!["1"]
    );
}

#[test]
fn test_math_log_one() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.log(1));
"#
        ),
        vec!["0"]
    );
}

#[test]
fn test_math_log_e() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.log(Math.E).toFixed(4));
"#
        ),
        vec!["1.0000"]
    );
}

#[test]
fn test_math_log2_eight() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.log2(8));
"#
        ),
        vec!["3"]
    );
}

#[test]
fn test_math_log10_thousand() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.log10(1000));
"#
        ),
        vec!["3"]
    );
}

#[test]
fn test_math_expm1_zero() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.expm1(0));
"#
        ),
        vec!["0"]
    );
}

#[test]
fn test_math_log1p_zero() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.log1p(0));
"#
        ),
        vec!["0"]
    );
}

#[test]
fn test_math_hypot_three_four() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.hypot(3, 4));
"#
        ),
        vec!["5"]
    );
}

#[test]
fn test_math_cbrt_twentyseven() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.cbrt(27));
"#
        ),
        vec!["3"]
    );
}

#[test]
fn test_math_clz32_one() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.clz32(1));
"#
        ),
        vec!["31"]
    );
}

#[test]
fn test_math_clz32_zero() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.clz32(0));
"#
        ),
        vec!["32"]
    );
}

#[test]
fn test_math_imul() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.imul(3, 4));
"#
        ),
        vec!["12"]
    );
}

#[test]
fn test_math_fround_rounding() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.fround(1.337).toFixed(4));
"#
        ),
        vec!["1.3370"]
    );
}

#[test]
fn test_math_sign() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.sign(-5));
console.log(Math.sign(0));
console.log(Math.sign(5));
"#
        ),
        vec!["-1", "0", "1"]
    );
}

#[test]
fn test_math_trunc() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.trunc(4.9));
console.log(Math.trunc(-4.9));
"#
        ),
        vec!["4", "-4"]
    );
}

#[test]
fn test_math_max_spread() {
    assert_eq!(
        run_js(
            r#"
let nums = [3, 1, 4, 1, 5, 9, 2, 6];
console.log(Math.max(...nums));
"#
        ),
        vec!["9"]
    );
}

#[test]
fn test_math_min_spread() {
    assert_eq!(
        run_js(
            r#"
let nums = [3, 1, 4, 1, 5, 9, 2, 6];
console.log(Math.min(...nums));
"#
        ),
        vec!["1"]
    );
}

#[test]
fn test_math_pi_constant() {
    assert_eq!(
        run_js(
            r#"
console.log(Math.PI.toFixed(5));
"#
        ),
        vec!["3.14159"]
    );
}
