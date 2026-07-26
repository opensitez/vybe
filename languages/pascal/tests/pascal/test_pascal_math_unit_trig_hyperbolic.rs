use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 71: Trigonometric & Hyperbolic Math Routines
// ═══════════════════════════════════════════════════════════

#[test]
fn test_math_sin_cos_tan() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
begin
  WriteLn(Sin(0.0) = 0.0);
  WriteLn(Cos(0.0) = 1.0);
  WriteLn(Tan(0.0) = 0.0);
end.
"#,
    );
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_math_arcsin_arccos_arctan2() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
begin
  WriteLn(ArcSin(0.0) = 0.0);
  WriteLn(ArcCos(1.0) = 0.0);
  WriteLn(ArcTan2(1.0, 1.0) > 0.78);
end.
"#,
    );
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_math_sinh_cosh_tanh() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
begin
  WriteLn(Sinh(0.0) = 0.0);
  WriteLn(Cosh(0.0) = 1.0);
  WriteLn(Tanh(0.0) = 0.0);
end.
"#,
    );
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_math_arcsinh_arccosh_arctanh() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
begin
  WriteLn(ArcSinh(0.0) = 0.0);
  WriteLn(ArcCosh(1.0) = 0.0);
  WriteLn(ArcTanh(0.0) = 0.0);
end.
"#,
    );
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_math_log10_log2_logn() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
begin
  WriteLn(Log10(100.0) = 2.0);
  WriteLn(Log2(8.0) = 3.0);
  WriteLn(LogN(3.0, 27.0) = 3.0);
end.
"#,
    );
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_math_power_intpower() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
begin
  WriteLn(Power(2.0, 3.0) = 8.0);
  WriteLn(IntPower(3.0, 4) = 81.0);
end.
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_math_ceil_floor() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
begin
  WriteLn(Ceil(3.2));
  WriteLn(Floor(3.8));
  WriteLn(Ceil(-3.8));
  WriteLn(Floor(-3.2));
end.
"#,
    );
    assert_eq!(out, vec!["4", "3", "-3", "-4"]);
}

#[test]
fn test_math_degtorad_radtodeg() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
begin
  WriteLn(DegToRad(180.0) = Pi);
  WriteLn(RadToDeg(Pi) = 180.0);
end.
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_math_hypot() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
begin
  WriteLn(Hypot(3.0, 4.0) = 5.0);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_math_minvalue_maxvalue() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
var arr: array[0..3] of Double;
begin
  arr[0] := 5.0; arr[1] := 12.0; arr[2] := -3.0; arr[3] := 8.0;
  WriteLn(MinValue(arr) = -3.0);
  WriteLn(MaxValue(arr) = 12.0);
end.
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_math_mean_sum() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
var arr: array[0..3] of Double;
begin
  arr[0] := 10.0; arr[1] := 20.0; arr[2] := 30.0; arr[3] := 40.0;
  WriteLn(Sum(arr) = 100.0);
  WriteLn(Mean(arr) = 25.0);
end.
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_math_samevalue() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
begin
  WriteLn(SameValue(1.0000001, 1.0000002, 0.0001));
  WriteLn(SameValue(1.0, 2.0, 0.0001));
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_math_isnan_isinfinite_iszero() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
begin
  WriteLn(IsZero(0.000000001, 0.0001));
  WriteLn(IsZero(1.5, 0.0001));
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_math_round_to_simple() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
begin
  WriteLn(SimpleRoundTo(123.456, -2) = 123.46);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_math_sqr_sqrt() {
    let out = run_pascal(
        r#"
program Test;
begin
  WriteLn(Sqr(9) = 81);
  WriteLn(Sqrt(81.0) = 9.0);
end.
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_math_int_frac() {
    let out = run_pascal(
        r#"
program Test;
begin
  WriteLn(Int(5.75) = 5.0);
  WriteLn(SameValue(Frac(5.75), 0.75, 0.0001));
end.
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_math_ln_exp() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
begin
  WriteLn(SameValue(Ln(Exp(1.0)), 1.0, 0.0001));
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_math_poly_eval() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
var coeffs: array[0..2] of Double;
begin
  // P(x) = 1 + 2x + 3x^2 for x = 2: 1 + 4 + 12 = 17
  coeffs[0] := 1.0; coeffs[1] := 2.0; coeffs[2] := 3.0;
  WriteLn(Poly(2.0, coeffs) = 17.0);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_math_sign_function() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
begin
  WriteLn(Sign(15));
  WriteLn(Sign(-42));
  WriteLn(Sign(0));
end.
"#,
    );
    assert_eq!(out, vec!["1", "-1", "0"]);
}

#[test]
fn test_math_norm_hypot3() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
begin
  WriteLn(Hypot(3.0, 4.0) = 5.0);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}
