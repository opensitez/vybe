/// Tests for additional math functions in Pascal/Delphi:
/// Odd/Even predicates, Int/Frac, Exp/Ln, trig functions,
/// extended arithmetic patterns not covered in test_builtins.rs.
use super::helpers::run_pascal;

// ===================================================================
// ODD / EVEN
// ===================================================================

#[test]
fn odd_true() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Odd(7));
end."#
        ),
        &["TRUE"]
    );
}

#[test]
fn odd_false() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Odd(4));
end."#
        ),
        &["FALSE"]
    );
}

#[test]
fn odd_zero() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Odd(0));
end."#
        ),
        &["FALSE"]
    );
}

#[test]
fn odd_negative() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Odd(-3));
end."#
        ),
        &["TRUE"]
    );
}

#[test]
fn odd_in_loop() {
    assert_eq!(
        run_pascal(
            r#"program T;
var i: Integer;
begin
  for i := 1 to 5 do
    if Odd(i) then Write(IntToStr(i) + ' ');
  WriteLn('');
end."#
        ),
        &["1 3 5 "]
    );
}

// ===================================================================
// INT AND FRAC
// ===================================================================

#[test]
fn int_of_positive() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Int(3.7));
end."#
        ),
        &["3"]
    );
}

#[test]
fn int_of_negative() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Int(-3.7));
end."#
        ),
        &["-3"]
    );
}

#[test]
fn frac_of_real() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Frac(3.25) = 0.25);
end."#
        ),
        &["TRUE"]
    );
}

#[test]
fn int_whole_number() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Int(5.0));
end."#
        ),
        &["5"]
    );
}

// ===================================================================
// EXP AND LN
// ===================================================================

#[test]
fn exp_of_zero() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Exp(0));
end."#
        ),
        &["1"]
    );
}

#[test]
fn ln_of_one() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Ln(1.0));
end."#
        ),
        &["0"]
    );
}

#[test]
fn exp_ln_roundtrip() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Round(Exp(Ln(5.0))));
end."#
        ),
        &["5"]
    );
}

// ===================================================================
// TRIGONOMETRIC FUNCTIONS
// ===================================================================

#[test]
fn sin_of_zero() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Sin(0.0));
end."#
        ),
        &["0"]
    );
}

#[test]
fn cos_of_zero() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Cos(0.0));
end."#
        ),
        &["1"]
    );
}

#[test]
fn arctan_of_zero() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(ArcTan(0.0));
end."#
        ),
        &["0"]
    );
}

#[test]
fn pi_constant() {
    assert_eq!(
        run_pascal(
            r#"program T;
const Pi = 3.14159265;
begin
  WriteLn(Pi > 3.0);
  WriteLn(Pi < 4.0);
end."#
        ),
        &["TRUE", "TRUE"]
    );
}

// ===================================================================
// SQRT EXTENDED
// ===================================================================

#[test]
fn sqrt_of_nine() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Sqrt(9.0));
end."#
        ),
        &["3"]
    );
}

#[test]
fn sqrt_of_two_approx() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Sqrt(2.0) > 1.4);
  WriteLn(Sqrt(2.0) < 1.5);
end."#
        ),
        &["TRUE", "TRUE"]
    );
}

// ===================================================================
// ABS EXTENDED
// ===================================================================

#[test]
fn abs_real_value() {
    assert_eq!(
        run_pascal(
            r#"program T;
var x: Real;
begin
  x := -2.5;
  WriteLn(Abs(x));
end."#
        ),
        &["2.5"]
    );
}

#[test]
fn sqr_of_three() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Sqr(3));
end."#
        ),
        &["9"]
    );
}

// -------------------------------------------------------------------
// from test_math_power_roots.rs
// -------------------------------------------------------------------
#[test]
fn sqrt_of_perfect_square_nine() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Sqrt(9.0):0:0); end."#),
        &["3"]
    );
}

#[test]
fn sqrt_of_two_approximate() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Format('%.3f', [Sqrt(2.0)])); end."#),
        &["1.414"]
    );
}

#[test]
fn int_returns_integer_part_of_real() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Int(3.9)); end."#),
        &["3"]
    );
}

#[test]
fn frac_returns_fractional_part() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Format('%.1f', [Frac(3.9)])); end."#),
        &["0.9"]
    );
}

#[test]
fn round_half_up_positive() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Round(2.6)); end."#),
        &["3"]
    );
}

#[test]
fn trunc_toward_zero_negative() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(-2.9)); end."#),
        &["-2"]
    );
}

#[test]
fn abs_integer_negative() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Abs(-17)); end."#),
        &["17"]
    );
}

#[test]
fn abs_real_negative() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Format('%.1f', [Abs(-4.5)])); end."#),
        &["4.5"]
    );
}

#[test]
fn sqr_of_negative_integer() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Sqr(-6)); end."#),
        &["36"]
    );
}

#[test]
fn power_integer_exponent_via_loop() {
    assert_eq!(
        run_pascal(
            r#"program T;
function PowInt(base, exp: Integer): Integer;
var i, r: Integer;
begin
  r := 1;
  for i := 1 to exp do r := r * base;
  Result := r;
end;
begin
  WriteLn(PowInt(2, 10));
end."#
        ),
        &["1024"]
    );
}

#[test]
fn hypot_three_four_five_triangle() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Hypot(a, b: Real): Real;
begin
  Result := Sqrt(a * a + b * b);
end;
begin
  WriteLn(Hypot(3.0, 4.0):0:0);
end."#
        ),
        &["5"]
    );
}

#[test]
fn mod_wrapping_positive_modulus() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(17 mod 5); end."#),
        &["2"]
    );
}

#[test]
fn div_integer_division_truncates() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(17 div 5); end."#),
        &["3"]
    );
}

#[test]
fn sin_quarter_pi_one() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Format('%.3f', [Sin(Pi / 2.0)])); end."#),
        &["1.000"]
    );
}

#[test]
fn cos_pi_negative_one() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Format('%.0f', [Cos(Pi)])); end."#),
        &["-1"]
    );
}

#[test]
fn tan_zero_is_zero() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Format('%.1f', [Tan(0.0)])); end."#),
        &["0.0"]
    );
}

#[test]
fn deg_to_rad_converts_ninety() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Format('%.4f', [DegToRad(90.0)])); end."#),
        &["1.5708"]
    );
}

#[test]
fn rad_to_deg_converts_pi() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Format('%.1f', [RadToDeg(Pi)])); end."#),
        &["180.0"]
    );
}

#[test]
fn power_function_integer_exponent() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Power(2.0, 10.0):0:0); end."#),
        &["1024"]
    );
}

#[test]
fn min_of_two_integers() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Min(3, 8)); end."#),
        &["3"]
    );
}

#[test]
fn max_of_two_integers() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Max(3, 8)); end."#),
        &["8"]
    );
}

#[test]
fn inc_procedure_mutates_integer() {
    assert_eq!(
        run_pascal(r#"program T; var n: Integer; begin n := 5; Inc(n, 2); WriteLn(n); end."#),
        &["7"]
    );
}

#[test]
fn dec_procedure_mutates_integer() {
    assert_eq!(
        run_pascal(r#"program T; var n: Integer; begin n := 5; Dec(n, 3); WriteLn(n); end."#),
        &["2"]
    );
}

#[test]
fn abs_negative_integer() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Abs(-9)); end."#),
        &["9"]
    );
}

#[test]
fn sqrt_of_perfect_square() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(Sqrt(81.0))); end."#),
        &["9"]
    );
}

#[test]
fn round_half_up_to_integer() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Round(2.6)); end."#),
        &["3"]
    );
}

#[test]
fn trunc_toward_zero_for_positive() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(3.9)); end."#),
        &["3"]
    );
}

#[test]
fn int_truncates_toward_zero() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Int(-3.9)); end."#),
        &["-3"]
    );
}
