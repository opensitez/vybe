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
        &["true"]
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
        &["false"]
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
        &["false"]
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
        &["true"]
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
        &["true"]
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
        &["true", "true"]
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
        &["true", "true"]
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
