/// Extended math builtins: trig, rounding, random, utility — distinct from test_math_extra.rs.
use super::helpers::run_pascal;

#[test]
fn abs_positive_integer() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Abs(7)); end."#),
        &["7"]
    );
}

#[test]
fn abs_negative_integer() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Abs(-7)); end."#),
        &["7"]
    );
}

#[test]
fn sqr_small_integer() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Sqr(6)); end."#),
        &["36"]
    );
}

#[test]
fn sqrt_perfect_square_16() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Sqrt(16.0):0:0); end."#),
        &["4"]
    );
}

#[test]
fn sqrt_perfect_square_81() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(Sqrt(81.0))); end."#),
        &["9"]
    );
}

#[test]
fn power_two_cubed() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Power(2.0, 3.0):0:0); end."#),
        &["8"]
    );
}

#[test]
fn power_ten_squared() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Power(10.0, 2.0):0:0); end."#),
        &["100"]
    );
}

#[test]
fn min_of_three_via_nested() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Min(Min(5,2),8)); end."#),
        &["2"]
    );
}

#[test]
fn max_of_three_via_nested() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Max(Max(5,2),8)); end."#),
        &["8"]
    );
}

#[test]
fn min_equal_returns_either() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Min(4,4)); end."#),
        &["4"]
    );
}

#[test]
fn max_equal_returns_either() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Max(4,4)); end."#),
        &["4"]
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
fn round_half_down_positive() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Round(2.4)); end."#),
        &["2"]
    );
}

#[test]
fn trunc_positive_fraction() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(9.99)); end."#),
        &["9"]
    );
}

#[test]
fn trunc_negative_fraction() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(-9.99)); end."#),
        &["-9"]
    );
}

#[test]
fn frac_zero_for_integer() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Frac(5.0)=0.0); end."#),
        &["true"]
    );
}

#[test]
fn int_negative_trunc_toward_zero() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Int(-2.9)); end."#),
        &["-2"]
    );
}

#[test]
fn sin_pi_over_six_half() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Format('%.1f', [Sin(Pi/6.0)])); end."#),
        &["0.5"]
    );
}

#[test]
fn cos_pi_over_three_half() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Format('%.1f', [Cos(Pi/3.0)])); end."#),
        &["0.5"]
    );
}

#[test]
fn tan_pi_over_four_one() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Format('%.0f', [Tan(Pi/4.0)])); end."#),
        &["1"]
    );
}

#[test]
fn arctan_one_is_pi_over_four() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Format('%.2f', [ArcTan(1.0)])); end."#),
        &["0.79"]
    );
}

#[test]
fn deg_to_rad_180_is_pi() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Format('%.2f', [DegToRad(180.0)])); end."#),
        &["3.14"]
    );
}

#[test]
fn rad_to_deg_pi_is_180() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Format('%.0f', [RadToDeg(Pi)])); end."#),
        &["180"]
    );
}

#[test]
fn hypot_5_12_is_13() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Format('%.0f', [Sqrt(5.0*5.0+12.0*12.0)])); end."#),
        &["13"]
    );
}

#[test]
fn inc_integer_by_one() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=5; Inc(n); WriteLn(n); end."#),
        &["6"]
    );
}

#[test]
fn inc_integer_by_delta() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=5; Inc(n,3); WriteLn(n); end."#),
        &["8"]
    );
}

#[test]
fn dec_integer_by_one() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=5; Dec(n); WriteLn(n); end."#),
        &["4"]
    );
}

#[test]
fn dec_integer_by_delta() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=10; Dec(n,4); WriteLn(n); end."#),
        &["6"]
    );
}

#[test]
fn pred_integer() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=5; n:=Pred(n); WriteLn(n); end."#),
        &["4"]
    );
}

#[test]
fn succ_integer() {
    assert_eq!(
        run_pascal(r#"program T; var n:Integer; begin n:=5; n:=Succ(n); WriteLn(n); end."#),
        &["6"]
    );
}

#[test]
fn odd_one_is_true() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Odd(1)); end."#),
        &["true"]
    );
}

#[test]
fn even_eight_is_true() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(not Odd(8)); end."#),
        &["true"]
    );
}

#[test]
fn exp_one_is_e() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Exp(1.0)>2.7); WriteLn(Exp(1.0)<2.8); end."#),
        &["true", "true"]
    );
}

#[test]
fn ln_e_is_one() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Format('%.0f', [Ln(Exp(1.0))])); end."#),
        &["1"]
    );
}

#[test]
fn random_bounded_below_six() {
    assert_eq!(
        run_pascal(
            r#"program T; var r:Integer; begin Randomize; r:=Random(6); WriteLn((r>=0) and (r<=5)); end."#
        ),
        &["true"]
    );
}

#[test]
fn random_modulo_bucket() {
    assert_eq!(
        run_pascal(r#"program T; var r:Integer; begin Randomize; r:=Random(1); WriteLn(r); end."#),
        &["0"]
    );
}

#[test]
fn ceil_via_neg_trunc() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(-Trunc(-3.2)); end."#),
        &["4"]
    );
}

#[test]
fn sign_positive_branch() {
    assert_eq!(
        run_pascal(
            r#"program T; function Sign(n:Integer):Integer; begin if n>0 then Result:=1 else if n<0 then Result:=-1 else Result:=0; end; begin WriteLn(Sign(9)); end."#
        ),
        &["1"]
    );
}

#[test]
fn sign_negative_branch() {
    assert_eq!(
        run_pascal(
            r#"program T; function Sign(n:Integer):Integer; begin if n>0 then Result:=1 else if n<0 then Result:=-1 else Result:=0; end; begin WriteLn(Sign(-2)); end."#
        ),
        &["-1"]
    );
}

#[test]
fn sign_zero_branch() {
    assert_eq!(
        run_pascal(
            r#"program T; function Sign(n:Integer):Integer; begin if n>0 then Result:=1 else if n<0 then Result:=-1 else Result:=0; end; begin WriteLn(Sign(0)); end."#
        ),
        &["0"]
    );
}

#[test]
fn mean_of_two_integers() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn((3+7) div 2); end."#),
        &["5"]
    );
}

#[test]
fn clamp_via_min_max() {
    assert_eq!(
        run_pascal(
            r#"program T; function Clamp(v,lo,hi:Integer):Integer; begin Result:=Max(lo,Min(v,hi)); end; begin WriteLn(Clamp(15,0,10)); end."#
        ),
        &["10"]
    );
}

#[test]
fn pythagorean_triple_check() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Sqr(3)+Sqr(4)=Sqr(5)); end."#),
        &["true"]
    );
}

#[test]
fn mod_always_non_negative_dividend() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(17 mod 5); end."#),
        &["2"]
    );
}

#[test]
fn div_integer_quotient() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(17 div 5); end."#),
        &["3"]
    );
}
