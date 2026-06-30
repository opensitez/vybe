/// Sin, Cos, Tan, and Arctan2 distinct trigonometric cases.
use super::helpers::run_pascal;

#[test]
fn sin_zero_is_zero() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(Sin(0.0)*1000)); end."#),
        &["0"]
    );
}

#[test]
fn sin_pi_over_two_one() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(Sin(1.5707963)*1000)); end."#),
        &["1000"]
    );
}

#[test]
fn sin_pi_is_near_zero() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(Abs(Sin(3.1415926))*1000)); end."#),
        &["0"]
    );
}

#[test]
fn sin_negative_angle() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(Sin(-0.5235988)*1000)); end."#),
        &["-500"]
    );
}

#[test]
fn cos_zero_is_one() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(Cos(0.0)*1000)); end."#),
        &["1000"]
    );
}

#[test]
fn cos_pi_over_two_near_zero() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(Abs(Cos(1.5707963))*1000)); end."#),
        &["0"]
    );
}

#[test]
fn cos_pi_negative_one() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(Cos(3.1415926)*1000)); end."#),
        &["-1000"]
    );
}

#[test]
fn cos_pi_over_three_half() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(Cos(1.0471976)*1000)); end."#),
        &["500"]
    );
}

#[test]
fn tan_zero_is_zero() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(Tan(0.0)*1000)); end."#),
        &["0"]
    );
}

#[test]
fn tan_pi_over_four_one() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(Tan(0.7853982)*1000)); end."#),
        &["1000"]
    );
}

#[test]
fn tan_small_angle_approx() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(Tan(0.1)*1000)); end."#),
        &["100"]
    );
}

#[test]
fn arctan_zero_zero() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(ArcTan(0.0)*1000)); end."#),
        &["0"]
    );
}

#[test]
fn arctan_one_pi_over_four() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(ArcTan(1.0)*1000)); end."#),
        &["785"]
    );
}

#[test]
fn arctan_negative_input() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(ArcTan(-1.0)*1000)); end."#),
        &["-785"]
    );
}

#[test]
fn arctan_large_value_near_pi_over_two() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(ArcTan(1000.0)*100)); end."#),
        &["157"]
    );
}

#[test]
fn arctan2_origin_zero() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(ArcTan2(0.0,0.0)*1000)); end."#),
        &["0"]
    );
}

#[test]
fn arctan2_positive_x_axis() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(ArcTan2(0.0,1.0)*1000)); end."#),
        &["0"]
    );
}

#[test]
fn arctan2_positive_y_axis() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(ArcTan2(1.0,0.0)*1000)); end."#),
        &["1570"]
    );
}

#[test]
fn arctan2_negative_x_axis() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(Abs(ArcTan2(0.0,-1.0))*1000)); end."#),
        &["3141"]
    );
}

#[test]
fn arctan2_quadrant_one() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(ArcTan2(1.0,1.0)*1000)); end."#),
        &["785"]
    );
}

#[test]
fn arctan2_quadrant_two() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(ArcTan2(1.0,-1.0)*100)); end."#),
        &["233"]
    );
}

#[test]
fn arctan2_quadrant_three() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(Abs(ArcTan2(-1.0,-1.0))*100)); end."#),
        &["233"]
    );
}

#[test]
fn arctan2_quadrant_four() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(ArcTan2(-1.0,1.0)*1000)); end."#),
        &["-785"]
    );
}

#[test]
fn sin_cos_pythagorean_unit() {
    assert_eq!(
        run_pascal(r#"program T; var t:Real; begin t:=0.7; WriteLn(Trunc((Sin(t)*Sin(t)+Cos(t)*Cos(t))*1000)); end."#),
        &["1000"]
    );
}

#[test]
fn sin_double_angle_identity() {
    assert_eq!(
        run_pascal(r#"program T; var a:Real; begin a:=0.4; WriteLn(Trunc((Sin(2*a)-2*Sin(a)*Cos(a))*10000)); end."#),
        &["0"]
    );
}

#[test]
fn tan_sin_over_cos() {
    assert_eq!(
        run_pascal(r#"program T; var a:Real; begin a:=0.3; WriteLn(Trunc((Tan(a)-Sin(a)/Cos(a))*10000)); end."#),
        &["0"]
    );
}

#[test]
fn arctan_tan_roundtrip_small() {
    assert_eq!(
        run_pascal(r#"program T; var a:Real; begin a:=0.2; WriteLn(Trunc((ArcTan(Tan(a))-a)*10000)); end."#),
        &["0"]
    );
}

#[test]
fn sin_pi_over_six_half() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(Sin(0.5235988)*1000)); end."#),
        &["500"]
    );
}

#[test]
fn cos_pi_over_six_sqrt_three_over_two() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(Cos(0.5235988)*1000)); end."#),
        &["866"]
    );
}

#[test]
fn sin_three_pi_over_two_negative_one() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(Sin(4.71238898)*1000)); end."#),
        &["-1000"]
    );
}

#[test]
fn cos_two_pi_is_one() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(Cos(6.2831853)*1000)); end."#),
        &["1000"]
    );
}

#[test]
fn tan_negative_angle() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(Tan(-0.7853982)*1000)); end."#),
        &["-1000"]
    );
}

#[test]
fn arctan2_equal_coords() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(ArcTan2(5.0,5.0)*1000)); end."#),
        &["785"]
    );
}

#[test]
fn arctan2_y_zero_positive_x() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(ArcTan2(0.0,3.0)*1000)); end."#),
        &["0"]
    );
}

#[test]
fn arctan2_y_zero_negative_x() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(Abs(ArcTan2(0.0,-3.0))*100)); end."#),
        &["314"]
    );
}

#[test]
fn sin_half_pi_minus_x_equals_cos() {
    assert_eq!(
        run_pascal(r#"program T; var x:Real; begin x:=0.6; WriteLn(Trunc((Sin(1.5707963-x)-Cos(x))*10000)); end."#),
        &["0"]
    );
}

#[test]
fn cos_half_pi_minus_x_equals_sin() {
    assert_eq!(
        run_pascal(r#"program T; var x:Real; begin x:=0.6; WriteLn(Trunc((Cos(1.5707963-x)-Sin(x))*10000)); end."#),
        &["0"]
    );
}

#[test]
fn arctan_fraction_less_than_one() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(ArcTan(0.5)*1000)); end."#),
        &["463"]
    );
}

#[test]
fn arctan2_small_y_large_x() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(ArcTan2(0.1,10.0)*10000)); end."#),
        &["999"]
    );
}

#[test]
fn sin_squared_plus_cos_squared_various() {
    assert_eq!(
        run_pascal(r#"program T; var t:Real; begin t:=1.1; WriteLn(Trunc((Sin(t)*Sin(t)+Cos(t)*Cos(t))*1000)); end."#),
        &["1000"]
    );
}

#[test]
fn tan_arctan_inverse_near_one() {
    assert_eq!(
        run_pascal(r#"program T; var v:Real; begin v:=2.0; WriteLn(Trunc((Tan(ArcTan(v))-v)*10000)); end."#),
        &["0"]
    );
}

#[test]
fn arctan2_negative_y_positive_x() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(ArcTan2(-2.0,2.0)*1000)); end."#),
        &["-785"]
    );
}
