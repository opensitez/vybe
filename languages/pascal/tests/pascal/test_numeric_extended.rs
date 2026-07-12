/// Int64, Currency, Extended, and rounding edge cases.
use super::helpers::run_pascal;

#[test]
fn int64_large_positive_literal() {
    assert_eq!(
        run_pascal(r#"program T; var i:Int64; begin i:=9223372036854775807; WriteLn(i>0); end."#),
        &["true"]
    );
}

#[test]
fn int64_negative_arithmetic() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Int64; begin a:=-1000000; b:=2000000; WriteLn(a+b); end."#
        ),
        &["1000000"]
    );
}

#[test]
fn int64_multiply_small_factors() {
    assert_eq!(
        run_pascal(r#"program T; var i:Int64; begin i:=1000; WriteLn(i*1000); end."#),
        &["1000000"]
    );
}

#[test]
fn int64_div_integer_truncates() {
    assert_eq!(
        run_pascal(r#"program T; var i:Int64; begin i:=7; WriteLn(i div 2); end."#),
        &["3"]
    );
}

#[test]
fn int64_mod_operation() {
    assert_eq!(
        run_pascal(r#"program T; var i:Int64; begin i:=10; WriteLn(i mod 3); end."#),
        &["1"]
    );
}

#[test]
fn int64_inc_by_delta() {
    assert_eq!(
        run_pascal(r#"program T; var i:Int64; begin i:=100; Inc(i,50); WriteLn(i); end."#),
        &["150"]
    );
}

#[test]
fn int64_in_record_field() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record Total:Int64; end; var r:TR; begin r.Total:=5000000000; WriteLn(r.Total>4000000000); end."#
        ),
        &["true"]
    );
}

#[test]
fn currency_add_two_amounts() {
    assert_eq!(
        run_pascal(r#"program T; var a,b:Currency; begin a:=1.25; b:=2.50; WriteLn(a+b); end."#),
        &["3.75"]
    );
}

#[test]
fn currency_multiply_quantity() {
    assert_eq!(
        run_pascal(
            r#"program T; var price,qty,total:Currency; begin price:=9.99; qty:=3; total:=price*qty; WriteLn(total>29.0); end."#
        ),
        &["true"]
    );
}

#[test]
fn currency_round_half_up() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Currency; begin c:=12.345; WriteLn(Round(c*100)/100); end."#
        ),
        &["12.35"]
    );
}

#[test]
fn currency_compare_equality() {
    assert_eq!(
        run_pascal(r#"program T; var a,b:Currency; begin a:=10.00; b:=10.00; WriteLn(a=b); end."#),
        &["true"]
    );
}

#[test]
fn currency_negative_balance() {
    assert_eq!(
        run_pascal(r#"program T; var c:Currency; begin c:=-5.50; WriteLn(c<0); end."#),
        &["true"]
    );
}

#[test]
fn extended_high_precision_add() {
    assert_eq!(
        run_pascal(
            r#"program T; var e:Extended; begin e:=1.23456789; e:=e+0.00000001; WriteLn(e>1.234567); end."#
        ),
        &["true"]
    );
}

#[test]
fn extended_division_small_fraction() {
    assert_eq!(
        run_pascal(r#"program T; var e:Extended; begin e:=1.0/3.0; WriteLn(e>0.333); end."#),
        &["true"]
    );
}

#[test]
fn extended_abs_negative() {
    assert_eq!(
        run_pascal(r#"program T; var e:Extended; begin e:=-99.5; WriteLn(Abs(e)); end."#),
        &["99.5"]
    );
}

#[test]
fn round_half_to_even_at_point_five() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Round(2.5)); end."#),
        &["2"]
    );
}

#[test]
fn round_negative_half() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Round(-2.5)); end."#),
        &["-2"]
    );
}

#[test]
fn trunc_toward_zero_positive() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(9.99)); end."#),
        &["9"]
    );
}

#[test]
fn trunc_toward_zero_negative() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(-9.99)); end."#),
        &["-9"]
    );
}

#[test]
fn floor_large_negative_fraction() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Floor(-1.1)); end."#),
        &["-2"]
    );
}

#[test]
fn ceil_negative_fraction() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Ceil(-1.9)); end."#),
        &["-1"]
    );
}

#[test]
fn frac_returns_fractional_part() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Frac(3.75)>0.7); end."#),
        &["true"]
    );
}

#[test]
fn int64_to_str_and_back() {
    assert_eq!(
        run_pascal(
            r#"program T; var i:Int64; s:String; begin s:='1234567890123'; i:=StrToInt64(s); WriteLn(i); end."#
        ),
        &["1234567890123"]
    );
}

#[test]
fn currency_format_two_decimals() {
    assert_eq!(
        run_pascal(r#"program T; var c:Currency; begin c:=3.5; WriteLn(Format('%.2f',[c])); end."#),
        &["3.50"]
    );
}

#[test]
fn extended_compare_with_double() {
    assert_eq!(
        run_pascal(
            r#"program T; var e:Extended; d:Double; begin e:=1.0; d:=1.0; WriteLn(e=d); end."#
        ),
        &["true"]
    );
}

#[test]
fn round_bankers_three_point_five() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Round(3.5)); end."#),
        &["4"]
    );
}

#[test]
fn round_zero_unchanged() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Round(0.0)); end."#),
        &["0"]
    );
}

#[test]
fn int64_shl_small() {
    assert_eq!(
        run_pascal(r#"program T; var i:Int64; begin i:=1; i:=i shl 10; WriteLn(i); end."#),
        &["1024"]
    );
}

#[test]
fn int64_shr_small() {
    assert_eq!(
        run_pascal(r#"program T; var i:Int64; begin i:=1024; i:=i shr 10; WriteLn(i); end."#),
        &["1"]
    );
}

#[test]
fn currency_subtract_tax() {
    assert_eq!(
        run_pascal(
            r#"program T; var gross,tax,net:Currency; begin gross:=100.00; tax:=8.25; net:=gross-tax; WriteLn(net); end."#
        ),
        &["91.75"]
    );
}

#[test]
fn extended_sqrt_two_approx() {
    assert_eq!(
        run_pascal(r#"program T; var e:Extended; begin e:=Sqrt(2.0); WriteLn(e>1.414); end."#),
        &["true"]
    );
}

#[test]
fn round_large_real_value() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Round(999999.49)); end."#),
        &["999999"]
    );
}

#[test]
fn trunc_zero_fraction() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(42.0)); end."#),
        &["42"]
    );
}

#[test]
fn floor_exact_integer_real() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Floor(5.0)); end."#),
        &["5"]
    );
}

#[test]
fn ceil_exact_integer_real() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Ceil(5.0)); end."#),
        &["5"]
    );
}

#[test]
fn int64_min_comparison() {
    assert_eq!(
        run_pascal(r#"program T; var a,b:Int64; begin a:=-5; b:=-10; WriteLn(Min(a,b)); end."#),
        &["-10"]
    );
}

#[test]
fn int64_max_comparison() {
    assert_eq!(
        run_pascal(r#"program T; var a,b:Int64; begin a:=100; b:=200; WriteLn(Max(a,b)); end."#),
        &["200"]
    );
}

#[test]
fn currency_int_part_via_trunc() {
    assert_eq!(
        run_pascal(r#"program T; var c:Currency; begin c:=19.99; WriteLn(Trunc(c)); end."#),
        &["19"]
    );
}

#[test]
fn extended_power_small_exp() {
    assert_eq!(
        run_pascal(r#"program T; var e:Extended; begin e:=Power(2.0,10.0); WriteLn(e); end."#),
        &["1024"]
    );
}

#[test]
fn round_tie_negative_three_point_five() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Round(-3.5)); end."#),
        &["-4"]
    );
}

#[test]
fn int64_abs_negative_large() {
    assert_eq!(
        run_pascal(r#"program T; var i:Int64; begin i:=-999999999; WriteLn(Abs(i)); end."#),
        &["999999999"]
    );
}
