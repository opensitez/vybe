/// Extended type coercion, ordinal builtins, and numeric conversions.
use super::helpers::run_pascal;

#[test]
fn chr_ord_roundtrip_ascii() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Ord('A')); WriteLn(Chr(66)); end."#),
        &["65", "B"]
    );
}

#[test]
fn succ_pred_on_char() {
    assert_eq!(
        run_pascal(r#"program T; var c: Char; begin c:='m'; WriteLn(Succ(c)); WriteLn(Pred(c)); end."#),
        &["n", "l"]
    );
}

#[test]
fn succ_pred_on_integer() {
    assert_eq!(
        run_pascal(r#"program T; var n: Integer; begin n:=10; WriteLn(Succ(n)); WriteLn(Pred(n)); end."#),
        &["11", "9"]
    );
}

#[test]
fn ord_on_enum_value() {
    assert_eq!(
        run_pascal(
            r#"program T; type TS=(Zero,One,Two); var s: TS; begin s:=Two; WriteLn(Ord(s)); end."#
        ),
        &["2"]
    );
}

#[test]
fn trunc_toward_zero_positive() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(3.9)); end."#),
        &["3"]
    );
}

#[test]
fn trunc_toward_zero_negative() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(-3.9)); end."#),
        &["-3"]
    );
}

#[test]
fn frac_returns_fractional_part() {
    assert_eq!(
        run_pascal(r#"program T; var f: Double; begin f:=Frac(3.75); WriteLn(Round(f*100)); end."#),
        &["75"]
    );
}

#[test]
fn int_part_via_trunc_on_real() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trunc(12.01)); end."#),
        &["12"]
    );
}

#[test]
fn booltostr_default_format() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(BoolToStr(true)); WriteLn(BoolToStr(false)); end."#),
        &["True", "False"]
    );
}

#[test]
fn booltostr_lowercase_format() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(BoolToStr(true, true)); WriteLn(BoolToStr(false, true)); end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn strtobool_true_literal() {
    assert_eq!(
        run_pascal(r#"program T; var b: Boolean; begin b:=StrToBool('True'); WriteLn(b); end."#),
        &["true"]
    );
}

#[test]
fn strtobool_false_literal() {
    assert_eq!(
        run_pascal(r#"program T; var b: Boolean; begin b:=StrToBool('False'); WriteLn(b); end."#),
        &["false"]
    );
}

#[test]
fn strtoint_valid_digits() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(StrToInt('12345')); end."#),
        &["12345"]
    );
}

#[test]
fn strtoint_negative() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(StrToInt('-27')); end."#),
        &["-27"]
    );
}

#[test]
fn inttostr_zero() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(IntToStr(0)); end."#),
        &["0"]
    );
}

#[test]
fn floattostr_basic() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(FloatToStr(1.5)); end."#),
        &["1.5"]
    );
}

#[test]
fn strtofloat_basic() {
    assert_eq!(
        run_pascal(r#"program T; var r: Double; begin r:=StrToFloat('2.25'); WriteLn(r*4); end."#),
        &["9"]
    );
}

#[test]
fn val_procedure_integer_part() {
    assert_eq!(
        run_pascal(
            r#"program T; var s: string; v: Integer; code: Integer; begin s:='  42abc'; Val(s,v,code); WriteLn(v); WriteLn(code); end."#
        ),
        &["42", "3"]
    );
}

#[test]
fn val_procedure_invalid_prefix() {
    assert_eq!(
        run_pascal(
            r#"program T; var s: string; v: Integer; code: Integer; begin s:='x9'; Val(s,v,code); WriteLn(code); end."#
        ),
        &["1"]
    );
}

#[test]
fn chr_from_ord_zero() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Ord(Chr(0))); end."#),
        &["0"]
    );
}

#[test]
fn implicit_integer_to_real_addition() {
    assert_eq!(
        run_pascal(
            r#"program T; var r: Double; begin r:=2.5; r:=r+1; WriteLn(Round(r)); end."#
        ),
        &["4"]
    );
}

#[test]
fn integer_division_truncates() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(7 div 2); end."#),
        &["3"]
    );
}

#[test]
fn mod_on_negative_values() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(-7 mod 3); end."#),
        &["-1"]
    );
}

#[test]
fn round_bankers_or_half_up() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Round(2.5)); WriteLn(Round(3.5)); end."#),
        &["2", "4"]
    );
}

#[test]
fn extended_ascii_char_compare() {
    assert_eq!(
        run_pascal(r#"program T; begin if 'a'<'b' then WriteLn('less'); end."#),
        &["less"]
    );
}

#[test]
fn set_char_from_integer_ord() {
    assert_eq!(
        run_pascal(
            r#"program T; var c: Char; begin c:=Chr(90); WriteLn(c); end."#
        ),
        &["Z"]
    );
}

#[test]
fn enum_succ_wraps_within_type() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD=(Mon,Tue,Wed); var d: TD; begin d:=Mon; d:=Succ(d); WriteLn(Ord(d)); end."#
        ),
        &["1"]
    );
}

#[test]
fn hex_digit_via_inttohex_single() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(IntToHex(15,1)); end."#),
        &["F"]
    );
}

#[test]
fn string_to_int_via_inttostr_roundtrip() {
    assert_eq!(
        run_pascal(
            r#"program T; var n: Integer; s: string; begin n:=808; s:=IntToStr(n); WriteLn(StrToInt(s)); end."#
        ),
        &["808"]
    );
}
