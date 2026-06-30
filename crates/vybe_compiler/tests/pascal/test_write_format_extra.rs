/// Format/WriteLn width, precision, and output patterns.
use super::helpers::run_pascal;

#[test]
fn format_int_zero_pad_four() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Format('%.4d',[7])); end."#
        ),
        &["0007"]
    );
}

#[test]
fn format_int_zero_pad_two() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Format('%.2d',[42])); end."#
        ),
        &["42"]
    );
}

#[test]
fn format_int_width_eight() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Format('%8d',[123])); end."#
        ),
        &["     123"]
    );
}

#[test]
fn format_string_left() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Format('%-5s',['hi'])); end."#
        ),
        &["hi   "]
    );
}

#[test]
fn format_float_two_decimals() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(FormatFloat('0.00',3.1415)); end."#
        ),
        &["3.14"]
    );
}

#[test]
fn format_float_one_decimal() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(FormatFloat('0.0',2.55)); end."#
        ),
        &["2.6"]
    );
}

#[test]
fn format_float_fixed_width() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(FormatFloat('00.00',1.2)); end."#
        ),
        &["01.20"]
    );
}

#[test]
fn writeln_multiple_ints() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(1,2,3); end."#
        ),
        &["1 2 3"]
    );
}

#[test]
fn writeln_mixed_types() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn('n',42); end."#
        ),
        &["n42"]
    );
}

#[test]
fn write_concat_no_ln() {
    assert_eq!(
        run_pascal(
            r#"program T; begin Write('a'); Write('b'); WriteLn; end."#
        ),
        &["ab"]
    );
}

#[test]
fn writeln_empty_line() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn; WriteLn('x'); end."#
        ),
        &["", "x"]
    );
}

#[test]
fn format_percent_d() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Format('%d',[99])); end."#
        ),
        &["99"]
    );
}

#[test]
fn format_percent_s_twice() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Format('%s-%s',['foo','bar'])); end."#
        ),
        &["foo-bar"]
    );
}

#[test]
fn format_three_placeholders() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Format('%d/%d=%d',[10,2,5])); end."#
        ),
        &["10/2=5"]
    );
}

#[test]
fn format_float_percent_f() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Format('%.1f',[2.5])); end."#
        ),
        &["2.5"]
    );
}

#[test]
fn format_hex_style() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(IntToHex(255,2)); end."#
        ),
        &["FF"]
    );
}

#[test]
fn writeln_bool_values() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(true); WriteLn(false); end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn writeln_char_value() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn('Z'); end."#
        ),
        &["Z"]
    );
}

#[test]
fn format_padding_zeros() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Format('%.6d',[12])); end."#
        ),
        &["000012"]
    );
}

#[test]
fn format_string_and_int() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Format('id=%d',[7])); end."#
        ),
        &["id=7"]
    );
}

#[test]
fn writeln_real_default() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(1.5); end."#
        ),
        &["1.5"]
    );
}

#[test]
fn format_date_template() {
    assert_eq!(
        run_pascal(
            r#"program T; var d:TDateTime; begin d:=EncodeDate(2021,6,15); WriteLn(FormatDateTime('yyyy-mm-dd',d)); end."#
        ),
        &["2021-06-15"]
    );
}

#[test]
fn format_time_template() {
    assert_eq!(
        run_pascal(
            r#"program T; var t:TDateTime; begin t:=EncodeTime(9,5,7,0); WriteLn(FormatDateTime('hh:nn',t)); end."#
        ),
        &["09:05"]
    );
}

#[test]
fn write_ln_separate_calls() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn('a'); WriteLn('b'); end."#
        ),
        &["a", "b"]
    );
}

#[test]
fn format_right_align() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Format('%5s',['x'])); end."#
        ),
        &["    x"]
    );
}

#[test]
fn format_multiple_ints_array() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Format('%d %d %d',[1,2,3])); end."#
        ),
        &["1 2 3"]
    );
}

#[test]
fn format_float_zero() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(FormatFloat('0.00',0.0)); end."#
        ),
        &["0.00"]
    );
}

#[test]
fn format_large_integer() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Format('%d',[1000000])); end."#
        ),
        &["1000000"]
    );
}

#[test]
fn writeln_integer_hex() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(IntToStr($2A)); end."#
        ),
        &["42"]
    );
}

#[test]
fn format_nested_concat() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Format('(%d)',[5])); end."#
        ),
        &["(5)"]
    );
}

#[test]
fn write_partial_line() {
    assert_eq!(
        run_pascal(
            r#"program T; begin Write('>>'); WriteLn('<<'); end."#
        ),
        &[">><<"]
    );
}

#[test]
fn format_width_star_style() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Format('%10d',[7])); end."#
        ),
        &["         7"]
    );
}

#[test]
fn format_negative_int() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Format('%d',[-8])); end."#
        ),
        &["-8"]
    );
}

#[test]
fn format_percent_literal() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Format('100%%',[ ])); end."#
        ),
        &["100%"]
    );
}

#[test]
fn writeln_empty_string() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(''); WriteLn('ok'); end."#
        ),
        &["", "ok"]
    );
}

#[test]
fn format_float_scientific_style() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(FormatFloat('0.000',0.5)); end."#
        ),
        &["0.500"]
    );
}

#[test]
fn format_two_strings() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Format('%s%s',['ab','cd'])); end."#
        ),
        &["abcd"]
    );
}

#[test]
fn writeln_ord_char() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Ord('A')); end."#
        ),
        &["65"]
    );
}

#[test]
fn format_combined_text() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Format('val=%.2f',[3.5])); end."#
        ),
        &["val=3.50"]
    );
}

#[test]
fn write_tab_separated() {
    assert_eq!(
        run_pascal(
            r#"program T; begin Write('a'); Write(#9); WriteLn('b'); end."#
        ),
        &["a\tb"]
    );
}

#[test]
fn format_leading_zeros_year() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Format('%.4d',[7])); end."#
        ),
        &["0007"]
    );
}

#[test]
fn writeln_trunc_real() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Trunc(9.9)); end."#
        ),
        &["9"]
    );
}

