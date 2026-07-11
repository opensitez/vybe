/// String formatting, Write/WriteLn combinations, and text output patterns.
use super::helpers::run_pascal;

#[test]
fn writeln_multiple_values_default_sep() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(1,2,3); end."#),
        &["1 2 3"]
    );
}

#[test]
fn write_without_ln_concatenates_same_line() {
    assert_eq!(
        run_pascal(r#"program T; begin Write('a'); Write('b'); WriteLn; end."#),
        &["ab"]
    );
}

#[test]
fn writeln_empty_produces_blank_line() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn; WriteLn('x'); end."#),
        &["", "x"]
    );
}

#[test]
fn format_float_two_decimals() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(FormatFloat('0.00', 3.1415)); end."#),
        &["3.14"]
    );
}

#[test]
fn format_int_zero_padded() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Format('%.4d', [7])); end."#),
        &["0007"]
    );
}

#[test]
fn format_string_placeholder() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Format('hello %s', ['world'])); end."#),
        &["hello world"]
    );
}

#[test]
fn format_multiple_args() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Format('%d+%d=%d', [2,3,5])); end."#),
        &["2+3=5"]
    );
}

#[test]
fn writeln_bool_default() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(true); WriteLn(false); end."#),
        &["true", "false"]
    );
}

#[test]
fn writeln_char_literal() {
    assert_eq!(run_pascal(r#"program T; begin WriteLn('Z'); end."#), &["Z"]);
}

#[test]
fn writeln_mixed_types() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn('n=', 42); end."#),
        &["n= 42"]
    );
}

#[test]
fn write_tab_char_separator() {
    assert_eq!(
        run_pascal(r#"program T; begin Write('a'); Write(#9); WriteLn('b'); end."#),
        &["a\tb"]
    );
}

#[test]
fn writeln_nested_call_result() {
    assert_eq!(
        run_pascal(
            r#"program T; function Twice(n:Integer):Integer; begin Result:=n*2; end; begin WriteLn(Twice(6)); end."#
        ),
        &["12"]
    );
}

#[test]
fn format_datetime_style_string() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Format('%s-%s', ['2024', '06'])); end."#),
        &["2024-06"]
    );
}

#[test]
fn writeln_in_loop_counter() {
    assert_eq!(
        run_pascal(r#"program T; var i:Integer; begin for i:=1 to 3 do WriteLn(i); end."#),
        &["1", "2", "3"]
    );
}

#[test]
fn write_hex_via_inttohex() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn('0x'+IntToHex(255,2)); end."#),
        &["0xFF"]
    );
}

#[test]
fn writeln_string_concat_expression() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn('a'+'b'+'c'); end."#),
        &["abc"]
    );
}

#[test]
fn format_fixed_width_string() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Format('[%s]', ['ok'])); end."#),
        &["[ok]"]
    );
}

#[test]
fn writeln_real_default_format() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(1.5); end."#),
        &["1.5"]
    );
}

#[test]
fn write_flush_style_partial_lines() {
    assert_eq!(
        run_pascal(r#"program T; begin Write('wait'); WriteLn(' ok'); end."#),
        &["wait ok"]
    );
}

#[test]
fn format_percent_literal_escape() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Format('100%%', [])); end."#),
        &["100%"]
    );
}
