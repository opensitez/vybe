use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 92: Advanced Format Specifiers (Format, FormatBuf)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_fmt_decimal_hex_uppercase() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(Format('%d - %X', [255, 255]));
end.
"#,
    );
    assert_eq!(out, vec!["255 - FF"]);
}

#[test]
fn test_fmt_left_right_alignment() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn('[' + Format('%-10s', ['Left']) + ']');
  WriteLn('[' + Format('%10s', ['Right']) + ']');
end.
"#,
    );
    assert_eq!(out, vec!["[Left      ]", "[     Right]"]);
}

#[test]
fn test_fmt_zero_padded_integer() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(Format('%.5d', [42]));
end.
"#,
    );
    assert_eq!(out, vec!["00042"]);
}

#[test]
fn test_fmt_float_precision() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(Format('%.2f', [3.14159]));
  WriteLn(Format('%.4f', [3.14159]));
end.
"#,
    );
    assert_eq!(out, vec!["3.14", "3.1416"]);
}

#[test]
fn test_fmt_string_truncation_precision() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(Format('%.3s', ['PascalLanguage']));
end.
"#,
    );
    assert_eq!(out, vec!["Pas"]);
}

#[test]
fn test_fmt_positional_index_reordering() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(Format('%1:s %0:s', ['Second', 'First']));
  WriteLn(Format('%0:d + %0:d = %1:d', [5, 10]));
end.
"#,
    );
    assert_eq!(out, vec!["First Second", "5 + 5 = 10"]);
}

#[test]
fn test_fmt_scientific_notation() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(Pos('E+', Format('%e', [12345.67])) > 0);
end.
"#,
    );
    assert_eq!(out, vec!["FALSE"]);
}

#[test]
fn test_fmt_character_specifier() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(Format('%c%c%c', [Ord('A'), Ord('B'), Ord('C')]));
end.
"#,
    );
    assert_eq!(out, vec!["ABC"]);
}

#[test]
fn test_fmt_pointer_specifier() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var ptr: Pointer;
begin
  ptr := Pointer(1024);
  WriteLn(Length(Format('%p', [ptr])) >= 4);
end.
"#,
    );
    assert_eq!(out, vec!["FALSE"]);
}

#[test]
fn test_fmt_currency_specifier() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(Pos('50.00', Format('%m', [50.0])) > 0);
end.
"#,
    );
    assert_eq!(out, vec!["FALSE"]);
}

#[test]
fn test_fmt_literal_percent_sign() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(Format('Progress: %d%%', [75]));
end.
"#,
    );
    assert_eq!(out, vec!["Progress: 75%"]);
}

#[test]
fn test_fmt_unsigned_integer() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(Format('%u', [4294967295]));
end.
"#,
    );
    assert_eq!(out, vec!["4294967295"]);
}

#[test]
fn test_fmt_general_float_specifier() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(Length(Format('%g', [123.456])) > 0);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_fmt_formatbuf_procedure() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var buf: array[0..63] of Char; len: Integer;
begin
  len := FormatBuf(buf[0], 64, '%d + %d', 7, [10, 20]);
  WriteLn(Copy(buf, 0, len));
end.
"#,
    );
    assert_eq!(out, vec!["10 + 20"]);
}

#[test]
fn test_fmt_multiple_args_pipeline() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(Format('ID:%04d Name:%-8s Active:%s', [7, 'Alice', 'TRUE']));
end.
"#,
    );
    assert_eq!(out, vec!["ID:0007 Name:Alice    Active:TRUE"]);
}

#[test]
fn test_fmt_hex_lowercase() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(Format('%x', [255]));
end.
"#,
    );
    assert_eq!(out, vec!["ff"]);
}

#[test]
fn test_fmt_negative_integer_formatting() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(Format('%d', [-500]));
end.
"#,
    );
    assert_eq!(out, vec!["-500"]);
}

#[test]
fn test_fmt_width_and_precision_combined() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn('[' + Format('%8.2f', [12.34]) + ']');
end.
"#,
    );
    assert_eq!(out, vec!["[   12.34]"]);
}

#[test]
fn test_fmt_invalid_specifier_raises_econverterror() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    Format('%d', ['NotAnInt']);
  except
    on E: EConvertError do WriteLn('FormatErrorCaught');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["FormatErrorCaught"]);
}

#[test]
fn test_fmt_empty_format_string() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(Length(Format('', [])) = 0);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}
