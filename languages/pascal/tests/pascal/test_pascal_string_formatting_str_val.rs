use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 35: String Formatting, Str, Val & Conversion Functions
// ═══════════════════════════════════════════════════════════

#[test]
fn test_str_procedure_integer_formatting() {
    let out = run_pascal(r#"
program Test;
var s: String;
    num: Integer;
begin
  num := 42;
  Str(num, s);
  WriteLn(s);
end.
"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_str_procedure_width_specifier() {
    let out = run_pascal(r#"
program Test;
var s: String;
begin
  Str(99:5, s);
  WriteLn('[' + s + ']');
end.
"#);
    assert_eq!(out, vec!["[   99]"]);
}

#[test]
fn test_str_procedure_float_precision() {
    let out = run_pascal(r#"
program Test;
var s: String;
begin
  Str(3.14159:0:2, s);
  WriteLn(s);
end.
"#);
    assert_eq!(out, vec!["3.14"]);
}

#[test]
fn test_val_procedure_integer_valid() {
    let out = run_pascal(r#"
program Test;
var val, code: Integer;
begin
  Val('1234', val, code);
  WriteLn(val);
  WriteLn(code);
end.
"#);
    assert_eq!(out, vec!["1234", "0"]);
}

#[test]
fn test_val_procedure_integer_invalid_code_position() {
    let out = run_pascal(r#"
program Test;
var val, code: Integer;
begin
  Val('12a4', val, code);
  WriteLn(code);
end.
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_val_procedure_float_conversion() {
    let out = run_pascal(r#"
program Test;
var r: Real; code: Integer;
begin
  Val('45.67', r, code);
  WriteLn(r);
  WriteLn(code);
end.
"#);
    assert_eq!(out, vec!["45.67", "0"]);
}

#[test]
fn test_format_function_basic_d_s_f() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var s: String;
begin
  s := Format('Item %s ID %d Price %.2f', ['Book', 10, 19.95]);
  WriteLn(s);
end.
"#);
    assert_eq!(out, vec!["Item Book ID 10 Price 19.95"]);
}

#[test]
fn test_format_hex_specifier() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  WriteLn(Format('Hex: %.4x', [255]));
end.
"#);
    assert_eq!(out, vec!["Hex: 00FF"]);
}

#[test]
fn test_inttohex_function() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  WriteLn(IntToHex(255, 4));
  WriteLn(IntToHex(16, 2));
end.
"#);
    assert_eq!(out, vec!["00FF", "10"]);
}

#[test]
fn test_strtointdef_fallback() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  WriteLn(StrToIntDef('500', -1));
  WriteLn(StrToIntDef('invalid', -1));
end.
"#);
    assert_eq!(out, vec!["500", "-1"]);
}

#[test]
fn test_booltostr_and_strtobool() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  WriteLn(BoolToStr(True, True));
  WriteLn(StrToBool('True'));
end.
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_floattostrf_fixed_format() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  WriteLn(FloatToStrF(12.3456, ffFixed, 8, 2));
end.
"#);
    assert_eq!(out, vec!["12.35"]);
}

#[test]
fn test_formatfloat_pattern() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  WriteLn(FormatFloat('000.00', 5.2));
end.
"#);
    assert_eq!(out, vec!["005.20"]);
}

#[test]
fn test_format_width_padding() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  WriteLn('[' + Format('%10s', ['Right']) + ']');
  WriteLn('[' + Format('%-10s', ['Left']) + ']');
end.
"#);
    assert_eq!(out, vec!["[     Right]", "[Left      ]"]);
}

#[test]
fn test_format_escaped_percent() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  WriteLn(Format('Discount: %d%%', [20]));
end.
"#);
    assert_eq!(out, vec!["Discount: 20%"]);
}

#[test]
fn test_str_negative_numbers() {
    let out = run_pascal(r#"
program Test;
var s: String;
begin
  Str(-100, s);
  WriteLn(s);
end.
"#);
    assert_eq!(out, vec!["-100"]);
}

#[test]
fn test_val_negative_numbers() {
    let out = run_pascal(r#"
program Test;
var val, code: Integer;
begin
  Val('-456', val, code);
  WriteLn(val);
  WriteLn(code);
end.
"#);
    assert_eq!(out, vec!["-456", "0"]);
}

#[test]
fn test_strtofloatdef_fallback() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  WriteLn(StrToFloatDef('12.5', 0.0));
  WriteLn(StrToFloatDef('bad', -1.0));
end.
"#);
    assert_eq!(out, vec!["12.5", "-1"]);
}

#[test]
fn test_format_multiple_arguments_array() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  WriteLn(Format('%d + %d = %d', [5, 10, 15]));
end.
"#);
    assert_eq!(out, vec!["5 + 10 = 15"]);
}

#[test]
fn test_inttostr_basic() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  WriteLn(IntToStr(789));
end.
"#);
    assert_eq!(out, vec!["789"]);
}
