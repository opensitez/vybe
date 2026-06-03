/// Tests for type conversion functions in Pascal/Delphi:
/// IntToHex, HexToInt patterns, BoolToStr, StrToBool,
/// Val procedure, numeric string validation, conversion edge cases.
use super::helpers::run_pascal;

// ===================================================================
// INTTOHEX
// ===================================================================

#[test]
fn inttohex_basic() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(IntToHex(255, 2));
end."#
        ),
        &["FF"]
    );
}

#[test]
fn inttohex_zero() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(IntToHex(0, 4));
end."#
        ),
        &["0000"]
    );
}

#[test]
fn inttohex_with_padding() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(IntToHex(16, 4));
end."#
        ),
        &["0010"]
    );
}

#[test]
fn inttohex_large() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(IntToHex(256, 3));
end."#
        ),
        &["100"]
    );
}

// ===================================================================
// BOOLTOSTR
// ===================================================================

#[test]
fn booltostr_true() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(BoolToStr(True, True));
end."#
        ),
        &["True"]
    );
}

#[test]
fn booltostr_false() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(BoolToStr(False, True));
end."#
        ),
        &["False"]
    );
}

// ===================================================================
// STRTOBOOL
// ===================================================================

#[test]
fn strtobool_true_string() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(StrToBool('true'));
end."#
        ),
        &["true"]
    );
}

#[test]
fn strtobool_false_string() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(StrToBool('false'));
end."#
        ),
        &["false"]
    );
}

// ===================================================================
// STRTOINTDEF (with default on failure)
// ===================================================================

#[test]
fn strtointdef_valid() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(StrToIntDef('42', 0));
end."#
        ),
        &["42"]
    );
}

#[test]
fn strtointdef_invalid() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(StrToIntDef('abc', -1));
end."#
        ),
        &["-1"]
    );
}

// ===================================================================
// STRTOFLOAT / FLOATTOSTR PATTERNS
// ===================================================================

#[test]
fn floattostr_integer_value() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(FloatToStr(100.0));
end."#
        ),
        &["100"]
    );
}

#[test]
fn floattostr_decimal() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(FloatToStr(3.5));
end."#
        ),
        &["3.5"]
    );
}

#[test]
fn strtofloat_then_calc() {
    assert_eq!(
        run_pascal(
            r#"program T;
var f: Real;
begin
  f := StrToFloat('2.5');
  WriteLn(f * 2.0);
end."#
        ),
        &["5"]
    );
}

// ===================================================================
// INTTOSTR / STRTOINT PATTERNS
// ===================================================================

#[test]
fn inttostr_negative() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(IntToStr(-42));
end."#
        ),
        &["-42"]
    );
}

#[test]
fn strtoint_then_math() {
    assert_eq!(
        run_pascal(
            r#"program T;
var n: Integer;
begin
  n := StrToInt('100');
  WriteLn(n * n);
end."#
        ),
        &["10000"]
    );
}

// ===================================================================
// NUMERIC STRING VALIDATION PATTERNS
// ===================================================================

#[test]
fn is_numeric_string() {
    assert_eq!(
        run_pascal(
            r#"program T;
function IsNumeric(s: String): Boolean;
var i: Integer;
begin
  Result := Length(s) > 0;
  for i := 1 to Length(s) do
    if (Ord(s[i]) < Ord('0')) or (Ord(s[i]) > Ord('9')) then
      Result := False;
end;
begin
  WriteLn(IsNumeric('123'));
  WriteLn(IsNumeric('12x'));
end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn string_to_int_sum() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a, b, c: Integer;
begin
  a := StrToInt('10');
  b := StrToInt('20');
  c := a + b;
  WriteLn(IntToStr(c));
end."#
        ),
        &["30"]
    );
}

// ===================================================================
// EXPLICIT TYPE CASTS IN EXPRESSIONS
// ===================================================================

#[test]
fn integer_cast_in_expr() {
    assert_eq!(
        run_pascal(
            r#"program T;
var r: Real;
    i: Integer;
begin
  r := 7.9;
  i := Trunc(r);
  WriteLn(i);
end."#
        ),
        &["7"]
    );
}

#[test]
fn round_then_inttostr() {
    assert_eq!(
        run_pascal(
            r#"program T;
var r: Real;
begin
  r := 3.6;
  WriteLn(IntToStr(Round(r)));
end."#
        ),
        &["4"]
    );
}

#[test]
fn format_with_conversion() {
    assert_eq!(
        run_pascal(
            r#"program T;
var n: Integer;
begin
  n := 42;
  WriteLn(Format('Hex: %s', [IntToHex(n, 4)]));
end."#
        ),
        &["Hex: 002A"]
    );
}
