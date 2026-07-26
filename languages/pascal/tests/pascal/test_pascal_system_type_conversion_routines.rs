use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 77: System Type Conversion Routines & Casts
// ═══════════════════════════════════════════════════════════

#[test]
fn test_conv_inttostr_strtoint() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var s: String; i: Integer;
begin
  s := IntToStr(12345);
  i := StrToInt(s);
  WriteLn(i);
end.
"#,
    );
    assert_eq!(out, vec!["12345"]);
}

#[test]
fn test_conv_strtointdef_fallback() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(StrToIntDef('99', -1));
  WriteLn(StrToIntDef('Invalid', -1));
end.
"#,
    );
    assert_eq!(out, vec!["99", "-1"]);
}

#[test]
fn test_conv_floattostr_strtofloat() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var s: String; f: Double;
begin
  s := FloatToStr(12.34);
  f := StrToFloat(s);
  WriteLn(f);
end.
"#,
    );
    assert_eq!(out, vec!["12.34"]);
}

#[test]
fn test_conv_strtofloatdef_fallback() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(StrToFloatDef('3.14', 0.0) = 3.14);
  WriteLn(StrToFloatDef('NotAFloat', 0.0) = 0.0);
end.
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_conv_booltostr_strtobool() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var s: String; b: Boolean;
begin
  s := BoolToStr(True, True); // 'True'
  b := StrToBool(s);
  WriteLn(b);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_conv_strtobooldef_fallback() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(StrToBoolDef('True', False));
  WriteLn(StrToBoolDef('Invalid', False));
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_conv_inttohex() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(IntToHex(255, 2));
  WriteLn(IntToHex(255, 4));
  WriteLn(IntToHex(4096, 4));
end.
"#,
    );
    assert_eq!(out, vec!["FF", "00FF", "1000"]);
}

#[test]
fn test_conv_ord_chr() {
    let out = run_pascal(
        r#"
program Test;
var ch: Char; code: Integer;
begin
  code := Ord('A');
  ch := Chr(code);
  WriteLn(code.ToString + ':' + ch);
end.
"#,
    );
    assert_eq!(out, vec!["65:A"]);
}

#[test]
fn test_conv_move_procedure() {
    let out = run_pascal(
        r#"
program Test;
var src, dst: Integer;
begin
  src := 123456789;
  dst := 0;
  Move(src, dst, SizeOf(Integer));
  WriteLn(dst);
end.
"#,
    );
    assert_eq!(out, vec!["123456789"]);
}

#[test]
fn test_conv_pointer_nativeint_cast() {
    let out = run_pascal(
        r#"
program Test;
var p: Pointer; n: NativeInt;
begin
  n := 4096;
  p := Pointer(n);
  WriteLn(NativeInt(p));
end.
"#,
    );
    assert_eq!(out, vec!["4096"]);
}

#[test]
fn test_conv_pchar_cast() {
    let out = run_pascal(
        r#"
program Test;
var s: String; pc: PChar;
begin
  s := 'PCharTest';
  pc := PChar(s);
  WriteLn(String(pc));
end.
"#,
    );
    assert_eq!(out, vec!["PCharTest"]);
}

#[test]
fn test_conv_trunc_round() {
    let out = run_pascal(
        r#"
program Test;
begin
  WriteLn(Trunc(3.9));
  WriteLn(Round(3.9));
  WriteLn(Round(3.1));
end.
"#,
    );
    assert_eq!(out, vec!["3", "4", "3"]);
}

#[test]
fn test_conv_val_procedure() {
    let out = run_pascal(
        r#"
program Test;
var num, code: Integer;
begin
  Val('500', num, code);
  WriteLn(num.ToString + ':' + code.ToString);
  Val('50x', num, code);
  WriteLn(code <> 0);
end.
"#,
    );
    assert_eq!(out, vec!["500:0", "True"]);
}

#[test]
fn test_conv_str_procedure() {
    let out = run_pascal(
        r#"
program Test;
var val: Integer; s: String;
begin
  val := 777;
  Str(val, s);
  WriteLn(s);
end.
"#,
    );
    assert_eq!(out, vec!["777"]);
}

#[test]
fn test_conv_enum_to_int_cast() {
    let out = run_pascal(
        r#"
program Test;
type TLevel = (lvlLow, lvlMed, lvlHigh);
var l: TLevel; i: Integer;
begin
  l := lvlMed;
  i := Ord(l);
  WriteLn(i);
  l := TLevel(2);
  WriteLn(Ord(l));
end.
"#,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn test_conv_int64tostr_strtoint64() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var v64: Int64; s: String;
begin
  v64 := 9000000000000;
  s := IntToStr(v64);
  WriteLn(StrToInt64(s));
end.
"#,
    );
    assert_eq!(out, vec!["9000000000000"]);
}

#[test]
fn test_conv_bin_to_hex_str() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var b: Byte;
begin
  b := 15;
  WriteLn(IntToHex(b, 2));
end.
"#,
    );
    assert_eq!(out, vec!["0F"]);
}

#[test]
fn test_conv_currencytostr() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var c: Currency;
begin
  c := 49.99;
  WriteLn(CurrToStr(c));
end.
"#,
    );
    assert_eq!(out, vec!["49.99"]);
}

#[test]
fn test_conv_strtocurrdef() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(StrToCurrDef('49.99', 0.0) = 49.99);
  WriteLn(StrToCurrDef('invalid', 0.0) = 0.0);
end.
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_conv_record_to_byte_array_move() {
    let out = run_pascal(
        r#"
program Test;
type TRec = packed record X, Y: Word; end;
var r1, r2: TRec; bytes: array[0..3] of Byte;
begin
  r1.X := 10; r1.Y := 20;
  Move(r1, bytes[0], SizeOf(TRec));
  Move(bytes[0], r2, SizeOf(TRec));
  WriteLn(r2.X.ToString + ':' + r2.Y.ToString);
end.
"#,
    );
    assert_eq!(out, vec!["10:20"]);
}
