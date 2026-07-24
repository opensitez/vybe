use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 34: UTF8String, RawByteString & CodePage Conversions
// ═══════════════════════════════════════════════════════════

#[test]
fn test_utf8string_declaration() {
    let out = run_pascal(r#"
program Test;
var u: UTF8String;
begin
  u := 'UTF8Text';
  WriteLn(u);
  WriteLn(Length(u));
end.
"#);
    assert_eq!(out, vec!["UTF8Text", "8"]);
}

#[test]
fn test_rawbytestring_declaration() {
    let out = run_pascal(r#"
program Test;
var r: RawByteString;
begin
  r := 'RawByteData';
  WriteLn(r);
  WriteLn(Length(r));
end.
"#);
    assert_eq!(out, vec!["RawByteData", "11"]);
}

#[test]
fn test_cp_utf8_constant_value() {
    let out = run_pascal(r#"
program Test;
begin
  WriteLn(CP_UTF8);
end.
"#);
    assert_eq!(out, vec!["65001"]);
}

#[test]
fn test_utf8string_codepage_query() {
    let out = run_pascal(r#"
program Test;
var u: UTF8String;
begin
  u := 'CheckCodePage';
  WriteLn(StringCodePage(u));
end.
"#);
    assert_eq!(out, vec!["65001"]);
}

#[test]
fn test_utf8encode_utf8decode_routines() {
    let out = run_pascal(r#"
program Test;
var uni: UnicodeString;
    utf8: UTF8String;
    decoded: UnicodeString;
begin
  uni := 'UnicodeText';
  utf8 := UTF8Encode(uni);
  decoded := UTF8Decode(utf8);
  WriteLn(decoded);
end.
"#);
    assert_eq!(out, vec!["UnicodeText"]);
}

#[test]
fn test_setcodepage_no_convert() {
    let out = run_pascal(r#"
program Test;
var r: RawByteString;
begin
  r := 'RawData';
  SetCodePage(r, CP_UTF8, False);
  WriteLn(StringCodePage(r));
end.
"#);
    assert_eq!(out, vec!["65001"]);
}

#[test]
fn test_utf8string_to_unicodestring_implicit_conversion() {
    let out = run_pascal(r#"
program Test;
var utf8: UTF8String;
    uni: UnicodeString;
begin
  utf8 := 'UTF8ToUnicode';
  uni := utf8;
  WriteLn(uni);
end.
"#);
    assert_eq!(out, vec!["UTF8ToUnicode"]);
}

#[test]
fn test_unicodestring_to_utf8string_implicit_conversion() {
    let out = run_pascal(r#"
program Test;
var uni: UnicodeString;
    utf8: UTF8String;
begin
  uni := 'UnicodeToUTF8';
  utf8 := uni;
  WriteLn(utf8);
end.
"#);
    assert_eq!(out, vec!["UnicodeToUTF8"]);
}

#[test]
fn test_putf8char_pointer_casting() {
    let out = run_pascal(r#"
program Test;
var u: UTF8String;
    p: PUTF8Char;
begin
  u := 'PUTF8Test';
  p := PUTF8Char(u);
  WriteLn(p^);
  WriteLn(p);
end.
"#);
    assert_eq!(out, vec!["P", "PUTF8Test"]);
}

#[test]
fn test_rawbytestring_prevents_implicit_conversion() {
    let out = run_pascal(r#"
program Test;
procedure ProcessRaw(const r: RawByteString);
begin
  WriteLn(r);
end;
var s: UTF8String;
begin
  s := 'RawPassThrough';
  ProcessRaw(s);
end.
"#);
    assert_eq!(out, vec!["RawPassThrough"]);
}

#[test]
fn test_utf8string_concatenation() {
    let out = run_pascal(r#"
program Test;
var u1, u2, u3: UTF8String;
begin
  u1 := 'UTF8_'; u2 := 'Concat';
  u3 := u1 + u2;
  WriteLn(u3);
end.
"#);
    assert_eq!(out, vec!["UTF8_Concat"]);
}

#[test]
fn test_utf8string_comparisons() {
    let out = run_pascal(r#"
program Test;
var u1, u2: UTF8String;
begin
  u1 := 'Alpha'; u2 := 'Beta';
  WriteLn(u1 < u2);
  WriteLn(u1 = 'Alpha');
  WriteLn(u1 <> u2);
end.
"#);
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_utf8string_record_field() {
    let out = run_pascal(r#"
program Test;
type TUTF8Rec = record
  ID: Integer;
  Data: UTF8String;
end;
var rec: TUTF8Rec;
begin
  rec.ID := 10; rec.Data := 'RecordUTF8';
  WriteLn(rec.ID);
  WriteLn(rec.Data);
end.
"#);
    assert_eq!(out, vec!["10", "RecordUTF8"]);
}

#[test]
fn test_utf8string_class_property() {
    let out = run_pascal(r#"
program Test;
type TUTF8Holder = class
  private FData: UTF8String;
  public property Data: UTF8String read FData write FData;
end;
var h: TUTF8Holder;
begin
  h := TUTF8Holder.Create;
  h.Data := 'ClassUTF8Property';
  WriteLn(h.Data);
  h.Free;
end.
"#);
    assert_eq!(out, vec!["ClassUTF8Property"]);
}

#[test]
fn test_utf8string_var_parameter() {
    let out = run_pascal(r#"
program Test;
procedure AppendUTF8(var u: UTF8String; suffix: UTF8String);
begin
  u := u + suffix;
end;
var text: UTF8String;
begin
  text := 'Base';
  AppendUTF8(text, '_UTF8');
  WriteLn(text);
end.
"#);
    assert_eq!(out, vec!["Base_UTF8"]);
}

#[test]
fn test_utf8string_function_return() {
    let out = run_pascal(r#"
program Test;
function GetUTF8Text: UTF8String;
begin
  Result := 'FunctionUTF8Return';
end;
begin
  WriteLn(GetUTF8Text);
end.
"#);
    assert_eq!(out, vec!["FunctionUTF8Return"]);
}

#[test]
fn test_utf8string_setlength() {
    let out = run_pascal(r#"
program Test;
var u: UTF8String;
begin
  u := 'TruncateUTF8Text';
  SetLength(u, 8);
  WriteLn(u);
end.
"#);
    assert_eq!(out, vec!["Truncate"]);
}

#[test]
fn test_rawbytestring_move_to_byte_array() {
    let out = run_pascal(r#"
program Test;
var r: RawByteString;
    bytes: array[0..3] of Byte;
begin
  r := 'ABCD';
  Move(r[1], bytes[0], 4);
  WriteLn(bytes[0]);
  WriteLn(bytes[3]);
end.
"#);
    assert_eq!(out, vec!["65", "68"]);
}

#[test]
fn test_utf8string_empty_check() {
    let out = run_pascal(r#"
program Test;
var u: UTF8String;
begin
  u := '';
  WriteLn(Length(u));
  WriteLn(u = '');
end.
"#);
    assert_eq!(out, vec!["0", "True"]);
}

#[test]
fn test_utf8string_high_low_bounds() {
    let out = run_pascal(r#"
program Test;
var u: UTF8String;
begin
  u := 'Sample';
  WriteLn(Low(u));
  WriteLn(High(u));
end.
"#);
    assert_eq!(out, vec!["1", "6"]);
}
