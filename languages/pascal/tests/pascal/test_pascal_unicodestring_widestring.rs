use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 33: UnicodeString, WideString & UTF-16 Encoding
// ═══════════════════════════════════════════════════════════

#[test]
fn test_unicodestring_basic_declaration() {
    let out = run_pascal(
        r#"
program Test;
var u: UnicodeString;
begin
  u := 'UnicodeText';
  WriteLn(u);
  WriteLn(Length(u));
end.
"#,
    );
    assert_eq!(out, vec!["UnicodeText", "11"]);
}

#[test]
fn test_widestring_basic_declaration() {
    let out = run_pascal(
        r#"
program Test;
var w: WideString;
begin
  w := 'WideStringData';
  WriteLn(w);
  WriteLn(Length(w));
end.
"#,
    );
    assert_eq!(out, vec!["WideStringData", "14"]);
}

#[test]
fn test_widechar_sizeof() {
    let out = run_pascal(
        r#"
program Test;
begin
  WriteLn(SizeOf(WideChar));
end.
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_unicodestring_utf16_codepoints() {
    let out = run_pascal(
        r#"
program Test;
var u: UnicodeString;
begin
  u := 'A';
  WriteLn(Ord(u[1]));
end.
"#,
    );
    assert_eq!(out, vec!["65"]);
}

#[test]
fn test_unicodestring_concatenation() {
    let out = run_pascal(
        r#"
program Test;
var u1, u2, u3: UnicodeString;
begin
  u1 := 'Hello '; u2 := 'Unicode';
  u3 := u1 + u2;
  WriteLn(u3);
end.
"#,
    );
    assert_eq!(out, vec!["Hello Unicode"]);
}

#[test]
fn test_unicodestring_to_ansistring_conversion() {
    let out = run_pascal(
        r#"
program Test;
var u: UnicodeString;
    a: AnsiString;
begin
  u := 'ConvertMe';
  a := AnsiString(u);
  WriteLn(a);
end.
"#,
    );
    assert_eq!(out, vec!["ConvertMe"]);
}

#[test]
fn test_ansistring_to_unicodestring_conversion() {
    let out = run_pascal(
        r#"
program Test;
var a: AnsiString;
    u: UnicodeString;
begin
  a := 'AnsiToUnicode';
  u := UnicodeString(a);
  WriteLn(u);
end.
"#,
    );
    assert_eq!(out, vec!["AnsiToUnicode"]);
}

#[test]
fn test_widestring_to_unicodestring_conversion() {
    let out = run_pascal(
        r#"
program Test;
var w: WideString;
    u: UnicodeString;
begin
  w := 'WideToUnicode';
  u := w;
  WriteLn(u);
end.
"#,
    );
    assert_eq!(out, vec!["WideToUnicode"]);
}

#[test]
fn test_pwidechar_pointer_casting() {
    let out = run_pascal(
        r#"
program Test;
var u: UnicodeString;
    pw: PWideChar;
begin
  u := 'PWideCharTest';
  pw := PWideChar(u);
  WriteLn(pw^);
  WriteLn(pw);
end.
"#,
    );
    assert_eq!(out, vec!["P", "PWideCharTest"]);
}

#[test]
fn test_unicodestring_indexing_and_mutation() {
    let out = run_pascal(
        r#"
program Test;
var u: UnicodeString;
begin
  u := 'Core';
  u[1] := 'M';
  WriteLn(u);
end.
"#,
    );
    assert_eq!(out, vec!["More"]);
}

#[test]
fn test_unicodestring_setlength_expansion() {
    let out = run_pascal(
        r#"
program Test;
var u: UnicodeString;
begin
  u := 'Base';
  SetLength(u, 8);
  WriteLn(Length(u));
end.
"#,
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn test_unicodestring_setlength_truncation() {
    let out = run_pascal(
        r#"
program Test;
var u: UnicodeString;
begin
  u := 'TruncateUnicode';
  SetLength(u, 8);
  WriteLn(u);
end.
"#,
    );
    assert_eq!(out, vec!["Truncate"]);
}

#[test]
fn test_unicodestring_comparisons() {
    let out = run_pascal(
        r#"
program Test;
var u1, u2: UnicodeString;
begin
  u1 := 'Apple'; u2 := 'Banana';
  WriteLn(u1 < u2);
  WriteLn(u1 = 'Apple');
  WriteLn(u1 <> u2);
end.
"#,
    );
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_unicodestring_var_parameter() {
    let out = run_pascal(
        r#"
program Test;
procedure AppendUnicode(var u: UnicodeString; suffix: UnicodeString);
begin
  u := u + suffix;
end;
var text: UnicodeString;
begin
  text := 'Prefix';
  AppendUnicode(text, '_Suffix');
  WriteLn(text);
end.
"#,
    );
    assert_eq!(out, vec!["Prefix_Suffix"]);
}

#[test]
fn test_unicodestring_const_parameter() {
    let out = run_pascal(
        r#"
program Test;
function FormatUnicode(const u: UnicodeString): UnicodeString;
begin
  Result := '[U]: ' + u;
end;
begin
  WriteLn(FormatUnicode('Val'));
end.
"#,
    );
    assert_eq!(out, vec!["[U]: Val"]);
}

#[test]
fn test_unicodestring_record_field() {
    let out = run_pascal(
        r#"
program Test;
type TUniRec = record
  ID: Integer;
  Text: UnicodeString;
end;
var rec: TUniRec;
begin
  rec.ID := 100; rec.Text := 'RecordUnicode';
  WriteLn(rec.ID);
  WriteLn(rec.Text);
end.
"#,
    );
    assert_eq!(out, vec!["100", "RecordUnicode"]);
}

#[test]
fn test_unicodestring_class_property() {
    let out = run_pascal(
        r#"
program Test;
type TUniHolder = class
  private FTitle: UnicodeString;
  public property Title: UnicodeString read FTitle write FTitle;
end;
var h: TUniHolder;
begin
  h := TUniHolder.Create;
  h.Title := 'ClassUnicodeProperty';
  WriteLn(h.Title);
  h.Free;
end.
"#,
    );
    assert_eq!(out, vec!["ClassUnicodeProperty"]);
}

#[test]
fn test_unicodestring_function_return() {
    let out = run_pascal(
        r#"
program Test;
function GetUniText: UnicodeString;
begin
  Result := 'FunctionUnicodeReturn';
end;
begin
  WriteLn(GetUniText);
end.
"#,
    );
    assert_eq!(out, vec!["FunctionUnicodeReturn"]);
}

#[test]
fn test_unicodestring_empty_string_check() {
    let out = run_pascal(
        r#"
program Test;
var u: UnicodeString;
begin
  u := '';
  WriteLn(Length(u));
  WriteLn(u = '');
end.
"#,
    );
    assert_eq!(out, vec!["0", "True"]);
}

#[test]
fn test_unicodestring_high_low_bounds() {
    let out = run_pascal(
        r#"
program Test;
var u: UnicodeString;
begin
  u := 'Bounds';
  WriteLn(Low(u));
  WriteLn(High(u));
end.
"#,
    );
    assert_eq!(out, vec!["1", "6"]);
}
