use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 31: ShortString Semantics & Pascal Length Byte (S[0])
// ═══════════════════════════════════════════════════════════

#[test]
fn test_shortstring_length_byte_read() {
    let out = run_pascal(
        r#"
program Test;
var s: ShortString;
begin
  s := 'Hello';
  WriteLn(Ord(s[0]));
  WriteLn(Length(s));
end.
"#,
    );
    assert_eq!(out, vec!["5", "5"]);
}

#[test]
fn test_shortstring_length_byte_mutation() {
    let out = run_pascal(
        r#"
program Test;
var s: ShortString;
begin
  s := 'PascalLanguage';
  s[0] := Chr(6);
  WriteLn(s);
  WriteLn(Length(s));
end.
"#,
    );
    assert_eq!(out, vec!["Pascal", "6"]);
}

#[test]
fn test_shortstring_fixed_size_declaration() {
    let out = run_pascal(
        r#"
program Test;
var s: String[10];
begin
  s := '123456789012345';
  WriteLn(s);
  WriteLn(Length(s));
end.
"#,
    );
    assert_eq!(out, vec!["1234567890", "10"]);
}

#[test]
fn test_shortstring_one_based_indexing() {
    let out = run_pascal(
        r#"
program Test;
var s: ShortString;
begin
  s := 'ABCDE';
  WriteLn(s[1]);
  WriteLn(s[5]);
end.
"#,
    );
    assert_eq!(out, vec!["A", "E"]);
}

#[test]
fn test_shortstring_element_mutation() {
    let out = run_pascal(
        r#"
program Test;
var s: ShortString;
begin
  s := 'Core';
  s[1] := 'M';
  WriteLn(s);
end.
"#,
    );
    assert_eq!(out, vec!["More"]);
}

#[test]
fn test_shortstring_concatenation() {
    let out = run_pascal(
        r#"
program Test;
var s1, s2, s3: ShortString;
begin
  s1 := 'Short'; s2 := 'String';
  s3 := s1 + ' ' + s2;
  WriteLn(s3);
  WriteLn(Length(s3));
end.
"#,
    );
    assert_eq!(out, vec!["Short String", "12"]);
}

#[test]
fn test_shortstring_high_low_bounds() {
    let out = run_pascal(
        r#"
program Test;
var s: ShortString;
begin
  WriteLn(Low(s));
  WriteLn(High(s));
end.
"#,
    );
    assert_eq!(out, vec!["0", "255"]);
}

#[test]
fn test_shortstring_to_ansistring_conversion() {
    let out = run_pascal(
        r#"
program Test;
var ss: ShortString;
    ansi: AnsiString;
begin
  ss := 'ConvertedShortString';
  ansi := ss;
  WriteLn(ansi);
end.
"#,
    );
    assert_eq!(out, vec!["ConvertedShortString"]);
}

#[test]
fn test_ansistring_to_shortstring_truncation() {
    let out = run_pascal(
        r#"
program Test;
var ansi: AnsiString;
    ss: String[5];
begin
  ansi := 'LongAnsiText';
  ss := ansi;
  WriteLn(ss);
  WriteLn(Length(ss));
end.
"#,
    );
    assert_eq!(out, vec!["LongA", "5"]);
}

#[test]
fn test_shortstring_var_parameter_mutation() {
    let out = run_pascal(
        r#"
program Test;
procedure AppendShort(var s: ShortString; suffix: ShortString);
begin
  s := s + suffix;
end;
var text: ShortString;
begin
  text := 'Base';
  AppendShort(text, '_Suffix');
  WriteLn(text);
end.
"#,
    );
    assert_eq!(out, vec!["Base_Suffix"]);
}

#[test]
fn test_shortstring_comparisons() {
    let out = run_pascal(
        r#"
program Test;
var s1, s2: ShortString;
begin
  s1 := 'Apple'; s2 := 'Banana';
  WriteLn(s1 < s2);
  WriteLn(s1 = 'Apple');
  WriteLn(s1 <> s2);
end.
"#,
    );
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_shortstring_empty_initialization() {
    let out = run_pascal(
        r#"
program Test;
var s: ShortString;
begin
  s := '';
  WriteLn(Length(s));
  WriteLn(Ord(s[0]));
end.
"#,
    );
    assert_eq!(out, vec!["0", "0"]);
}

#[test]
fn test_shortstring_fillchar_zero_reset() {
    let out = run_pascal(
        r#"
program Test;
var s: ShortString;
begin
  s := 'Populated';
  FillChar(s, SizeOf(s), 0);
  WriteLn(Length(s));
end.
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_shortstring_record_field() {
    let out = run_pascal(
        r#"
program Test;
type TShortRec = record
  Code: Integer;
  Name: String[20];
end;
var rec: TShortRec;
begin
  rec.Code := 1; rec.Name := 'ShortRecordName';
  WriteLn(rec.Code);
  WriteLn(rec.Name);
end.
"#,
    );
    assert_eq!(out, vec!["1", "ShortRecordName"]);
}

#[test]
fn test_shortstring_array_elements() {
    let out = run_pascal(
        r#"
program Test;
var list: array[1..2] of String[15];
begin
  list[1] := 'ItemOne'; list[2] := 'ItemTwo';
  WriteLn(list[1] + ' & ' + list[2]);
end.
"#,
    );
    assert_eq!(out, vec!["ItemOne & ItemTwo"]);
}

#[test]
fn test_shortstring_case_statement_match() {
    let out = run_pascal(
        r#"
program Test;
var s: ShortString;
begin
  s := 'YES';
  case s of
    'NO': WriteLn('Denied');
    'YES': WriteLn('Approved');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["Approved"]);
}

#[test]
fn test_shortstring_escaped_quote_literal() {
    let out = run_pascal(
        r#"
program Test;
var s: ShortString;
begin
  s := 'It''s Pascal';
  WriteLn(s);
end.
"#,
    );
    assert_eq!(out, vec!["It's Pascal"]);
}

#[test]
fn test_shortstring_pchar_cast() {
    let out = run_pascal(
        r#"
program Test;
var s: ShortString;
    pc: PChar;
begin
  s := 'PCharFromShort' + #0;
  pc := @s[1];
  WriteLn(pc);
end.
"#,
    );
    assert_eq!(out, vec!["PCharFromShort"]);
}

#[test]
fn test_shortstring_setlength_routine() {
    let out = run_pascal(
        r#"
program Test;
var s: ShortString;
begin
  s := '123456789';
  SetLength(s, 4);
  WriteLn(s);
  WriteLn(Ord(s[0]));
end.
"#,
    );
    assert_eq!(out, vec!["1234", "4"]);
}

#[test]
fn test_shortstring_max_capacity_overflow() {
    let out = run_pascal(
        r#"
program Test;
var s1, s2, s3: ShortString;
    i: Integer;
begin
  s1 := ''; s2 := '';
  for i := 1 to 200 do s1 := s1 + 'A';
  for i := 1 to 100 do s2 := s2 + 'B';
  s3 := s1 + s2;
  WriteLn(Length(s3));
end.
"#,
    );
    assert_eq!(out, vec!["255"]);
}
