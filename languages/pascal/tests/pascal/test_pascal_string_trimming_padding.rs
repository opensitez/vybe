use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 38: String Trimming, Padding & Alignment Routines
// ═══════════════════════════════════════════════════════════

#[test]
fn test_trim_spaces() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn('[' + Trim('   Hello World   ') + ']');
end.
"#,
    );
    assert_eq!(out, vec!["[Hello World]"]);
}

#[test]
fn test_trimleft_spaces() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn('[' + TrimLeft('   Hello World   ') + ']');
end.
"#,
    );
    assert_eq!(out, vec!["[Hello World   ]"]);
}

#[test]
fn test_trimright_spaces() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn('[' + TrimRight('   Hello World   ') + ']');
end.
"#,
    );
    assert_eq!(out, vec!["[   Hello World]"]);
}

#[test]
fn test_trim_whitespace_control_characters() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var s: String;
begin
  s := #9 + #10 + #13 + ' Data ' + #13 + #10 + #9;
  WriteLn('[' + Trim(s) + ']');
end.
"#,
    );
    assert_eq!(out, vec!["[Data]"]);
}

#[test]
fn test_trim_empty_and_spaces_only() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn('[' + Trim('') + ']');
  WriteLn('[' + Trim('     ') + ']');
end.
"#,
    );
    assert_eq!(out, vec!["[]", "[]"]);
}

#[test]
fn test_padleft_custom_character() {
    let out = run_pascal(
        r#"
program Test;
uses StrUtils;
begin
  WriteLn(PadLeft('42', 5, '0'));
end.
"#,
    );
    assert_eq!(out, vec!["00042"]);
}

#[test]
fn test_padright_custom_character() {
    let out = run_pascal(
        r#"
program Test;
uses StrUtils;
begin
  WriteLn(PadRight('Item', 8, '*'));
end.
"#,
    );
    assert_eq!(out, vec!["Item****"]);
}

#[test]
fn test_dupestring_multiplier() {
    let out = run_pascal(
        r#"
program Test;
uses StrUtils;
begin
  WriteLn(DupeString('ABC', 3));
end.
"#,
    );
    assert_eq!(out, vec!["ABCABCABC"]);
}

#[test]
fn test_reversestring_function() {
    let out = run_pascal(
        r#"
program Test;
uses StrUtils;
begin
  WriteLn(ReverseString('Pascal'));
end.
"#,
    );
    assert_eq!(out, vec!["lacsaP"]);
}

#[test]
fn test_padding_string_already_longer_than_length() {
    let out = run_pascal(
        r#"
program Test;
uses StrUtils;
begin
  WriteLn(PadLeft('LongerText', 5, '0'));
end.
"#,
    );
    assert_eq!(out, vec!["LongerText"]);
}

#[test]
fn test_dupestring_zero_count() {
    let out = run_pascal(
        r#"
program Test;
uses StrUtils;
begin
  WriteLn('[' + DupeString('Text', 0) + ']');
end.
"#,
    );
    assert_eq!(out, vec!["[]"]);
}

#[test]
fn test_reversestring_palindromic_check() {
    let out = run_pascal(
        r#"
program Test;
uses StrUtils;
var s: String;
begin
  s := 'radar';
  WriteLn(s = ReverseString(s));
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_trim_shortstring() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var ss: ShortString;
begin
  ss := '   ShortData   ';
  ss := Trim(ss);
  WriteLn(ss);
end.
"#,
    );
    assert_eq!(out, vec!["ShortData"]);
}

#[test]
fn test_trim_unicodestring() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var u: UnicodeString;
begin
  u := '   UnicodeData   ';
  u := Trim(u);
  WriteLn(u);
end.
"#,
    );
    assert_eq!(out, vec!["UnicodeData"]);
}

#[test]
fn test_trim_before_integer_conversion() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var numStr: String;
begin
  numStr := '   777   ';
  WriteLn(StrToInt(Trim(numStr)));
end.
"#,
    );
    assert_eq!(out, vec!["777"]);
}

#[test]
fn test_chained_trim_and_pad() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils, StrUtils;
var raw: String;
begin
  raw := '   Code   ';
  WriteLn(PadRight(Trim(raw), 10, '-'));
end.
"#,
    );
    assert_eq!(out, vec!["Code------"]);
}

#[test]
fn test_trim_array_elements() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var arr: array[0..1] of String;
begin
  arr[0] := '  A  '; arr[1] := '  B  ';
  arr[0] := Trim(arr[0]); arr[1] := Trim(arr[1]);
  WriteLn(arr[0] + arr[1]);
end.
"#,
    );
    assert_eq!(out, vec!["AB"]);
}

#[test]
fn test_padleft_default_space_character() {
    let out = run_pascal(
        r#"
program Test;
uses StrUtils;
begin
  WriteLn('[' + PadLeft('Val', 6) + ']');
end.
"#,
    );
    assert_eq!(out, vec!["[   Val]"]);
}

#[test]
fn test_padright_default_space_character() {
    let out = run_pascal(
        r#"
program Test;
uses StrUtils;
begin
  WriteLn('[' + PadRight('Val', 6) + ']');
end.
"#,
    );
    assert_eq!(out, vec!["[Val   ]"]);
}

#[test]
fn test_trim_inside_property_setter() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type TCleanItem = class
  private FName: String;
  private procedure SetName(v: String);
  public property Name: String read FName write SetName;
end;
procedure TCleanItem.SetName(v: String); begin FName := Trim(v); end;
var item: TCleanItem;
begin
  item := TCleanItem.Create;
  item.Name := '   Padded   ';
  WriteLn(item.Name);
  item.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Padded"]);
}
