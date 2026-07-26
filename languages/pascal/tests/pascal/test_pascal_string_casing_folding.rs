use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 37: String Casing & Case Folding (UpperCase, LowerCase, SameText)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_uppercase_basic() {
    let out = run_pascal(
        r#"
program Test;
begin
  WriteLn(UpperCase('hello pascal 123'));
end.
"#,
    );
    assert_eq!(out, vec!["HELLO PASCAL 123"]);
}

#[test]
fn test_lowercase_basic() {
    let out = run_pascal(
        r#"
program Test;
begin
  WriteLn(LowerCase('HELLO PASCAL 123'));
end.
"#,
    );
    assert_eq!(out, vec!["hello pascal 123"]);
}

#[test]
fn test_upcase_char_function() {
    let out = run_pascal(
        r#"
program Test;
var c1, c2: Char;
begin
  c1 := UpCase('a');
  c2 := UpCase('Z');
  WriteLn(c1);
  WriteLn(c2);
end.
"#,
    );
    assert_eq!(out, vec!["A", "Z"]);
}

#[test]
fn test_sametext_case_insensitive_equality() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(SameText('Pascal', 'PASCAL'));
  WriteLn(SameText('Pascal', 'Python'));
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_samestr_case_sensitive_equality() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(SameStr('Pascal', 'Pascal'));
  WriteLn(SameStr('Pascal', 'PASCAL'));
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_ansicomparetext_case_insensitive() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(AnsiCompareText('abc', 'ABC') = 0);
  WriteLn(AnsiCompareText('abc', 'DEF') < 0);
end.
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_ansicomparestr_case_sensitive() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(AnsiCompareStr('abc', 'ABC') <> 0);
  WriteLn(AnsiCompareStr('abc', 'abc') = 0);
end.
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_ansiuppercase_and_ansilowercase() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(AnsiUpperCase('object pascal'));
  WriteLn(AnsiLowerCase('OBJECT PASCAL'));
end.
"#,
    );
    assert_eq!(out, vec!["OBJECT PASCAL", "object pascal"]);
}

#[test]
fn test_casing_on_empty_string() {
    let out = run_pascal(
        r#"
program Test;
begin
  WriteLn('[' + UpperCase('') + ']');
  WriteLn('[' + LowerCase('') + ']');
end.
"#,
    );
    assert_eq!(out, vec!["[]", "[]"]);
}

#[test]
fn test_casing_with_numbers_and_punctuation() {
    let out = run_pascal(
        r#"
program Test;
begin
  WriteLn(UpperCase('123-abc_DEF!'));
  WriteLn(LowerCase('123-abc_DEF!'));
end.
"#,
    );
    assert_eq!(out, vec!["123-ABC_DEF!", "123-abc_def!"]);
}

#[test]
fn test_casing_in_shortstring() {
    let out = run_pascal(
        r#"
program Test;
var ss: ShortString;
begin
  ss := 'ShortText';
  ss := UpperCase(ss);
  WriteLn(ss);
end.
"#,
    );
    assert_eq!(out, vec!["SHORTTEXT"]);
}

#[test]
fn test_casing_in_unicodestring() {
    let out = run_pascal(
        r#"
program Test;
var u: UnicodeString;
begin
  u := 'UnicodeText';
  u := LowerCase(u);
  WriteLn(u);
end.
"#,
    );
    assert_eq!(out, vec!["unicodetext"]);
}

#[test]
fn test_casing_array_of_strings() {
    let out = run_pascal(
        r#"
program Test;
var arr: array[0..1] of String;
begin
  arr[0] := 'first'; arr[1] := 'second';
  arr[0] := UpperCase(arr[0]);
  arr[1] := UpperCase(arr[1]);
  WriteLn(arr[0] + ' ' + arr[1]);
end.
"#,
    );
    assert_eq!(out, vec!["FIRST SECOND"]);
}

#[test]
fn test_casing_record_fields() {
    let out = run_pascal(
        r#"
program Test;
type TUserRec = record
  Username: String;
end;
var u: TUserRec;
begin
  u.Username := 'john_doe';
  u.Username := UpperCase(u.Username);
  WriteLn(u.Username);
end.
"#,
    );
    assert_eq!(out, vec!["JOHN_DOE"]);
}

#[test]
fn test_casing_procedure_parameter_transform() {
    let out = run_pascal(
        r#"
program Test;
procedure NormalizeName(var s: String);
begin
  s := UpperCase(Trim(s));
end;
var name: String;
begin
  name := '   alice   ';
  NormalizeName(name);
  WriteLn(name);
end.
"#,
    );
    assert_eq!(out, vec!["ALICE"]);
}

#[test]
fn test_casing_first_character_capitalization() {
    let out = run_pascal(
        r#"
program Test;
function CapitalizeWord(s: String): String;
begin
  if Length(s) = 0 then Result := ''
  else
  begin
    Result := LowerCase(s);
    Result[1] := UpCase(Result[1]);
  end;
end;
begin
  WriteLn(CapitalizeWord('pASCAL'));
end.
"#,
    );
    assert_eq!(out, vec!["Pascal"]);
}

#[test]
fn test_sametext_in_conditional_branch() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
procedure HandleCommand(cmd: String);
begin
  if SameText(cmd, 'START') then WriteLn('CommandStart')
  else if SameText(cmd, 'STOP') then WriteLn('CommandStop');
end;
begin
  HandleCommand('Start');
  HandleCommand('stop');
end.
"#,
    );
    assert_eq!(out, vec!["CommandStart", "CommandStop"]);
}

#[test]
fn test_upcase_character_loop() {
    let out = run_pascal(
        r#"
program Test;
var s: String; i: Integer;
begin
  s := 'hello';
  for i := 1 to Length(s) do
    s[i] := UpCase(s[i]);
  WriteLn(s);
end.
"#,
    );
    assert_eq!(out, vec!["HELLO"]);
}

#[test]
fn test_casing_in_class_property_setter() {
    let out = run_pascal(
        r#"
program Test;
type TUpperHolder = class
  private FTitle: String;
  private procedure SetTitle(t: String);
  public property Title: String read FTitle write SetTitle;
end;
procedure TUpperHolder.SetTitle(t: String); begin FTitle := UpperCase(t); end;
var h: TUpperHolder;
begin
  h := TUpperHolder.Create;
  h.Title := 'lowercase title';
  WriteLn(h.Title);
  h.Free;
end.
"#,
    );
    assert_eq!(out, vec!["LOWERCASE TITLE"]);
}

#[test]
fn test_samestr_exact_case_matching() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(SameStr('admin', 'ADMIN'));
  WriteLn(SameStr('admin', 'admin'));
end.
"#,
    );
    assert_eq!(out, vec!["False", "True"]);
}
