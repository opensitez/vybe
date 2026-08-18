use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 32: AnsiString Reference Counting & Copy-on-Write
// ═══════════════════════════════════════════════════════════

#[test]
fn test_ansistring_refcount_query() {
    let out = run_pascal(
        r#"
program Test;
var s1, s2: AnsiString;
begin
  s1 := 'SharedAnsiStringData';
  s2 := s1;
  WriteLn(StringRefCount(s1) > 1);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_ansistring_copy_on_write_mutation() {
    let out = run_pascal(
        r#"
program Test;
var s1, s2: AnsiString;
begin
  s1 := 'Hello World';
  s2 := s1;
  s2[1] := 'J';
  WriteLn(s1);
  WriteLn(s2);
end.
"#,
    );
    assert_eq!(out, vec!["Hello World", "Jello World"]);
}

#[test]
fn test_ansistring_uniquestring_forcing_unique_copy() {
    let out = run_pascal(
        r#"
program Test;
var s1, s2: AnsiString;
begin
  s1 := 'UniqueData';
  s2 := s1;
  UniqueString(s2);
  WriteLn(StringRefCount(s1));
  WriteLn(StringRefCount(s2));
end.
"#,
    );
    assert_eq!(out, vec!["1", "1"]);
}

#[test]
fn test_ansistring_empty_points_to_nil() {
    let out = run_pascal(
        r#"
program Test;
var s: AnsiString;
begin
  s := '';
  WriteLn(StringRefCount(s));
  WriteLn(Length(s));
end.
"#,
    );
    assert_eq!(out, vec!["0", "0"]);
}

#[test]
fn test_ansistring_setlength_expansion() {
    let out = run_pascal(
        r#"
program Test;
var s: AnsiString;
begin
  s := 'Base';
  SetLength(s, 10);
  WriteLn(Length(s));
end.
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_ansistring_setlength_truncation() {
    let out = run_pascal(
        r#"
program Test;
var s: AnsiString;
begin
  s := 'LongAnsiStringText';
  SetLength(s, 4);
  WriteLn(s);
  WriteLn(Length(s));
end.
"#,
    );
    assert_eq!(out, vec!["Long", "4"]);
}

#[test]
fn test_ansistring_pansichar_casting() {
    let out = run_pascal(
        r#"
program Test;
var s: AnsiString;
    p: PAnsiChar;
begin
  s := 'PAnsiCharPointerTest';
  p := PAnsiChar(s);
  WriteLn(p^);
  WriteLn(p);
end.
"#,
    );
    assert_eq!(out, vec!["P", "PAnsiCharPointerTest"]);
}

#[test]
fn test_ansistring_codepage_query() {
    let out = run_pascal(
        r#"
program Test;
var s: AnsiString;
begin
  s := 'CodePageText';
  WriteLn(StringCodePage(s) >= 0);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_ansistring_reassignment_decrements_refcount() {
    let out = run_pascal(
        r#"
program Test;
var s1, s2: AnsiString;
begin
  s1 := 'Shared';
  s2 := s1;
  WriteLn(StringRefCount(s1));
  s2 := 'Different';
  WriteLn(StringRefCount(s1));
end.
"#,
    );
    assert_eq!(out, vec!["2", "1"]);
}

#[test]
fn test_ansistring_pass_by_const_prevents_refcount_bump() {
    let out = run_pascal(
        r#"
program Test;
procedure InspectConst(const s: AnsiString);
begin
  WriteLn(StringRefCount(s));
end;
var text: AnsiString;
begin
  text := 'ConstText';
  InspectConst(text);
end.
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_ansistring_concatenation_allocates_new_buffer() {
    let out = run_pascal(
        r#"
program Test;
var s1, s2, s3: AnsiString;
begin
  s1 := 'Hello '; s2 := 'World';
  s3 := s1 + s2;
  WriteLn(s3);
  WriteLn(StringRefCount(s3));
end.
"#,
    );
    assert_eq!(out, vec!["Hello World", "1"]);
}

#[test]
fn test_ansistring_element_read_indexing() {
    let out = run_pascal(
        r#"
program Test;
var s: AnsiString;
begin
  s := 'Pascal';
  WriteLn(s[1]);
  WriteLn(s[6]);
end.
"#,
    );
    assert_eq!(out, vec!["P", "l"]);
}

#[test]
fn test_ansistring_comparisons() {
    let out = run_pascal(
        r#"
program Test;
var s1, s2: AnsiString;
begin
  s1 := 'Alpha'; s2 := 'Beta';
  WriteLn(s1 < s2);
  WriteLn(s1 = 'Alpha');
  WriteLn(s1 <> s2);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE", "TRUE", "TRUE"]);
}

#[test]
fn test_ansistring_array_elements_refcounting() {
    let out = run_pascal(
        r#"
program Test;
var arr: array[0..1] of AnsiString;
    shared: AnsiString;
begin
  shared := 'ArraySharedString';
  arr[0] := shared; arr[1] := shared;
  WriteLn(StringRefCount(shared));
end.
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_ansistring_record_field_cleanup() {
    let out = run_pascal(
        r#"
program Test;
type TRec = record Text: AnsiString; end;
procedure RunScope;
var r: TRec;
begin
  r.Text := 'RecordAnsiString';
  WriteLn(r.Text);
end;
begin
  RunScope;
end.
"#,
    );
    assert_eq!(out, vec!["RecordAnsiString"]);
}

#[test]
fn test_ansistring_class_field_destruction() {
    let out = run_pascal(
        r#"
program Test;
type TClassWithStr = class
  public Title: AnsiString;
  constructor Create(T: AnsiString);
end;
constructor TClassWithStr.Create(T: AnsiString); begin Title := T; end;
var obj: TClassWithStr;
begin
  obj := TClassWithStr.Create('ClassAnsiString');
  WriteLn(obj.Title);
  obj.Free;
end.
"#,
    );
    assert_eq!(out, vec!["ClassAnsiString"]);
}

#[test]
fn test_ansistring_function_return_value() {
    let out = run_pascal(
        r#"
program Test;
function BuildString(prefix, suffix: AnsiString): AnsiString;
begin
  Result := prefix + '_' + suffix;
end;
var res: AnsiString;
begin
  res := BuildString('Start', 'End');
  WriteLn(res);
end.
"#,
    );
    assert_eq!(out, vec!["Start_End"]);
}

#[test]
fn test_ansistring_clear_frees_heap_buffer() {
    let out = run_pascal(
        r#"
program Test;
var s: AnsiString;
begin
  s := 'TemporaryHeapBuffer';
  WriteLn(Length(s) > 0);
  s := '';
  WriteLn(Length(s));
  WriteLn(s = '');
end.
"#,
    );
    assert_eq!(out, vec!["TRUE", "0", "TRUE"]);
}

#[test]
fn test_ansistring_high_low_bounds() {
    let out = run_pascal(
        r#"
program Test;
var s: AnsiString;
begin
  s := 'Sample';
  WriteLn(Low(s));
  WriteLn(High(s));
end.
"#,
    );
    assert_eq!(out, vec!["1", "6"]);
}

#[test]
fn test_ansistring_loop_character_traversal() {
    let out = run_pascal(
        r#"
program Test;
var s: AnsiString;
    i: Integer;
    upperCount: Integer;
begin
  s := 'PascalLanguageTest';
  upperCount := 0;
  for i := 1 to Length(s) do
    if (s[i] >= 'A') and (s[i] <= 'Z') then Inc(upperCount);
  WriteLn(upperCount);
end.
"#,
    );
    assert_eq!(out, vec!["3"]);
}
