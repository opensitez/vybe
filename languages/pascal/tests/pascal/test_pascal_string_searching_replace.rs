use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 36: String Searching & Replacement (Pos, PosEx, StringReplace)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_pos_substring_index() {
    let out = run_pascal(
        r#"
program Test;
var idx: Integer;
begin
  idx := Pos('World', 'Hello World');
  WriteLn(idx);
end.
"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn test_pos_not_found_returns_zero() {
    let out = run_pascal(
        r#"
program Test;
var idx: Integer;
begin
  idx := Pos('Pascal', 'Hello World');
  WriteLn(idx);
end.
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_posex_with_start_offset() {
    let out = run_pascal(
        r#"
program Test;
uses StrUtils;
var idx1, idx2: Integer;
    s: String;
begin
  s := 'abc abc abc';
  idx1 := PosEx('abc', s, 1);
  idx2 := PosEx('abc', s, 4);
  WriteLn(idx1);
  WriteLn(idx2);
end.
"#,
    );
    assert_eq!(out, vec!["1", "5"]);
}

#[test]
fn test_stringreplace_replace_all() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var res: String;
begin
  res := StringReplace('foo bar foo baz', 'foo', 'qux', [rfReplaceAll]);
  WriteLn(res);
end.
"#,
    );
    assert_eq!(out, vec!["qux bar qux baz"]);
}

#[test]
fn test_stringreplace_first_occurrence_only() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var res: String;
begin
  res := StringReplace('foo bar foo baz', 'foo', 'qux', []);
  WriteLn(res);
end.
"#,
    );
    assert_eq!(out, vec!["qux bar foo baz"]);
}

#[test]
fn test_stringreplace_ignore_case() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var res: String;
begin
  res := StringReplace('Foo bar FOO baz', 'foo', 'qux', [rfReplaceAll, rfIgnoreCase]);
  WriteLn(res);
end.
"#,
    );
    assert_eq!(out, vec!["qux bar qux baz"]);
}

#[test]
fn test_replacestr_case_sensitive() {
    let out = run_pascal(
        r#"
program Test;
uses StrUtils;
begin
  WriteLn(ReplaceStr('Test test Test', 'Test', 'Pass'));
end.
"#,
    );
    assert_eq!(out, vec!["Pass test Pass"]);
}

#[test]
fn test_replacetext_case_insensitive() {
    let out = run_pascal(
        r#"
program Test;
uses StrUtils;
begin
  WriteLn(ReplaceText('Test test TEST', 'test', 'Pass'));
end.
"#,
    );
    assert_eq!(out, vec!["Pass Pass Pass"]);
}

#[test]
fn test_pos_single_char_search() {
    let out = run_pascal(
        r#"
program Test;
begin
  WriteLn(Pos('X', 'ABCXDEF'));
end.
"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn test_pos_at_start_of_string() {
    let out = run_pascal(
        r#"
program Test;
begin
  WriteLn(Pos('Start', 'StartOfText'));
end.
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_pos_at_end_of_string() {
    let out = run_pascal(
        r#"
program Test;
var s: String;
begin
  s := 'TextAtTheEnd';
  WriteLn(Pos('End', s));
end.
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_posex_loop_count_occurrences() {
    let out = run_pascal(
        r#"
program Test;
uses StrUtils;
var text: String; idx, count: Integer;
begin
  text := 'banana';
  count := 0;
  idx := PosEx('an', text, 1);
  while idx > 0 do
  begin
    Inc(count);
    idx := PosEx('an', text, idx + 1);
  end;
  WriteLn(count);
end.
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_stringreplace_deletion() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(StringReplace('A-B-C-D', '-', '', [rfReplaceAll]));
end.
"#,
    );
    assert_eq!(out, vec!["ABCD"]);
}

#[test]
fn test_stringreplace_expansion() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(StringReplace('A B C', ' ', '___', [rfReplaceAll]));
end.
"#,
    );
    assert_eq!(out, vec!["A___B___C"]);
}

#[test]
fn test_pos_case_insensitive_pattern() {
    let out = run_pascal(
        r#"
program Test;
function PosCaseInsensitive(const sub, main: String): Integer;
begin
  Result := Pos(LowerCase(sub), LowerCase(main));
end;
begin
  WriteLn(PosCaseInsensitive('world', 'HELLO WORLD'));
end.
"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn test_stringreplace_with_quotes() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(StringReplace('Hello "World"', '"World"', 'Pascal', [rfReplaceAll]));
end.
"#,
    );
    assert_eq!(out, vec!["Hello Pascal"]);
}

#[test]
fn test_pos_in_shortstring() {
    let out = run_pascal(
        r#"
program Test;
var ss: ShortString;
begin
  ss := 'ShortStringPattern';
  WriteLn(Pos('Pattern', ss));
end.
"#,
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn test_stringreplace_in_array_elements() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var arr: array[0..1] of String;
begin
  arr[0] := 'item_old'; arr[1] := 'data_old';
  arr[0] := StringReplace(arr[0], 'old', 'new', [rfReplaceAll]);
  arr[1] := StringReplace(arr[1], 'old', 'new', [rfReplaceAll]);
  WriteLn(arr[0] + ' ' + arr[1]);
end.
"#,
    );
    assert_eq!(out, vec!["item_new data_new"]);
}

#[test]
fn test_pos_empty_pattern_returns_zero() {
    let out = run_pascal(
        r#"
program Test;
begin
  WriteLn(Pos('', 'Text'));
end.
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_stringreplace_no_match_returns_original() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(StringReplace('UnchangedText', 'Missing', 'Replacement', [rfReplaceAll]));
end.
"#,
    );
    assert_eq!(out, vec!["UnchangedText"]);
}
