/// Tests for extra string functions and patterns in Pascal/Delphi:
/// CompareStr, CompareText, SameText, SameStr, AnsiUpperCase,
/// AnsiLowerCase, StringOfChar variations, Trim edge cases,
/// string reversal, splitting, joining patterns.
use super::helpers::run_pascal;

// ===================================================================
// COMPARESTR / COMPARETEXT
// ===================================================================

#[test]
fn comparestr_equal() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(CompareStr('hello', 'hello') = 0);
end."#
        ),
        &["true"]
    );
}

#[test]
fn comparestr_less() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(CompareStr('abc', 'abd') < 0);
end."#
        ),
        &["true"]
    );
}

#[test]
fn comparestr_greater() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(CompareStr('z', 'a') > 0);
end."#
        ),
        &["true"]
    );
}

#[test]
fn comparetext_case_insensitive() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(CompareText('Hello', 'hello') = 0);
end."#
        ),
        &["true"]
    );
}

#[test]
fn sametext_basic() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(SameText('ABC', 'abc'));
  WriteLn(SameText('abc', 'xyz'));
end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn samestr_case_sensitive() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(SameStr('Hello', 'Hello'));
  WriteLn(SameStr('Hello', 'hello'));
end."#
        ),
        &["true", "false"]
    );
}

// ===================================================================
// ANSI UPPER/LOWERCASE
// ===================================================================

#[test]
fn ansiuppercase_basic() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(AnsiUpperCase('hello world'));
end."#
        ),
        &["HELLO WORLD"]
    );
}

#[test]
fn ansilowercase_basic() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(AnsiLowerCase('HELLO WORLD'));
end."#
        ),
        &["hello world"]
    );
}

// ===================================================================
// STRINGOFCHAR VARIATIONS
// ===================================================================

#[test]
fn stringofchar_dash() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(StringOfChar('-', 10));
end."#
        ),
        &["----------"]
    );
}

#[test]
fn stringofchar_zero() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Length(StringOfChar('x', 0)));
end."#
        ),
        &["0"]
    );
}

#[test]
fn stringofchar_separator() {
    assert_eq!(
        run_pascal(
            r#"program T;
var sep: String;
begin
  sep := StringOfChar('=', 20);
  WriteLn(Length(sep));
end."#
        ),
        &["20"]
    );
}

// ===================================================================
// TRIM EDGE CASES
// ===================================================================

#[test]
fn trim_only_spaces() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Length(Trim('     ')));
end."#
        ),
        &["0"]
    );
}

#[test]
fn trim_no_spaces() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Trim('hello'));
end."#
        ),
        &["hello"]
    );
}

#[test]
fn trimleft_right() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(TrimLeft('  hello  '));
  WriteLn(TrimRight('  hello  '));
end."#
        ),
        &["hello  ", "  hello"]
    );
}

// ===================================================================
// STRING REVERSAL
// ===================================================================

#[test]
fn reverse_string_function() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Reverse(s: String): String;
var i: Integer;
begin
  Result := '';
  for i := Length(s) downto 1 do
    Result := Result + s[i];
end;
begin
  WriteLn(Reverse('pascal'));
  WriteLn(Reverse(''));
end."#
        ),
        &["lacsap", ""]
    );
}

// ===================================================================
// STRING SPLITTING PATTERN
// ===================================================================

#[test]
fn split_by_comma_count() {
    assert_eq!(
        run_pascal(
            r#"program T;
function CountParts(s, delim: String): Integer;
var i: Integer;
begin
  Result := 1;
  for i := 1 to Length(s) do
    if s[i] = delim[1] then Inc(Result);
end;
begin
  WriteLn(CountParts('a,b,c,d', ','));
end."#
        ),
        &["4"]
    );
}

#[test]
fn extract_first_word() {
    assert_eq!(
        run_pascal(
            r#"program T;
function FirstWord(s: String): String;
var p: Integer;
begin
  p := Pos(' ', s);
  if p > 0 then Result := Copy(s, 1, p - 1)
  else Result := s;
end;
begin
  WriteLn(FirstWord('hello world'));
  WriteLn(FirstWord('single'));
end."#
        ),
        &["hello", "single"]
    );
}

// ===================================================================
// STRING JOINING PATTERN
// ===================================================================

#[test]
fn join_with_separator() {
    assert_eq!(
        run_pascal(
            r#"program T;
var parts: array[1..3] of String;
    result: String;
    i: Integer;
begin
  parts[1] := 'one';
  parts[2] := 'two';
  parts[3] := 'three';
  result := '';
  for i := 1 to 3 do
  begin
    if i > 1 then result := result + ', ';
    result := result + parts[i];
  end;
  WriteLn(result);
end."#
        ),
        &["one, two, three"]
    );
}

// ===================================================================
// STRING PAD / ALIGN PATTERN
// ===================================================================

#[test]
fn pad_left_pattern() {
    assert_eq!(
        run_pascal(
            r#"program T;
function PadLeft(s: String; width: Integer; ch: Char): String;
begin
  while Length(s) < width do
    s := ch + s;
  Result := s;
end;
begin
  WriteLn(PadLeft('42', 5, '0'));
  WriteLn(PadLeft('hello', 5, ' '));
end."#
        ),
        &["00042", "hello"]
    );
}

#[test]
fn pad_right_pattern() {
    assert_eq!(
        run_pascal(
            r#"program T;
function PadRight(s: String; width: Integer; ch: Char): String;
begin
  while Length(s) < width do
    s := s + ch;
  Result := s;
end;
begin
  WriteLn(PadRight('hi', 5, '.'));
end."#
        ),
        &["hi..."]
    );
}

// ===================================================================
// STRING CONTAINS PATTERN
// ===================================================================

#[test]
fn string_count_char() {
    assert_eq!(
        run_pascal(
            r#"program T;
function CountChar(s: String; c: Char): Integer;
var i: Integer;
begin
  Result := 0;
  for i := 1 to Length(s) do
    if s[i] = c then Inc(Result);
end;
begin
  WriteLn(CountChar('mississippi', 's'));
end."#
        ),
        &["4"]
    );
}
