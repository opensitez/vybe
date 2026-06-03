/// String operations from standard Object Pascal / Delphi RTL.
/// Format(), Copy with length, Delete, Insert, Pos patterns,
/// char iteration, string building, common string algorithms.
use super::helpers::run_pascal;

// ===================================================================
// FORMAT FUNCTION
// ===================================================================

#[test]
fn format_integer() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Format('Value: %d', [42]));
end."#
        ),
        &["Value: 42"]
    );
}

#[test]
fn format_string() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Format('Hello %s!', ['World']));
end."#
        ),
        &["Hello World!"]
    );
}

#[test]
fn format_multiple() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Format('%s is %d years old', ['Alice', 30]));
end."#
        ),
        &["Alice is 30 years old"]
    );
}

#[test]
fn format_float() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Format('Pi is %.2f', [3.14159]));
end."#
        ),
        &["Pi is 3.14"]
    );
}

// ===================================================================
// COPY FUNCTION (Copy(s, index, count))
// ===================================================================

#[test]
fn copy_middle() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Copy('Hello World', 1, 5));
end."#
        ),
        &["Hello"]
    );
}

#[test]
fn copy_end() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Copy('Hello World', 7, 5));
end."#
        ),
        &["World"]
    );
}

#[test]
fn copy_single_char() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Copy('ABCDE', 3, 1));
end."#
        ),
        &["C"]
    );
}

// ===================================================================
// POS FUNCTION (find substring)
// ===================================================================

#[test]
fn pos_found() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Pos('World', 'Hello World'));
end."#
        ),
        &["7"]
    );
}

#[test]
fn pos_not_found() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Pos('XYZ', 'Hello World'));
end."#
        ),
        &["0"]
    );
}

#[test]
fn pos_at_start() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Pos('Hello', 'Hello World'));
end."#
        ),
        &["1"]
    );
}

// ===================================================================
// DELETE AND INSERT PROCEDURES
// ===================================================================

#[test]
fn delete_from_string() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: String;
begin
  s := 'Hello World';
  Delete(s, 6, 6);
  WriteLn(s);
end."#
        ),
        &["Hello"]
    );
}

#[test]
fn insert_into_string() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: String;
begin
  s := 'Hello World';
  Insert(' Beautiful', s, 6);
  WriteLn(s);
end."#
        ),
        &["Hello Beautiful World"]
    );
}

// ===================================================================
// CHAR OPERATIONS
// ===================================================================

#[test]
fn char_from_string_index() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: String;
begin
  s := 'ABCDE';
  WriteLn(s[1]);
  WriteLn(s[3]);
  WriteLn(s[5]);
end."#
        ),
        &["A", "C", "E"]
    );
}

#[test]
fn char_comparison() {
    assert_eq!(
        run_pascal(
            r#"program T;
var ch: String;
begin
  ch := 'B';
  if ch >= 'A' then WriteLn('yes');
  if ch < 'D' then WriteLn('also yes');
end."#
        ),
        &["yes", "also yes"]
    );
}

#[test]
fn iterate_chars_in_string() {
    assert_eq!(
        run_pascal(
            r#"program T;
var ch: String;
begin
  for ch in 'HELLO' do
    WriteLn(ch);
end."#
        ),
        &["H", "E", "L", "L", "O"]
    );
}

// ===================================================================
// STRING BUILDING PATTERNS
// ===================================================================

#[test]
fn build_csv() {
    assert_eq!(
        run_pascal(
            r#"program T;
var items: array of String;
var result: String;
var i: Integer;
begin
  items := ['apple', 'banana', 'cherry'];
  result := '';
  for i := 0 to High(items) do
  begin
    if i > 0 then result := result + ',';
    result := result + items[i];
  end;
  WriteLn(result);
end."#
        ),
        &["apple,banana,cherry"]
    );
}

#[test]
fn build_repeated_string() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: String;
var i: Integer;
begin
  s := '';
  for i := 1 to 5 do
    s := s + '*';
  WriteLn(s);
end."#
        ),
        &["*****"]
    );
}

#[test]
fn string_contains_check() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Contains(haystack, needle: String): Boolean;
begin
  Result := Pos(needle, haystack) > 0;
end;
begin
  if Contains('Hello World', 'World') then WriteLn('found');
  if not Contains('Hello World', 'xyz') then WriteLn('not found');
end."#
        ),
        &["found", "not found"]
    );
}

#[test]
fn string_starts_with() {
    assert_eq!(
        run_pascal(
            r#"program T;
function StartsWith(s, prefix: String): Boolean;
begin
  Result := Copy(s, 1, Length(prefix)) = prefix;
end;
begin
  WriteLn(StartsWith('Hello World', 'Hello'));
  WriteLn(StartsWith('Hello World', 'World'));
end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn string_ends_with() {
    assert_eq!(
        run_pascal(
            r#"program T;
function EndsWith(s, suffix: String): Boolean;
begin
  Result := Copy(s, Length(s) - Length(suffix) + 1, Length(suffix)) = suffix;
end;
begin
  WriteLn(EndsWith('Hello World', 'World'));
  WriteLn(EndsWith('Hello World', 'Hello'));
end."#
        ),
        &["true", "false"]
    );
}

// ===================================================================
// STRING REPLACE PATTERNS
// ===================================================================

#[test]
fn string_replace_all() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(StringReplace('aabaa', 'a', 'x'));
end."#
        ),
        &["xxbxx"]
    );
}

// ===================================================================
// STRING CONVERSION PATTERNS
// ===================================================================

#[test]
fn int_to_str_and_back() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: String;
var n: Integer;
begin
  n := 42;
  s := IntToStr(n);
  WriteLn(s);
  n := StrToInt(s) * 2;
  WriteLn(n);
end."#
        ),
        &["42", "84"]
    );
}

#[test]
fn float_to_str_and_back() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: String;
var f: Real;
begin
  f := 3.14;
  s := FloatToStr(f);
  WriteLn(s);
end."#
        ),
        &["3.14"]
    );
}

// ===================================================================
// STRING LENGTH PATTERNS
// ===================================================================

#[test]
fn empty_string_length() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: String;
begin
  s := '';
  if Length(s) = 0 then WriteLn('empty');
end."#
        ),
        &["empty"]
    );
}

#[test]
fn string_as_condition() {
    assert_eq!(
        run_pascal(
            r#"program T;
var name: String;
begin
  name := 'Alice';
  if Length(name) > 0 then WriteLn('has name')
  else WriteLn('no name');
end."#
        ),
        &["has name"]
    );
}
