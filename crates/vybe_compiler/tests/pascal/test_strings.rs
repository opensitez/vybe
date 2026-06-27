/// Tests for Pascal string operations beyond basic builtins.
use super::helpers::run_pascal;

// ===================================================================
// POS — find substring (1-based, 0 if not found)
// ===================================================================

#[test]
fn str_pos_found() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Pos('lo', 'hello')); end."),
        &["4"]
    );
}

#[test]
fn str_pos_not_found() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Pos('xyz', 'hello')); end."),
        &["0"]
    );
}

#[test]
fn str_pos_at_start() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Pos('he', 'hello')); end."),
        &["1"]
    );
}

// ===================================================================
// COPY — extract substring (1-based)
// ===================================================================

#[test]
fn str_copy_middle() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Copy('hello world', 7, 5)); end."),
        &["world"]
    );
}

#[test]
fn str_copy_from_start() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Copy('abcdef', 1, 3)); end."),
        &["abc"]
    );
}

// ===================================================================
// RIGHTSTR — right substring
// ===================================================================

#[test]
fn str_rightstr() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(RightStr('hello', 3)); end."),
        &["llo"]
    );
}

#[test]
fn str_rightstr_full() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(RightStr('abc', 3)); end."),
        &["abc"]
    );
}

// ===================================================================
// CHR / ORD — character conversion
// ===================================================================

#[test]
fn str_chr() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Chr(65)); end."),
        &["A"]
    );
}

#[test]
fn str_ord() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Ord('A')); end."),
        &["65"]
    );
}

#[test]
fn str_chr_ord_roundtrip() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Chr(Ord('Z'))); end."),
        &["Z"]
    );
}

// ===================================================================
// TRIMLEFT / TRIMRIGHT
// ===================================================================

#[test]
fn str_trimleft() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(TrimLeft('  hi  ')); end."),
        &["hi  "]
    );
}

#[test]
fn str_trimright() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(TrimRight('  hi  ')); end."),
        &["  hi"]
    );
}

// ===================================================================
// BOOLTOSTR / STRTOBOOL
// ===================================================================

#[test]
fn str_booltostr() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(BoolToStr(true)); end."),
        &["true"]
    );
}

#[test]
fn str_booltostr_false() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(BoolToStr(false)); end."),
        &["false"]
    );
}

// ===================================================================
// STRING INDEXING (Pascal strings are 1-based)
// ===================================================================

#[test]
fn str_index_first_char() {
    assert_eq!(
        run_pascal("program T; var s: String; begin s := 'hello'; WriteLn(s[1]); end."),
        &["h"]
    );
}

#[test]
fn str_index_last_char() {
    assert_eq!(
        run_pascal("program T; var s: String; begin s := 'hello'; WriteLn(s[5]); end."),
        &["o"]
    );
}

// ===================================================================
// COMPARESTR
// ===================================================================

#[test]
fn str_comparestr_equal() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(CompareStr('abc', 'abc')); end."),
        &["0"]
    );
}

// ===================================================================
// STRING IN EXPRESSIONS
// ===================================================================

#[test]
fn str_length_in_loop() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: String; i: Integer;
begin
  s := 'abc';
    for i := 1 to Length(s) do WriteLn(s[i]);
end."#
        ),
        &["a", "b", "c"]
    );
}

#[test]
fn str_build_reverse() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s, r: String; i: Integer;
begin
  s := 'abcd';
  r := '';
    for i := Length(s) downto 1 do r := r + s[i];
  WriteLn(r);
end."#
        ),
        &["dcba"]
    );
}

// -------------------------------------------------------------------
// from test_strings_pos_copy_delete.rs
// -------------------------------------------------------------------
#[test]
fn copy_extracts_middle_substring() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Copy('abcdef', 2, 3)); end."#
        ),
        &["bcd"]
    );
}

#[test]
fn length_counts_characters() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Length('pascal')); end."#
        ),
        &["6"]
    );
}

#[test]
fn pos_finds_substring_offset() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Pos('ca', 'pascal')); end."#
        ),
        &["3"]
    );
}

#[test]
fn pos_returns_zero_when_missing() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Pos('z', 'pascal')); end."#
        ),
        &["0"]
    );
}

#[test]
fn delete_removes_characters_at_index() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: String;
begin
  s := 'abcde';
  Delete(s, 2, 2);
  WriteLn(s);
end."#
        ),
        &["ade"]
    );
}

#[test]
fn insert_places_text_at_position() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: String;
begin
  s := 'ab';
  Insert('XY', s, 2);
  WriteLn(s);
end."#
        ),
        &["aXYb"]
    );
}

#[test]
fn uppercase_converts_ascii_letters() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(UpperCase('AbC')); end."#
        ),
        &["ABC"]
    );
}

#[test]
fn lowercase_converts_ascii_letters() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(LowerCase('AbC')); end."#
        ),
        &["abc"]
    );
}

#[test]
fn trim_removes_outer_spaces() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Trim('  hi  ')); end."#
        ),
        &["hi"]
    );
}

#[test]
fn string_concat_with_plus() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn('foo' + '-' + 'bar'); end."#
        ),
        &["foo-bar"]
    );
}

#[test]
fn string_replace_substitutes_first_occurrence() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(StringReplace('a-b-a', '-', '+', [])); end."#
        ),
        &["a+b-a"]
    );
}

#[test]
fn string_replace_all_occurrences_flag() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(StringReplace('a-b-a', '-', '+', [rfReplaceAll])); end."#
        ),
        &["a+b+a"]
    );
}

#[test]
fn int_to_str_formats_integer() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(IntToStr(1204)); end."#
        ),
        &["1204"]
    );
}

#[test]
fn str_to_int_def_returns_default_on_invalid() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(StrToIntDef('xy', 9)); end."#
        ),
        &["9"]
    );
}

#[test]
fn quoted_str_wraps_in_double_quotes() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(QuotedStr('hi')); end."#
        ),
        &["\"hi\""]
    );
}

#[test]
fn compare_text_case_insensitive_equal() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(CompareText('AbC', 'abc')); end."#
        ),
        &["0"]
    );
}

#[test]
fn same_text_ignores_case() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(SameText('Hello', 'hello')); end."#
        ),
        &["true"]
    );
}

#[test]
fn format_inserts_multiple_placeholders() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Format('%s=%d', ['x', 7])); end."#
        ),
        &["x=7"]
    );
}

#[test]
fn string_of_char_repeats_character() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(StringOfChar('*', 4)); end."#
        ),
        &["****"]
    );
}

#[test]
fn trim_left_removes_leading_spaces_only() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(TrimLeft('  hi ')); end."#
        ),
        &["hi "]
    );
}

#[test]
fn trim_right_removes_trailing_spaces_only() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(TrimRight(' hi  ')); end."#
        ),
        &[" hi"]
    );
}

#[test]
fn copy_string_subrange() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Copy('abcdef', 2, 3)); end."#
        ),
        &["bcd"]
    );
}

#[test]
fn pos_finds_substring_index() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Pos('lo', 'hello')); end."#
        ),
        &["4"]
    );
}

#[test]
fn delete_removes_characters_from_string() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Delete('abcdef', 2, 2)); end."#
        ),
        &["adef"]
    );
}

#[test]
fn insert_puts_text_at_position() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Insert('XX', 'ab', 2)); end."#
        ),
        &["aXXb"]
    );
}

#[test]
fn string_concatenation_with_plus() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn('foo' + 'bar'); end."#
        ),
        &["foobar"]
    );
}



