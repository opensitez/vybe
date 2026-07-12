/// Advanced string behaviors: Copy/Insert/Delete/Pos/Trim/SameText and related.
use super::helpers::run_pascal;

#[test]
fn copy_beyond_length_returns_partial() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Copy('abc', 2, 10)); end."#),
        &["bc"]
    );
}

#[test]
fn copy_zero_count_yields_empty() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Length(Copy('hello', 1, 0))); end."#),
        &["0"]
    );
}

#[test]
fn copy_from_past_end_yields_empty() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Copy('hi', 5, 2)); end."#),
        &["0"]
    );
}

#[test]
fn insert_at_beginning() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:string; begin s:='world'; Insert('hello ', s, 1); WriteLn(s); end."#
        ),
        &["hello world"]
    );
}

#[test]
fn insert_at_end_position() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:string; begin s:='ab'; Insert('cd', s, 3); WriteLn(s); end."#
        ),
        &["abcd"]
    );
}

#[test]
fn delete_middle_substring() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:string; begin s:='abcdef'; Delete(s, 2, 3); WriteLn(s); end."#
        ),
        &["af"]
    );
}

#[test]
fn delete_from_start() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:string; begin s:='prefix_rest'; Delete(s, 1, 7); WriteLn(s); end."#
        ),
        &["rest"]
    );
}

#[test]
fn pos_case_sensitive_no_match() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Pos('A', 'alpha')); end."#),
        &["0"]
    );
}

#[test]
fn pos_overlapping_needle() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Pos('aa', 'aaa')); end."#),
        &["1"]
    );
}

#[test]
fn pos_single_char_at_end() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Pos('z', 'xyzz')); end."#),
        &["3"]
    );
}

#[test]
fn trim_only_left_spaces() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(TrimLeft('   data')); end."#),
        &["data"]
    );
}

#[test]
fn trim_only_right_spaces() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(TrimRight('data   ')); end."#),
        &["data"]
    );
}

#[test]
fn trim_inner_spaces_preserved() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trim('  a b  ')); end."#),
        &["a b"]
    );
}

#[test]
fn sametext_different_strings() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(SameText('abc', 'xyz')); end."#),
        &["false"]
    );
}

#[test]
fn sametext_mixed_case_true() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(SameText('Pascal', 'PASCAL')); end."#),
        &["true"]
    );
}

#[test]
fn samestr_case_sensitive_true() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(SameStr('ok', 'ok')); end."#),
        &["true"]
    );
}

#[test]
fn samestr_case_sensitive_false() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(SameStr('ok', 'OK')); end."#),
        &["false"]
    );
}

#[test]
fn leftstr_more_than_length() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(LeftStr('ab', 5)); end."#),
        &["ab"]
    );
}

#[test]
fn rightstr_zero_count_empty() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Length(RightStr('abc', 0))); end."#),
        &["0"]
    );
}

#[test]
fn stringofchar_repeat_five() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(StringOfChar('x', 5)); end."#),
        &["xxxxx"]
    );
}

#[test]
fn stringofchar_zero_length() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Length(StringOfChar('-', 0))); end."#),
        &["0"]
    );
}

#[test]
fn quotedstr_wraps_quotes() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(QuotedStr('hi')); end."#),
        &["'hi'"]
    );
}

#[test]
fn ansilowercase_basic() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(AnsiLowerCase('AbC')); end."#),
        &["abc"]
    );
}

#[test]
fn ansiuppercase_basic() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(AnsiUpperCase('AbC')); end."#),
        &["ABC"]
    );
}

#[test]
fn comparetext_insensitive_order() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(CompareText('b', 'A')); end."#),
        &["1"]
    );
}

#[test]
fn comparestr_sensitive_equal() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(CompareStr('a', 'a')); end."#),
        &["0"]
    );
}

#[test]
fn stringreplace_single_occurrence() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(StringReplace('one two', 'two', '2', [])); end."#),
        &["one 2"]
    );
}

#[test]
fn stringreplace_replace_all_flag() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(StringReplace('a-a-a', '-', '', [rfReplaceAll])); end."#
        ),
        &["aaa"]
    );
}

#[test]
fn format_string_with_int() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Format('n=%d', [42])); end."#),
        &["n=42"]
    );
}

#[test]
fn format_string_with_two_fields() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Format('%s:%d', ['id', 7])); end."#),
        &["id:7"]
    );
}

#[test]
fn length_after_concat() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Length('ab'+'cd')); end."#),
        &["4"]
    );
}

#[test]
fn empty_string_length_zero() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Length('')); end."#),
        &["0"]
    );
}

#[test]
fn copy_then_uppercase() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(AnsiUpperCase(Copy('Hello', 1, 1))); end."#),
        &["H"]
    );
}

#[test]
fn pos_drives_copy_slice() {
    assert_eq!(
        run_pascal(
            r#"program T; var s,p:Integer; t:string; begin s:='name=value'; p:=Pos('=', s); t:=Copy(s, p+1, Length(s)); WriteLn(t); end."#
        ),
        &["value"]
    );
}

#[test]
fn delete_all_via_loop_pattern() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:string; begin s:='a-b-c'; Delete(s, 2, 1); Delete(s, 3, 1); WriteLn(s); end."#
        ),
        &["abc"]
    );
}

#[test]
fn insert_builds_path() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:string; begin p:='file'; Insert('.txt', p, Length(p)+1); WriteLn(p); end."#
        ),
        &["file.txt"]
    );
}

#[test]
fn trim_then_length() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Length(Trim('  x  '))); end."#),
        &["1"]
    );
}

#[test]
fn sametext_empty_strings() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(SameText('', '')); end."#),
        &["true"]
    );
}

#[test]
fn copy_full_string_via_length() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:string; begin s:='clone'; WriteLn(Copy(s, 1, Length(s))); end."#
        ),
        &["clone"]
    );
}

#[test]
fn pos_in_empty_haystack() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Pos('a', '')); end."#),
        &["0"]
    );
}

#[test]
fn insert_empty_fragment_no_change() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:string; begin s:='keep'; Insert('', s, 1); WriteLn(s); end."#
        ),
        &["keep"]
    );
}
