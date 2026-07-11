/// String compare, case conversion, padding, and replace operations.
use super::helpers::run_pascal;

#[test]
fn comparestr_equal_returns_zero() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(CompareStr('abc','abc')); end."#),
        &["0"]
    );
}

#[test]
fn comparestr_less_returns_negative() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(CompareStr('abc','abd') < 0); end."#),
        &["true"]
    );
}

#[test]
fn comparestr_greater_returns_positive() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(CompareStr('z','a') > 0); end."#),
        &["true"]
    );
}

#[test]
fn comparetext_case_insensitive_equal() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(CompareText('Hello','hello')); end."#),
        &["0"]
    );
}

#[test]
fn comparetext_mixed_case_ordering() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(CompareText('Beta','alpha') > 0); end."#),
        &["true"]
    );
}

#[test]
fn sametext_true_for_case_variant() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(SameText('Vybe','vybe')); end."#),
        &["true"]
    );
}

#[test]
fn samestr_false_when_different() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(SameStr('a','b')); end."#),
        &["false"]
    );
}

#[test]
fn uppercase_all_lower_input() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(UpperCase('delphi')); end."#),
        &["DELPHI"]
    );
}

#[test]
fn lowercase_all_upper_input() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(LowerCase('PASCAL')); end."#),
        &["pascal"]
    );
}

#[test]
fn ansiuppercase_mixed_string() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(AnsiUpperCase('AbC123')); end."#),
        &["ABC123"]
    );
}

#[test]
fn ansilowercase_mixed_string() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(AnsiLowerCase('AbC123')); end."#),
        &["abc123"]
    );
}

#[test]
fn trim_leading_spaces_removed() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trim('   left')); end."#),
        &["left"]
    );
}

#[test]
fn trim_trailing_spaces_removed() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trim('right   ')); end."#),
        &["right"]
    );
}

#[test]
fn trim_both_ends() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trim('  mid  ')); end."#),
        &["mid"]
    );
}

#[test]
fn stringreplace_first_occurrence_only() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(StringReplace('a-a-a','-','+',[])); end."#),
        &["a+a-a"]
    );
}

#[test]
fn stringreplace_all_occurrences() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(StringReplace('a-a-a','-','+',[rfReplaceAll])); end."#
        ),
        &["a+a+a"]
    );
}

#[test]
fn stringreplace_empty_search_unchanged() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(StringReplace('hello','','x',[])); end."#),
        &["hello"]
    );
}

#[test]
fn pad_left_zeros_width_five() {
    assert_eq!(
        run_pascal(
            r#"program T;
function PadLeft(s:String;width:Integer;ch:Char):String;
var n:Integer;
begin
  n:=width-Length(s);
  if n<=0 then Result:=s else Result:=StringOfChar(ch,n)+s;
end;
begin WriteLn(PadLeft('42',5,'0')); end."#
        ),
        &["00042"]
    );
}

#[test]
fn pad_right_dots_width_four() {
    assert_eq!(
        run_pascal(
            r#"program T;
function PadRight(s:String;width:Integer;ch:Char):String;
var n:Integer;
begin
  n:=width-Length(s);
  if n<=0 then Result:=s else Result:=s+StringOfChar(ch,n);
end;
begin WriteLn(PadRight('go',4,'.')); end."#
        ),
        &["go.."]
    );
}

#[test]
fn stringofchar_repeat_dash() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(StringOfChar('-',5)); end."#),
        &["-----"]
    );
}

#[test]
fn quotedstr_wraps_simple_word() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(QuotedStr('hi')); end."#),
        &["'hi'"]
    );
}

#[test]
fn comparestr_empty_strings() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(CompareStr('','')); end."#),
        &["0"]
    );
}

#[test]
fn comparestr_prefix_is_less() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(CompareStr('app','apple') < 0); end."#),
        &["true"]
    );
}

#[test]
fn comparetext_empty_vs_nonempty() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(CompareText('','A') < 0); end."#),
        &["true"]
    );
}

#[test]
fn uppercase_preserves_digits() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(UpperCase('a1b2')); end."#),
        &["A1B2"]
    );
}

#[test]
fn lowercase_preserves_symbols() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(LowerCase('A_B')); end."#),
        &["a_b"]
    );
}

#[test]
fn stringreplace_word_boundary_style() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(StringReplace('cat catalog','cat','dog',[])); end."#
        ),
        &["dog catalog"]
    );
}

#[test]
fn trim_no_spaces_unchanged() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Trim('none')); end."#),
        &["none"]
    );
}

#[test]
fn copy_then_uppercase_pipeline() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(UpperCase(Copy('hello world',1,5))); end."#),
        &["HELLO"]
    );
}

#[test]
fn pos_then_copy_extract_token() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:Integer; begin p:=Pos('@','user@host'); WriteLn(Copy('user@host',1,p-1)); end."#
        ),
        &["user"]
    );
}

#[test]
fn comparestr_length_difference() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(CompareStr('aa','aaa') < 0); end."#),
        &["true"]
    );
}

#[test]
fn sametext_false_for_different_words() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(SameText('foo','bar')); end."#),
        &["false"]
    );
}

#[test]
fn pad_left_no_padding_when_wide_enough() {
    assert_eq!(
        run_pascal(
            r#"program T;
function PadLeft(s:String;width:Integer;ch:Char):String;
begin if Length(s)>=width then Result:=s else Result:=StringOfChar(ch,width-Length(s))+s; end;
begin WriteLn(PadLeft('hello',3,' ')); end."#
        ),
        &["hello"]
    );
}

#[test]
fn stringreplace_delete_by_empty() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(StringReplace('abc','b','',[rfReplaceAll])); end."#),
        &["ac"]
    );
}

#[test]
fn concat_then_comparestr_equal() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(CompareStr('ab'+'c','abc')=0); end."#),
        &["true"]
    );
}

#[test]
fn lowercase_then_comparetext_self() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(CompareText(LowerCase('MiXed'),'mixed')); end."#),
        &["0"]
    );
}

#[test]
fn quotedstr_on_empty_string() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Length(QuotedStr(''))); end."#),
        &["2"]
    );
}

#[test]
fn stringofchar_zero_length_empty() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Length(StringOfChar('x',0))); end."#),
        &["0"]
    );
}

#[test]
fn trim_left_only_via_custom() {
    assert_eq!(
        run_pascal(
            r#"program T;
function TrimLeft(const s:String):String;
var i:Integer;
begin i:=1; while (i<=Length(s)) and (s[i]=' ') do Inc(i); Result:=Copy(s,i,MaxInt); end;
begin WriteLn(TrimLeft('  x')); end."#
        ),
        &["x"]
    );
}

#[test]
fn replace_then_length_check() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:String; begin s:=StringReplace('aaaa','a','b',[rfReplaceAll]); WriteLn(Length(s)); end."#
        ),
        &["4"]
    );
}
