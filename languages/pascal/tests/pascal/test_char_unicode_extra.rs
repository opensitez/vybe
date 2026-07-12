/// UpCase, LowerCase, and char range operations.
use super::helpers::run_pascal;

#[test]
fn upcase_lowercase_a() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(UpCase('a')); end."#),
        &["A"]
    );
}

#[test]
fn upcase_already_upper_z() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(UpCase('Z')); end."#),
        &["Z"]
    );
}

#[test]
fn upcase_digit_unchanged() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(UpCase('7')); end."#),
        &["7"]
    );
}

#[test]
fn upcase_symbol_unchanged() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(UpCase('@')); end."#),
        &["@"]
    );
}

#[test]
fn lowercase_upper_m() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(LowerCase('M')); end."#),
        &["m"]
    );
}

#[test]
fn lowercase_already_lower() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(LowerCase('k')); end."#),
        &["k"]
    );
}

#[test]
fn lowercase_digit_unchanged() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(LowerCase('0')); end."#),
        &["0"]
    );
}

#[test]
fn lowercase_string_word() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(LowerCase('HELLO')); end."#),
        &["hello"]
    );
}

#[test]
fn uppercase_string_word() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(UpperCase('hello')); end."#),
        &["HELLO"]
    );
}

#[test]
fn upcase_var_char_field() {
    assert_eq!(
        run_pascal(r#"program T; var c:Char; begin c:='q'; WriteLn(UpCase(c)); end."#),
        &["Q"]
    );
}

#[test]
fn char_range_lowercase_in_set() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:set of Char; begin s:=['a'..'z']; WriteLn(Ord('m' in s)); WriteLn(Ord('M' in s)); end."#
        ),
        &["1", "0"]
    );
}

#[test]
fn char_range_uppercase_in_set() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:set of Char; begin s:=['A'..'Z']; WriteLn(Ord('B' in s)); end."#
        ),
        &["1"]
    );
}

#[test]
fn char_range_digits_in_set() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:set of Char; begin s:=['0'..'9']; WriteLn(Ord('4' in s)); WriteLn(Ord('a' in s)); end."#
        ),
        &["1", "0"]
    );
}

#[test]
fn char_range_ascii_printable_slice() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:set of Char; c:Char; n:Integer; begin s:=['!'..'~']; n:=0; for c in s do Inc(n); WriteLn(Ord(n>90)); end."#
        ),
        &["1"]
    );
}

#[test]
fn char_compare_less_ascii() {
    assert_eq!(
        run_pascal(r#"program T; begin if 'A'<'B' then WriteLn('yes') else WriteLn('no'); end."#),
        &["yes"]
    );
}

#[test]
fn char_compare_greater_digit() {
    assert_eq!(
        run_pascal(r#"program T; begin if '9'>'0' then WriteLn('yes') else WriteLn('no'); end."#),
        &["yes"]
    );
}

#[test]
fn char_ord_chr_roundtrip() {
    assert_eq!(
        run_pascal(r#"program T; var c:Char; begin c:=Chr(66); WriteLn(c); WriteLn(Ord(c)); end."#),
        &["B", "66"]
    );
}

#[test]
fn upcase_then_lower_back() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:='f'; c:=UpCase(c); c:=LowerCase(c); WriteLn(c); end."#
        ),
        &["f"]
    );
}

#[test]
fn lowercase_mixed_string() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(LowerCase('AbC123')); end."#),
        &["abc123"]
    );
}

#[test]
fn uppercase_mixed_string() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(UpperCase('xYz9')); end."#),
        &["XYZ9"]
    );
}

#[test]
fn char_range_vowels_only() {
    assert_eq!(
        run_pascal(
            r#"program T; var v:set of Char; begin v:=['a','e','i','o','u']; WriteLn(Ord('i' in v)); WriteLn(Ord('b' in v)); end."#
        ),
        &["1", "0"]
    );
}

#[test]
fn char_in_range_letters_loop_count() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:set of Char; c:Char; n:Integer; begin s:=['a'..'c']; n:=0; for c in s do Inc(n); WriteLn(n); end."#
        ),
        &["3"]
    );
}

#[test]
fn upcase_space_unchanged() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(UpCase(' ')); end."#),
        &[" "]
    );
}

#[test]
fn lowercase_space_unchanged() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(LowerCase(' ')); end."#),
        &[" "]
    );
}

#[test]
fn char_equality_case_sensitive() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Ord('a'='A')); end."#),
        &["0"]
    );
}

#[test]
fn char_inequality_case() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Ord('a'<>'A')); end."#),
        &["1"]
    );
}

#[test]
fn upcase_char_array_cell() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array[1..2] of Char; begin a[1]:='g'; a[2]:='h'; WriteLn(UpCase(a[1])); WriteLn(UpCase(a[2])); end."#
        ),
        &["G", "H"]
    );
}

#[test]
fn lowercase_char_from_string_index() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:string; c:Char; begin s:='Ab'; c:=s[1]; WriteLn(LowerCase(c)); end."#
        ),
        &["a"]
    );
}

#[test]
fn char_range_control_chars_low() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:set of Char; begin s:=['#0'..'#31']; WriteLn(Ord(#9 in s)); WriteLn(Ord('A' in s)); end."#
        ),
        &["1", "0"]
    );
}

#[test]
fn char_range_hex_letters() {
    assert_eq!(
        run_pascal(
            r#"program T; var h:set of Char; begin h:=['0'..'9','a'..'f']; WriteLn(Ord('c' in h)); WriteLn(Ord('g' in h)); end."#
        ),
        &["1", "0"]
    );
}

#[test]
fn upcase_string_empty() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Length(UpperCase(''))); end."#),
        &["0"]
    );
}

#[test]
fn lowercase_string_preserves_spaces() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(LowerCase('A B')); end."#),
        &["a b"]
    );
}

#[test]
fn char_succ_from_a_to_b() {
    assert_eq!(
        run_pascal(r#"program T; var c:Char; begin c:='a'; c:=Succ(c); WriteLn(c); end."#),
        &["b"]
    );
}

#[test]
fn char_pred_from_b_to_a() {
    assert_eq!(
        run_pascal(r#"program T; var c:Char; begin c:='b'; c:=Pred(c); WriteLn(c); end."#),
        &["a"]
    );
}

#[test]
fn upcase_in_case_statement() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; begin c:=UpCase('d'); case c of 'D': WriteLn('ok'); else WriteLn('no'); end; end."#
        ),
        &["ok"]
    );
}

#[test]
fn char_range_union_two_ranges() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:set of Char; begin s:=['a'..'c']+['x'..'z']; WriteLn(Ord('b' in s)); WriteLn(Ord('y' in s)); end."#
        ),
        &["1", "1"]
    );
}

#[test]
fn lowercase_then_compare_chars() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Char; begin a:=LowerCase('X'); b:='x'; WriteLn(Ord(a=b)); end."#
        ),
        &["1"]
    );
}

#[test]
fn upcase_y_is_y() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(UpCase('y')); end."#),
        &["Y"]
    );
}

#[test]
fn lowercase_z_is_z() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(LowerCase('z')); end."#),
        &["z"]
    );
}

#[test]
fn char_range_punctuation_subset() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:set of Char; begin p:=['.'..'/']; WriteLn(Ord('.' in p)); WriteLn(Ord('0' in p)); end."#
        ),
        &["1", "0"]
    );
}

#[test]
fn upcase_string_single_char() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(UpperCase('n')); end."#),
        &["N"]
    );
}

#[test]
fn char_is_letter_via_ranges() {
    assert_eq!(
        run_pascal(
            r#"program T; var letters:set of Char; c:Char; begin letters:=['A'..'Z']+['a'..'z']; c:='J'; WriteLn(Ord(c in letters)); c:='5'; WriteLn(Ord(c in letters)); end."#
        ),
        &["1", "0"]
    );
}
