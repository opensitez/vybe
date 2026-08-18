/// Typed constants, constant expressions, and resource-string style patterns.
use super::helpers::run_pascal;

#[test]
fn const_hex_integer_literal() {
    assert_eq!(
        run_pascal(r#"program T; const N = $FF; begin WriteLn(N); end."#),
        &["255"]
    );
}

#[test]
fn const_binary_integer_literal() {
    assert_eq!(
        run_pascal(r#"program T; const N = %1010; begin WriteLn(N); end."#),
        &["10"]
    );
}

#[test]
fn const_octal_style_via_expression() {
    assert_eq!(
        run_pascal(r#"program T; const N = 8 * 8 + 1; begin WriteLn(N); end."#),
        &["65"]
    );
}

#[test]
fn const_nested_arithmetic_chain() {
    assert_eq!(
        run_pascal(
            r#"program T; const A = 2; const B = 3; const C = (A + B) * 4; begin WriteLn(C); end."#
        ),
        &["20"]
    );
}

#[test]
fn const_string_multiline_concat() {
    assert_eq!(
        run_pascal(
            r#"program T; const P1 = 'hel'; const P2 = 'lo'; const S = P1 + P2; begin WriteLn(S); end."#
        ),
        &["hello"]
    );
}

#[test]
fn const_typed_byte_value() {
    assert_eq!(
        run_pascal(r#"program T; const B: Byte = 200; begin WriteLn(B); end."#),
        &["200"]
    );
}

#[test]
fn const_typed_word_value() {
    assert_eq!(
        run_pascal(r#"program T; const W: Word = 1000; begin WriteLn(W); end."#),
        &["1000"]
    );
}

#[test]
fn const_negative_integer() {
    assert_eq!(
        run_pascal(r#"program T; const N: Integer = -17; begin WriteLn(N); end."#),
        &["-17"]
    );
}

#[test]
fn const_real_division_expression() {
    assert_eq!(
        run_pascal(r#"program T; const R: Double = 10 / 4; begin WriteLn(R); end."#),
        &["2.5"]
    );
}

#[test]
fn const_array_three_elements() {
    assert_eq!(
        run_pascal(
            r#"program T; const A: array[0..2] of Integer = (1, 2, 3); begin WriteLn(A[2]); end."#
        ),
        &["3"]
    );
}

#[test]
fn const_array_string_elements() {
    assert_eq!(
        run_pascal(
            r#"program T; const Names: array[0..1] of string = ('a', 'b'); begin WriteLn(Names[1]); end."#
        ),
        &["b"]
    );
}

#[test]
fn const_record_two_fields() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR = record X, Y: Integer; end; const P: TR = (X: 3; Y: 4); begin WriteLn(P.X + P.Y); end."#
        ),
        &["7"]
    );
}

#[test]
fn const_enum_value_reference() {
    assert_eq!(
        run_pascal(
            r#"program T; type TLevel = (Low, Mid, High); const L: TLevel = High; begin WriteLn(Ord(L)); end."#
        ),
        &["2"]
    );
}

#[test]
fn const_set_union_expression() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD = (A, B, C); const S1 = [A]; const S2 = [B, C]; const S = S1 + S2; var x: TD; begin x := B; if x in S then WriteLn('in') else WriteLn('out'); end."#
        ),
        &["in"]
    );
}

#[test]
fn const_char_sequence_builds_marker() {
    assert_eq!(
        run_pascal(
            r#"program T; const C1 = '>'; const C2 = '>'; const Marker = C1 + C2; begin WriteLn(Marker); end."#
        ),
        &[">>"]
    );
}

#[test]
fn const_boolean_and_expression() {
    assert_eq!(
        run_pascal(
            r#"program T; const T1 = True; const T2 = True; const R = T1 and T2; begin if R then WriteLn('ok'); end."#
        ),
        &["ok"]
    );
}

#[test]
fn const_boolean_or_expression() {
    assert_eq!(
        run_pascal(
            r#"program T; const F = False; const T1 = True; const R = F or T1; begin if R then WriteLn('ok'); end."#
        ),
        &["ok"]
    );
}

#[test]
fn const_in_case_label() {
    assert_eq!(
        run_pascal(
            r#"program T; const CODE = 2; var x: Integer; begin x := CODE; case x of 1: WriteLn('one'); 2: WriteLn('two'); else WriteLn('other'); end; end."#
        ),
        &["two"]
    );
}

#[test]
fn const_used_in_for_bound() {
    assert_eq!(
        run_pascal(
            r#"program T; const LAST = 3; var i, s: Integer; begin s := 0; for i := 1 to LAST do s := s + i; WriteLn(s); end."#
        ),
        &["6"]
    );
}

#[test]
fn const_pointer_nil_compare() {
    assert_eq!(
        run_pascal(
            r#"program T; const P: Pointer = nil; begin if P = nil then WriteLn('nil'); end."#
        ),
        &["nil"]
    );
}

#[test]
fn const_subrange_upper_bound() {
    assert_eq!(
        run_pascal(
            r#"program T; type TRange = 1..5; const MAX: TRange = 5; begin WriteLn(MAX); end."#
        ),
        &["5"]
    );
}

#[test]
fn resourcestring_style_const_string() {
    assert_eq!(
        run_pascal(
            r#"program T; const SMsgNotFound = 'Record not found'; begin WriteLn(SMsgNotFound); end."#
        ),
        &["Record not found"]
    );
}

#[test]
fn resourcestring_style_error_template() {
    assert_eq!(
        run_pascal(
            r#"program T; const SErrPrefix = 'Error: '; const SErrCode = 'E001'; const SFull = SErrPrefix + SErrCode; begin WriteLn(SFull); end."#
        ),
        &["Error: E001"]
    );
}

#[test]
fn resourcestring_style_plural_label() {
    assert_eq!(
        run_pascal(
            r#"program T; const SItemSingular = 'item'; const SItemPlural = 'items'; const UsePlural = True; begin if UsePlural then WriteLn(SItemPlural) else WriteLn(SItemSingular); end."#
        ),
        &["items"]
    );
}

#[test]
fn resourcestring_style_button_caption() {
    assert_eq!(
        run_pascal(
            r#"program T; const SCaptionOk = 'OK'; const SCaptionCancel = 'Cancel'; begin WriteLn(SCaptionOk); WriteLn(SCaptionCancel); end."#
        ),
        &["OK", "Cancel"]
    );
}

#[test]
fn const_array_index_from_const() {
    assert_eq!(
        run_pascal(
            r#"program T; const IDX = 1; const A: array[0..2] of Integer = (10, 20, 30); begin WriteLn(A[IDX]); end."#
        ),
        &["20"]
    );
}

#[test]
fn const_mod_expression() {
    assert_eq!(
        run_pascal(r#"program T; const N = 17 mod 5; begin WriteLn(N); end."#),
        &["2"]
    );
}

#[test]
fn const_div_expression() {
    assert_eq!(
        run_pascal(r#"program T; const N = 17 div 5; begin WriteLn(N); end."#),
        &["3"]
    );
}

#[test]
fn const_shl_expression() {
    assert_eq!(
        run_pascal(r#"program T; const N = 1 shl 4; begin WriteLn(N); end."#),
        &["16"]
    );
}

#[test]
fn const_shr_expression() {
    assert_eq!(
        run_pascal(r#"program T; const N = 32 shr 2; begin WriteLn(N); end."#),
        &["8"]
    );
}

#[test]
fn const_record_nested_field_init() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInner = record V: Integer; end; type TOuter = record Inner: TInner; end; const O: TOuter = (Inner: (V: 9)); begin WriteLn(O.Inner.V); end."#
        ),
        &["9"]
    );
}

#[test]
fn const_multiple_in_same_const_section() {
    assert_eq!(
        run_pascal(r#"program T; const A = 1; B = 2; C = 3; begin WriteLn(A + B + C); end."#),
        &["6"]
    );
}

#[test]
fn const_string_of_char_repeat_style() {
    assert_eq!(
        run_pascal(
            r#"program T; const DASH = '-'; const LINE = DASH + DASH + DASH; begin WriteLn(LINE); end."#
        ),
        &["---"]
    );
}

#[test]
fn const_compare_in_if() {
    assert_eq!(
        run_pascal(
            r#"program T; const THRESH = 10; var n: Integer; begin n := 12; if n > THRESH then WriteLn('above'); end."#
        ),
        &["above"]
    );
}

#[test]
fn const_real_pi_approximation() {
    assert_eq!(
        run_pascal(r#"program T; const PI = 3.14159; begin WriteLn(PI > 3); end."#),
        &["TRUE"]
    );
}

#[test]
fn const_typed_ansistring() {
    assert_eq!(
        run_pascal(r#"program T; const S: AnsiString = 'ansi'; begin WriteLn(S); end."#),
        &["ansi"]
    );
}

#[test]
fn const_array_bounds_low_high() {
    assert_eq!(
        run_pascal(
            r#"program T; const A: array[5..7] of Integer = (1, 2, 3); begin WriteLn(Low(A)); WriteLn(High(A)); end."#
        ),
        &["5", "7"]
    );
}

#[test]
fn const_expression_with_ord() {
    assert_eq!(
        run_pascal(r#"program T; type T = (X, Y, Z); const N = Ord(Z); begin WriteLn(N); end."#),
        &["2"]
    );
}

#[test]
fn const_case_insensitive_string_compare_const() {
    assert_eq!(
        run_pascal(r#"program T; const GREET = 'Hello'; begin WriteLn(UpperCase(GREET)); end."#),
        &["HELLO"]
    );
}

#[test]
fn const_in_record_default_parameter() {
    assert_eq!(
        run_pascal(
            r#"program T; const DEFAULT_N = 5; procedure Show(n: Integer); begin WriteLn(n); end; begin Show(DEFAULT_N); end."#
        ),
        &["5"]
    );
}

#[test]
fn const_resource_lookup_table_entry() {
    assert_eq!(
        run_pascal(
            r#"program T; const Keys: array[0..2] of string = ('red', 'green', 'blue'); const Vals: array[0..2] of Integer = (1, 2, 3); var i: Integer; begin for i := 0 to 2 do if Keys[i] = 'green' then WriteLn(Vals[i]); end."#
        ),
        &["2"]
    );
}
