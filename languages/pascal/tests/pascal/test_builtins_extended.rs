/// Extended builtin coverage: Inc/Dec deltas, Odd, Random, Length/High/Low variants.
use super::helpers::run_pascal;

#[test]
fn inc_with_delta_two() {
    assert_eq!(
        run_pascal("program T; var x: Integer; begin x := 10; Inc(x, 2); WriteLn(x); end."),
        &["12"]
    );
}

#[test]
fn inc_with_delta_five_from_zero() {
    assert_eq!(
        run_pascal("program T; var x: Integer; begin x := 0; Inc(x, 5); WriteLn(x); end."),
        &["5"]
    );
}

#[test]
fn dec_with_delta_three() {
    assert_eq!(
        run_pascal("program T; var x: Integer; begin x := 20; Dec(x, 3); WriteLn(x); end."),
        &["17"]
    );
}

#[test]
fn dec_with_delta_to_zero() {
    assert_eq!(
        run_pascal("program T; var x: Integer; begin x := 4; Dec(x, 4); WriteLn(x); end."),
        &["0"]
    );
}

#[test]
fn inc_delta_on_byte_counter() {
    assert_eq!(
        run_pascal("program T; var b: Byte; begin b := 250; Inc(b); WriteLn(b); end."),
        &["251"]
    );
}

#[test]
fn dec_delta_on_word_value() {
    assert_eq!(
        run_pascal("program T; var w: Word; begin w := 100; Dec(w, 25); WriteLn(w); end."),
        &["75"]
    );
}

#[test]
fn odd_on_three_returns_true() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Odd(3)); end."),
        &["true"]
    );
}

#[test]
fn odd_on_four_returns_false() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Odd(4)); end."),
        &["false"]
    );
}

#[test]
fn odd_on_negative_three() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Odd(-3)); end."),
        &["true"]
    );
}

#[test]
fn odd_on_negative_four() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Odd(-4)); end."),
        &["false"]
    );
}

#[test]
fn odd_on_large_even_int64() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Odd(1000002)); end."),
        &["false"]
    );
}

#[test]
fn odd_on_large_odd_int64() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Odd(1000001)); end."),
        &["true"]
    );
}

#[test]
fn random_after_randomize_is_nonnegative() {
    assert_eq!(
        run_pascal(
            r#"program T; var r: Double; begin Randomize; r := Random; WriteLn(r >= 0.0); end."#
        ),
        &["true"]
    );
}

#[test]
fn random_range_single_value() {
    assert_eq!(
        run_pascal(
            r#"program T; var n: Integer; begin Randomize; n := Random(1); WriteLn(n); end."#
        ),
        &["0"]
    );
}

#[test]
fn random_range_upper_bound_exclusive() {
    assert_eq!(
        run_pascal(
            r#"program T; var n: Integer; begin Randomize; n := Random(3); WriteLn((n >= 0) and (n < 3)); end."#
        ),
        &["true"]
    );
}

#[test]
fn random_range_hundred_values() {
    assert_eq!(
        run_pascal(
            r#"program T; var n: Integer; begin Randomize; n := Random(100); WriteLn(n < 100); end."#
        ),
        &["true"]
    );
}

#[test]
fn length_on_ansistring_variable() {
    assert_eq!(
        run_pascal("program T; var s: String; begin s := 'pascal'; WriteLn(Length(s)); end."),
        &["6"]
    );
}

#[test]
fn length_on_concatenated_string() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Length('ab' + 'cd' + 'ef')); end."),
        &["6"]
    );
}

#[test]
fn length_on_single_char_string() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Length('x')); end."),
        &["1"]
    );
}

#[test]
fn length_on_dynamic_array_after_grow() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array of Integer; begin SetLength(a, 7); WriteLn(Length(a)); end."#
        ),
        &["7"]
    );
}

#[test]
fn length_on_set_of_three_chars() {
    assert_eq!(
        run_pascal(
            r#"program T; var s: set of Char; begin s := ['a', 'b', 'c']; WriteLn(Length(s)); end."#
        ),
        &["3"]
    );
}

#[test]
fn length_on_empty_set() {
    assert_eq!(
        run_pascal(r#"program T; var s: set of Integer; begin WriteLn(Length(s)); end."#),
        &["0"]
    );
}

#[test]
fn high_on_static_one_based_array() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[1..5] of Integer; begin a[1] := 1; WriteLn(High(a)); end."#
        ),
        &["5"]
    );
}

#[test]
fn low_on_static_one_based_array() {
    assert_eq!(
        run_pascal(r#"program T; var a: array[1..5] of Integer; begin WriteLn(Low(a)); end."#),
        &["1"]
    );
}

#[test]
fn high_low_on_negative_index_range() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[-2..2] of Integer; begin WriteLn(Low(a)); WriteLn(High(a)); end."#
        ),
        &["-2", "2"]
    );
}

#[test]
fn high_on_string_is_last_index() {
    assert_eq!(
        run_pascal(r#"program T; var s: String; begin s := 'abc'; WriteLn(High(s)); end."#),
        &["3"]
    );
}

#[test]
fn low_on_string_is_one() {
    assert_eq!(
        run_pascal(r#"program T; var s: String; begin s := 'xy'; WriteLn(Low(s)); end."#),
        &["1"]
    );
}

#[test]
fn succ_on_char_advances() {
    assert_eq!(
        run_pascal(r#"program T; var c: Char; begin c := 'A'; c := Succ(c); WriteLn(c); end."#),
        &["B"]
    );
}

#[test]
fn pred_on_char_retreats() {
    assert_eq!(
        run_pascal(r#"program T; var c: Char; begin c := 'C'; c := Pred(c); WriteLn(c); end."#),
        &["B"]
    );
}

#[test]
fn succ_on_enum_member() {
    assert_eq!(
        run_pascal(
            r#"program T; type T = (Red, Green, Blue); var c: T; begin c := Red; c := Succ(c); WriteLn(Ord(c)); end."#
        ),
        &["1"]
    );
}

#[test]
fn pred_on_enum_member() {
    assert_eq!(
        run_pascal(
            r#"program T; type T = (Red, Green, Blue); var c: T; begin c := Blue; c := Pred(c); WriteLn(Ord(c)); end."#
        ),
        &["1"]
    );
}

#[test]
fn inc_on_for_loop_counter() {
    assert_eq!(
        run_pascal(
            r#"program T; var i, sum: Integer; begin sum := 0; for i := 1 to 3 do begin Inc(sum, i); end; WriteLn(sum); end."#
        ),
        &["6"]
    );
}

#[test]
fn dec_in_while_loop_countdown() {
    assert_eq!(
        run_pascal(
            r#"program T; var n: Integer; begin n := 3; while n > 0 do begin WriteLn(n); Dec(n); end; end."#
        ),
        &["3", "2", "1"]
    );
}

#[test]
fn assigned_on_new_pointer_before_dispose() {
    assert_eq!(
        run_pascal(
            r#"program T; var p: ^Integer; begin New(p); WriteLn(Assigned(p)); Dispose(p); end."#
        ),
        &["true"]
    );
}

#[test]
fn paramstr_one_when_missing_is_empty() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Length(ParamStr(1)) = 0); end."#),
        &["true"]
    );
}

#[test]
fn upcase_on_digit_unchanged() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(UpCase('5')); end."#),
        &["5"]
    );
}

#[test]
fn lo_case_on_digit_unchanged() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(LoCase('9')); end."#),
        &["9"]
    );
}

#[test]
fn chr_ord_roundtrip_via_builtin() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Chr(Ord('M'))); end."#),
        &["M"]
    );
}

#[test]
fn min_of_three_via_nested_min() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Min(Min(9, 4), Min(7, 2))); end."#),
        &["2"]
    );
}

#[test]
fn max_of_three_via_nested_max() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Max(Max(1, 8), Max(3, 5))); end."#),
        &["8"]
    );
}
