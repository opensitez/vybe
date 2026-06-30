/// Type aliases, subranges, and ordinal type operations — extended coverage.
use super::helpers::run_pascal;

#[test]
fn alias_integer_to_count_type() {
    assert_eq!(
        run_pascal(
            r#"program T; type TCount = Integer; var n: TCount; begin n := 7; WriteLn(n); end."#
        ),
        &["7"]
    );
}

#[test]
fn alias_string_to_name_type() {
    assert_eq!(
        run_pascal(
            r#"program T; type TName = string; var s: TName; begin s := 'bob'; WriteLn(s); end."#
        ),
        &["bob"]
    );
}

#[test]
fn alias_nested_type_reference() {
    assert_eq!(
        run_pascal(
            r#"program T; type TId = Integer; type TUser = record Id: TId; end; var u: TUser; begin u.Id := 42; WriteLn(u.Id); end."#
        ),
        &["42"]
    );
}

#[test]
fn subrange_byte_values() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPercent = 0..100; var p: TPercent; begin p := 75; WriteLn(p); end."#
        ),
        &["75"]
    );
}

#[test]
fn subrange_char_letters() {
    assert_eq!(
        run_pascal(
            r#"program T; type TLetter = 'a'..'z'; var c: TLetter; begin c := 'm'; WriteLn(c); end."#
        ),
        &["m"]
    );
}

#[test]
fn subrange_digit_chars() {
    assert_eq!(
        run_pascal(
            r#"program T; type TDigit = '0'..'9'; var d: TDigit; begin d := '5'; WriteLn(d); end."#
        ),
        &["5"]
    );
}

#[test]
fn subrange_negative_small() {
    assert_eq!(
        run_pascal(
            r#"program T; type TSmall = -5..5; var x: TSmall; begin x := -3; WriteLn(x); end."#
        ),
        &["-3"]
    );
}

#[test]
fn subrange_in_for_loop() {
    assert_eq!(
        run_pascal(
            r#"program T; type TIdx = 1..3; var i: TIdx; begin for i := 1 to 3 do WriteLn(i); end."#
        ),
        &["1", "2", "3"]
    );
}

#[test]
fn ordinal_succ_on_subrange() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR = 10..20; var x: TR; begin x := 15; WriteLn(Succ(x)); end."#
        ),
        &["16"]
    );
}

#[test]
fn ordinal_pred_on_subrange() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR = 10..20; var x: TR; begin x := 15; WriteLn(Pred(x)); end."#
        ),
        &["14"]
    );
}

#[test]
fn ordinal_ord_on_char_subrange() {
    assert_eq!(
        run_pascal(
            r#"program T; type TLetter = 'A'..'Z'; var c: TLetter; begin c := 'C'; WriteLn(Ord(c)); end."#
        ),
        &["67"]
    );
}

#[test]
fn ordinal_chr_from_integer() {
    assert_eq!(
        run_pascal(
            r#"program T; var c: Char; begin c := Chr(65); WriteLn(c); end."#
        ),
        &["A"]
    );
}

#[test]
fn type_alias_array_of_integer() {
    assert_eq!(
        run_pascal(
            r#"program T; type TIntArray = array[0..2] of Integer; var a: TIntArray; begin a[1] := 9; WriteLn(a[1]); end."#
        ),
        &["9"]
    );
}

#[test]
fn type_alias_pointer_type() {
    assert_eq!(
        run_pascal(
            r#"program T; type PInt = ^Integer; var n: Integer; p: PInt; begin n := 11; p := @n; WriteLn(p^); end."#
        ),
        &["11"]
    );
}

#[test]
fn subrange_assignment_from_literal() {
    assert_eq!(
        run_pascal(
            r#"program T; type TMonth = 1..12; var m: TMonth; begin m := 12; WriteLn(m); end."#
        ),
        &["12"]
    );
}

#[test]
fn enum_as_ordinal_type() {
    assert_eq!(
        run_pascal(
            r#"program T; type TDir = (North, East, South, West); var d: TDir; begin d := South; WriteLn(Ord(d)); end."#
        ),
        &["2"]
    );
}

#[test]
fn enum_succ_wraps_forward() {
    assert_eq!(
        run_pascal(
            r#"program T; type TDay = (Mon, Tue, Wed); var d: TDay; begin d := Mon; WriteLn(Ord(Succ(d))); end."#
        ),
        &["1"]
    );
}

#[test]
fn enum_pred_steps_back() {
    assert_eq!(
        run_pascal(
            r#"program T; type TDay = (Mon, Tue, Wed); var d: TDay; begin d := Wed; WriteLn(Ord(Pred(d))); end."#
        ),
        &["1"]
    );
}

#[test]
fn set_of_subrange_type() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD = 1..3; var s: set of TD; x: TD; begin s := [1, 3]; x := 2; if x in s then WriteLn('in') else WriteLn('out'); end."#
        ),
        &["out"]
    );
}

#[test]
fn set_of_enum_membership() {
    assert_eq!(
        run_pascal(
            r#"program T; type TC = (Red, Green, Blue); var s: set of TC; c: TC; begin s := [Red, Blue]; c := Blue; if c in s then WriteLn('yes'); end."#
        ),
        &["yes"]
    );
}

#[test]
fn alias_record_type_name() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPoint = record X, Y: Integer; end; type TLocation = TPoint; var p: TLocation; begin p.X := 4; p.Y := 5; WriteLn(p.X + p.Y); end."#
        ),
        &["9"]
    );
}

#[test]
fn subrange_compare_less_than() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR = 0..9; var a, b: TR; begin a := 3; b := 7; if a < b then WriteLn('lt'); end."#
        ),
        &["lt"]
    );
}

#[test]
fn byte_type_range_value() {
    assert_eq!(
        run_pascal(
            r#"program T; var b: Byte; begin b := 255; WriteLn(b); end."#
        ),
        &["255"]
    );
}

#[test]
fn shortint_negative_value() {
    assert_eq!(
        run_pascal(
            r#"program T; var s: ShortInt; begin s := -128; WriteLn(s); end."#
        ),
        &["-128"]
    );
}

#[test]
fn cardinal_large_unsigned() {
    assert_eq!(
        run_pascal(
            r#"program T; var c: Cardinal; begin c := 4000000000; WriteLn(c > 1000000000); end."#
        ),
        &["True"]
    );
}

#[test]
fn int64_literal_value() {
    assert_eq!(
        run_pascal(
            r#"program T; var i: Int64; begin i := 10000000000; WriteLn(i div 1000000000); end."#
        ),
        &["10"]
    );
}

#[test]
fn single_float_type() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: Single; begin f := 2.5; WriteLn(f + f); end."#
        ),
        &["5"]
    );
}

#[test]
fn extended_float_type() {
    assert_eq!(
        run_pascal(
            r#"program T; var e: Extended; begin e := 1.25; WriteLn(e * 4); end."#
        ),
        &["5"]
    );
}

#[test]
fn currency_type_roundtrip() {
    assert_eq!(
        run_pascal(
            r#"program T; var c: Currency; begin c := 19.99; WriteLn(c > 19); end."#
        ),
        &["True"]
    );
}

#[test]
fn boolean_ord_values() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(Ord(False)); WriteLn(Ord(True)); end."#
        ),
        &["0", "1"]
    );
}

#[test]
fn char_succ_advances_ascii() {
    assert_eq!(
        run_pascal(
            r#"program T; var c: Char; begin c := 'A'; WriteLn(Succ(c)); end."#
        ),
        &["B"]
    );
}

#[test]
fn char_pred_steps_back_ascii() {
    assert_eq!(
        run_pascal(
            r#"program T; var c: Char; begin c := 'B'; WriteLn(Pred(c)); end."#
        ),
        &["A"]
    );
}

#[test]
fn typecast_integer_to_enum() {
    assert_eq!(
        run_pascal(
            r#"program T; type T = (A, B, C); var e: T; begin e := T(2); WriteLn(Ord(e)); end."#
        ),
        &["2"]
    );
}

#[test]
fn alias_procedure_type() {
    assert_eq!(
        run_pascal(
            r#"program T; type TProc = procedure; procedure P; begin WriteLn('run'); end; var fp: TProc; begin fp := @P; fp(); end."#
        ),
        &["run"]
    );
}

#[test]
fn alias_function_type() {
    assert_eq!(
        run_pascal(
            r#"program T; type TFunc = function(x: Integer): Integer; function Sq(x: Integer): Integer; begin Result := x * x; end; var f: TFunc; begin f := @Sq; WriteLn(f(6)); end."#
        ),
        &["36"]
    );
}

#[test]
fn subrange_array_index_type() {
    assert_eq!(
        run_pascal(
            r#"program T; type TIdx = 1..3; var a: array[TIdx] of Integer; i: TIdx; begin for i := 1 to 3 do a[i] := i; WriteLn(a[2]); end."#
        ),
        &["2"]
    );
}

#[test]
fn enum_explicit_value_gap() {
    assert_eq!(
        run_pascal(
            r#"program T; type T = (A = 10, B = 20, C = 30); var x: T; begin x := B; WriteLn(x); end."#
        ),
        &["20"]
    );
}

#[test]
fn set_union_two_enums() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD = (A, B, C); var s1, s2, s3: set of TD; begin s1 := [A]; s2 := [B, C]; s3 := s1 + s2; if A in s3 then WriteLn('has'); end."#
        ),
        &["has"]
    );
}

#[test]
fn set_difference_operation() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD = (A, B, C); var s1, s2, s3: set of TD; begin s1 := [A, B, C]; s2 := [B]; s3 := s1 - s2; if A in s3 then WriteLn('a'); if B in s3 then WriteLn('b'); end."#
        ),
        &["a"]
    );
}

#[test]
fn set_intersection_operation() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD = (A, B, C); var s1, s2, s3: set of TD; begin s1 := [A, B]; s2 := [B, C]; s3 := s1 * s2; if B in s3 then WriteLn('b'); if A in s3 then WriteLn('a'); end."#
        ),
        &["b", "a"]
    );
}

#[test]
fn ordinal_low_high_on_static_array() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[2..5] of Integer; begin WriteLn(Low(a)); WriteLn(High(a)); end."#
        ),
        &["2", "5"]
    );
}
