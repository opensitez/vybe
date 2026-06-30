/// Case labels with ranges and comma-separated lists — not covered in test_control_flow.rs.
use super::helpers::run_pascal;

#[test]
fn case_integer_range_matches_inside() {
    assert_eq!(
        run_pascal(
            r#"program T; var n: Integer; begin n := 3; case n of 1..5: WriteLn('band'); else WriteLn('other'); end; end."#
        ),
        &["band"]
    );
}

#[test]
fn case_integer_range_falls_to_else() {
    assert_eq!(
        run_pascal(
            r#"program T; var n: Integer; begin n := 9; case n of 1..5: WriteLn('band'); else WriteLn('other'); end; end."#
        ),
        &["other"]
    );
}

#[test]
fn case_negative_integer_range() {
    assert_eq!(
        run_pascal(
            r#"program T; var n: Integer; begin n := -2; case n of -5..-1: WriteLn('neg'); else WriteLn('other'); end; end."#
        ),
        &["neg"]
    );
}

#[test]
fn case_char_range_lowercase() {
    assert_eq!(
        run_pascal(
            r#"program T; var c: Char; begin c := 'm'; case c of 'a'..'m': WriteLn('first'); 'n'..'z': WriteLn('second'); end; end."#
        ),
        &["first"]
    );
}

#[test]
fn case_char_ascii_digit_range() {
    assert_eq!(
        run_pascal(
            r#"program T; var c: Char; begin c := '7'; case c of '0'..'9': WriteLn('digit'); else WriteLn('other'); end; end."#
        ),
        &["digit"]
    );
}

#[test]
fn case_comma_separated_disjoint_ranges() {
    assert_eq!(
        run_pascal(
            r#"program T; var n: Integer; begin n := 10; case n of 1..3, 8..12: WriteLn('hit'); else WriteLn('miss'); end; end."#
        ),
        &["hit"]
    );
}

#[test]
fn case_comma_separated_single_values() {
    assert_eq!(
        run_pascal(
            r#"program T; var d: Integer; begin d := 6; case d of 1, 3, 5, 7: WriteLn('odd'); 2, 4, 6, 8: WriteLn('even'); end; end."#
        ),
        &["even"]
    );
}

#[test]
fn case_enum_members_as_labels() {
    assert_eq!(
        run_pascal(
            r#"program T; type TColor = (Red, Green, Blue); var c: TColor; begin c := Green; case c of Red: WriteLn('r'); Green: WriteLn('g'); Blue: WriteLn('b'); end; end."#
        ),
        &["g"]
    );
}

#[test]
fn case_enum_subset_weekend() {
    assert_eq!(
        run_pascal(
            r#"program T; type TDay = (Mon, Tue, Wed, Thu, Fri, Sat, Sun); var d: TDay; begin d := Sun; case d of Sat, Sun: WriteLn('weekend'); else WriteLn('weekday'); end; end."#
        ),
        &["weekend"]
    );
}

#[test]
fn case_function_result_from_range_branch() {
    assert_eq!(
        run_pascal(
            r#"program T; function Sign(n: Integer): string; begin case n of -5..-1: Result := 'neg'; 0: Result := 'zero'; 1..5: Result := 'pos'; else Result := 'unknown'; end; end; begin WriteLn(Sign(-3)); WriteLn(Sign(0)); WriteLn(Sign(4)); end."#
        ),
        &["neg", "zero", "pos"]
    );
}

#[test]
fn case_nested_inner_range() {
    assert_eq!(
        run_pascal(
            r#"program T; var a, b: Integer; begin a := 1; b := 4; case a of 1: case b of 2..5: WriteLn('inner'); else WriteLn('outer'); end; else WriteLn('skip'); end; end."#
        ),
        &["inner"]
    );
}

#[test]
fn case_hex_nibble_char_range() {
    assert_eq!(
        run_pascal(
            r#"program T; var c: Char; v: Integer; begin c := 'C'; case c of '0'..'9': v := Ord(c) - Ord('0'); 'A'..'F': v := 10 + Ord(c) - Ord('A'); else v := -1; end; WriteLn(v); end."#
        ),
        &["12"]
    );
}
