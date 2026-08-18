/// Literal syntax: numbers, chars, strings, sets, arrays, booleans, hex/binary.
use super::helpers::run_pascal;

#[test]
fn hex_integer_literal() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn($FF); end."#),
        &["255"]
    );
}

#[test]
fn binary_integer_literal() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(%1010); end."#),
        &["10"]
    );
}

#[test]
fn octal_integer_literal() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(&17); end."#),
        &["15"]
    );
}

#[test]
fn scientific_notation_real() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Round(1.5e2)); end."#),
        &["150"]
    );
}

#[test]
fn char_control_literal_tab() {
    assert_eq!(
        run_pascal(r#"program T; var c:Char; begin c:=#9; WriteLn(Ord(c)); end."#),
        &["9"]
    );
}

#[test]
fn string_escaped_quote() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn('it''s'); end."#),
        &["it's"]
    );
}

#[test]
fn string_hash_literal_style() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(#65#66); end."#),
        &["AB"]
    );
}

#[test]
fn set_literal_empty() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:set of Integer; begin s:=[]; if s=[] then WriteLn('empty'); end."#
        ),
        &["empty"]
    );
}

#[test]
fn set_literal_range_members() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:set of Integer; n:Integer; c:Integer; begin s:=[1..3]; c:=0; for n in s do Inc(c); WriteLn(c); end."#
        ),
        &["3"]
    );
}

#[test]
fn array_literal_integer() {
    assert_eq!(
        run_pascal(r#"program T; var a:array of Integer; begin a:=[10,20]; WriteLn(a[1]); end."#),
        &["20"]
    );
}

#[test]
fn array_literal_string() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array of string; begin a:=['a','b']; WriteLn(a[0]+a[1]); end."#
        ),
        &["ab"]
    );
}

#[test]
fn boolean_literals_in_expression() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(true and false); WriteLn(true or false); end."#),
        &["FALSE", "TRUE"]
    );
}

#[test]
fn negative_integer_literal() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(-42); end."#),
        &["-42"]
    );
}

#[test]
fn underscore_in_numeric_literal() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(1_000); end."#),
        &["1000"]
    );
}

#[test]
fn wide_char_literal_if_supported() {
    assert_eq!(run_pascal(r#"program T; begin WriteLn('Ω'); end."#), &["Ω"]);
}

#[test]
fn multiline_string_concat() {
    assert_eq!(
        run_pascal(r#"program T; var s:string; begin s:='line1'+#10+'line2'; WriteLn(s); end."#),
        &["line1", "line2"]
    );
}

#[test]
fn typed_constant_literal_expression() {
    assert_eq!(
        run_pascal(r#"program T; const N:Integer=2*3; begin WriteLn(N); end."#),
        &["6"]
    );
}

#[test]
fn real_literal_fraction() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(0.125*8); end."#),
        &["1"]
    );
}

#[test]
fn char_literal_vs_ord() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Ord('0')); end."#),
        &["48"]
    );
}

#[test]
fn set_char_literals() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:set of Char; begin s:=['a'..'c']; if 'b' in s then WriteLn('in'); end."#
        ),
        &["in"]
    );
}
