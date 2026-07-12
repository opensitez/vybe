/// Typed constants and constant expressions in Object Pascal.
use super::helpers::run_pascal;

#[test]
fn typed_const_integer_preserves_type() {
    assert_eq!(
        run_pascal(r#"program T; const N: Integer = 42; begin WriteLn(N); end."#),
        &["42"]
    );
}

#[test]
fn typed_const_string_concat_in_const_expr() {
    assert_eq!(
        run_pascal(
            r#"program T; const Prefix = 'id-'; const S: string = Prefix + '7'; begin WriteLn(S); end."#
        ),
        &["id-7"]
    );
}

#[test]
fn typed_const_real_value() {
    assert_eq!(
        run_pascal(r#"program T; const PiApprox: Double = 3.25; begin WriteLn(PiApprox); end."#),
        &["3.25"]
    );
}

#[test]
fn typed_const_char_literal() {
    assert_eq!(
        run_pascal(r#"program T; const Sep: Char = '|'; begin WriteLn(Sep); end."#),
        &["|"]
    );
}

#[test]
fn typed_const_boolean_in_if() {
    assert_eq!(
        run_pascal(
            r#"program T; const Flag: Boolean = True; begin if Flag then WriteLn('yes'); end."#
        ),
        &["yes"]
    );
}

#[test]
fn typed_const_set_literal() {
    assert_eq!(
        run_pascal(
            r#"program T; type TD = (One, Two, Three); const S: set of TD = [One, Three]; var x: TD; begin x := Two; if not (x in S) then WriteLn('missing'); end."#
        ),
        &["missing"]
    );
}

#[test]
fn typed_const_array_of_integer() {
    assert_eq!(
        run_pascal(
            r#"program T; const Digits: array[0..2] of Integer = (2, 4, 6); begin WriteLn(Digits[1]); end."#
        ),
        &["4"]
    );
}

#[test]
fn typed_const_record_aggregate() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR = record X, Y: Integer; end; const Origin: TR = (X: 0; Y: 0); var r: TR; begin r := Origin; WriteLn(r.X); WriteLn(r.Y); end."#
        ),
        &["0", "0"]
    );
}

#[test]
fn const_expression_arithmetic_at_compile_time() {
    assert_eq!(
        run_pascal(
            r#"program T; const A = 6; const B = 7; const C = A * B; begin WriteLn(C); end."#
        ),
        &["42"]
    );
}
