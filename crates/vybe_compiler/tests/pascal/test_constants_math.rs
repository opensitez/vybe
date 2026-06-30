/// Constant expressions and compile-time evaluation patterns.
use super::helpers::run_pascal;

#[test]
fn const_addition() {
    assert_eq!(
        run_pascal(
            r#"program T; const N=2+3; begin WriteLn(N); end."#
        ),
        &["5"]
    );
}

#[test]
fn const_subtraction() {
    assert_eq!(
        run_pascal(
            r#"program T; const N=10-4; begin WriteLn(N); end."#
        ),
        &["6"]
    );
}

#[test]
fn const_multiplication() {
    assert_eq!(
        run_pascal(
            r#"program T; const N=6*7; begin WriteLn(N); end."#
        ),
        &["42"]
    );
}

#[test]
fn const_division() {
    assert_eq!(
        run_pascal(
            r#"program T; const N=15 div 4; begin WriteLn(N); end."#
        ),
        &["3"]
    );
}

#[test]
fn const_modulo() {
    assert_eq!(
        run_pascal(
            r#"program T; const N=17 mod 5; begin WriteLn(N); end."#
        ),
        &["2"]
    );
}

#[test]
fn const_hex_ff() {
    assert_eq!(
        run_pascal(
            r#"program T; const N=$FF; begin WriteLn(N); end."#
        ),
        &["255"]
    );
}

#[test]
fn const_binary_1010() {
    assert_eq!(
        run_pascal(
            r#"program T; const N=%1010; begin WriteLn(N); end."#
        ),
        &["10"]
    );
}

#[test]
fn const_nested_parens() {
    assert_eq!(
        run_pascal(
            r#"program T; const N=(2+3)*4; begin WriteLn(N); end."#
        ),
        &["20"]
    );
}

#[test]
fn const_chain_refs() {
    assert_eq!(
        run_pascal(
            r#"program T; const A=2; const B=A+3; const C=B*2; begin WriteLn(C); end."#
        ),
        &["10"]
    );
}

#[test]
fn const_negative() {
    assert_eq!(
        run_pascal(
            r#"program T; const N=-(3+4); begin WriteLn(N); end."#
        ),
        &["-7"]
    );
}

#[test]
fn const_string_concat() {
    assert_eq!(
        run_pascal(
            r#"program T; const S='ab'+'cd'; begin WriteLn(S); end."#
        ),
        &["abcd"]
    );
}

#[test]
fn const_typed_byte() {
    assert_eq!(
        run_pascal(
            r#"program T; const B:Byte=200; begin WriteLn(B); end."#
        ),
        &["200"]
    );
}

#[test]
fn const_typed_integer() {
    assert_eq!(
        run_pascal(
            r#"program T; const I:Integer=-42; begin WriteLn(I); end."#
        ),
        &["-42"]
    );
}

#[test]
fn const_real_pi_approx() {
    assert_eq!(
        run_pascal(
            r#"program T; const R:Real=3.0+0.14; begin WriteLn(Trunc(R*100)); end."#
        ),
        &["314"]
    );
}

#[test]
fn const_bool_and() {
    assert_eq!(
        run_pascal(
            r#"program T; const T=True and False; begin WriteLn(T); end."#
        ),
        &["false"]
    );
}

#[test]
fn const_bool_or() {
    assert_eq!(
        run_pascal(
            r#"program T; const T=False or True; begin WriteLn(T); end."#
        ),
        &["true"]
    );
}

#[test]
fn const_bool_not() {
    assert_eq!(
        run_pascal(
            r#"program T; const T=not False; begin WriteLn(T); end."#
        ),
        &["true"]
    );
}

#[test]
fn const_compare_eq() {
    assert_eq!(
        run_pascal(
            r#"program T; const T=(5=5); begin WriteLn(T); end."#
        ),
        &["true"]
    );
}

#[test]
fn const_compare_ne() {
    assert_eq!(
        run_pascal(
            r#"program T; const T=(5<>3); begin WriteLn(T); end."#
        ),
        &["true"]
    );
}

#[test]
fn const_power_style() {
    assert_eq!(
        run_pascal(
            r#"program T; const N=2*2*2*2; begin WriteLn(N); end."#
        ),
        &["16"]
    );
}

#[test]
fn const_ord_char() {
    assert_eq!(
        run_pascal(
            r#"program T; const N=Ord('A'); begin WriteLn(N); end."#
        ),
        &["65"]
    );
}

#[test]
fn const_length_string() {
    assert_eq!(
        run_pascal(
            r#"program T; const S='hello'; const N=Length(S); begin WriteLn(N); end."#
        ),
        &["5"]
    );
}

#[test]
fn const_array_index_expr() {
    assert_eq!(
        run_pascal(
            r#"program T; const I=1+2; const V=10*I; begin WriteLn(V); end."#
        ),
        &["30"]
    );
}

#[test]
fn const_multiple_in_expr() {
    assert_eq!(
        run_pascal(
            r#"program T; const A=1; const B=2; const C=3; const S=A+B+C; begin WriteLn(S); end."#
        ),
        &["6"]
    );
}

#[test]
fn const_shift_style() {
    assert_eq!(
        run_pascal(
            r#"program T; const N=1*2*2*2; begin WriteLn(N); end."#
        ),
        &["8"]
    );
}

#[test]
fn const_min_via_if() {
    assert_eq!(
        run_pascal(
            r#"program T; const A=7; const B=3; const M=Min(A,B); begin WriteLn(M); end."#
        ),
        &["3"]
    );
}

#[test]
fn const_abs_negative() {
    assert_eq!(
        run_pascal(
            r#"program T; const N=Abs(-9); begin WriteLn(N); end."#
        ),
        &["9"]
    );
}

#[test]
fn const_max_two() {
    assert_eq!(
        run_pascal(
            r#"program T; const N=Max(4,9); begin WriteLn(N); end."#
        ),
        &["9"]
    );
}

#[test]
fn const_min_two() {
    assert_eq!(
        run_pascal(
            r#"program T; const N=Min(4,9); begin WriteLn(N); end."#
        ),
        &["4"]
    );
}

#[test]
fn const_succ_pred() {
    assert_eq!(
        run_pascal(
            r#"program T; const N=Succ(Pred(5)); begin WriteLn(N); end."#
        ),
        &["5"]
    );
}

#[test]
fn const_char_add_offset() {
    assert_eq!(
        run_pascal(
            r#"program T; const C=Chr(Ord('a')+2); begin WriteLn(C); end."#
        ),
        &["c"]
    );
}

#[test]
fn const_real_int_mix() {
    assert_eq!(
        run_pascal(
            r#"program T; const N=Trunc(2.5+2.5); begin WriteLn(N); end."#
        ),
        &["5"]
    );
}

#[test]
fn const_nested_string() {
    assert_eq!(
        run_pascal(
            r#"program T; const P='x'; const Q=P+P+P; begin WriteLn(Q); end."#
        ),
        &["xxx"]
    );
}

#[test]
fn const_boolean_xor() {
    assert_eq!(
        run_pascal(
            r#"program T; const T=True xor False; begin WriteLn(T); end."#
        ),
        &["true"]
    );
}

#[test]
fn const_integer_div_real() {
    assert_eq!(
        run_pascal(
            r#"program T; const R=7/2; begin WriteLn(Trunc(R)); end."#
        ),
        &["3"]
    );
}

#[test]
fn const_factorial_style() {
    assert_eq!(
        run_pascal(
            r#"program T; const F=1*2*3*4; begin WriteLn(F); end."#
        ),
        &["24"]
    );
}

#[test]
fn const_bit_and_sim() {
    assert_eq!(
        run_pascal(
            r#"program T; const N=15 and 7; begin WriteLn(N); end."#
        ),
        &["7"]
    );
}

#[test]
fn const_bit_or_sim() {
    assert_eq!(
        run_pascal(
            r#"program T; const N=8 or 1; begin WriteLn(N); end."#
        ),
        &["9"]
    );
}

#[test]
fn const_expression_compare() {
    assert_eq!(
        run_pascal(
            r#"program T; const T=(100 div 10)=(5+5); begin WriteLn(T); end."#
        ),
        &["true"]
    );
}

#[test]
fn const_mixed_precedence() {
    assert_eq!(
        run_pascal(
            r#"program T; const N=2+3*4-5; begin WriteLn(N); end."#
        ),
        &["9"]
    );
}

#[test]
fn const_three_level() {
    assert_eq!(
        run_pascal(
            r#"program T; const A=1; const B=A+1; const C=B+1; const D=C+1; begin WriteLn(D); end."#
        ),
        &["4"]
    );
}

#[test]
fn const_string_repeat_style() {
    assert_eq!(
        run_pascal(
            r#"program T; const S='ha'; const T=S+S; begin WriteLn(T); end."#
        ),
        &["haha"]
    );
}

