use super::helpers::run_pascal;

// Arithmetic
#[test]
fn arith_add() {
    assert_eq!(run_pascal("program T; begin WriteLn(3 + 4); end."), &["7"]);
}
#[test]
fn arith_sub() {
    assert_eq!(run_pascal("program T; begin WriteLn(10 - 3); end."), &["7"]);
}
#[test]
fn arith_mul() {
    assert_eq!(run_pascal("program T; begin WriteLn(6 * 7); end."), &["42"]);
}
#[test]
fn arith_div_real() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(10 / 4); end."),
        &["2.5"]
    );
}
#[test]
fn arith_idiv() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(10 div 3); end."),
        &["3"]
    );
}
#[test]
fn arith_mod() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(10 mod 3); end."),
        &["1"]
    );
}
#[test]
fn arith_neg() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(-(3 + 4)); end."),
        &["-7"]
    );
}
#[test]
fn arith_precedence() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(2 + 3 * 4); end."),
        &["14"]
    );
}
#[test]
fn arith_parens() {
    assert_eq!(
        run_pascal("program T; begin WriteLn((2 + 3) * 4); end."),
        &["20"]
    );
}
#[test]
fn arith_chain() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(1 + 2 + 3 + 4); end."),
        &["10"]
    );
}
#[test]
fn arith_mixed() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(10 - 2 * 3); end."),
        &["4"]
    );
}
#[test]
fn arith_mod_zero() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(7 mod 7); end."),
        &["0"]
    );
}

// Comparison
#[test]
fn cmp_eq_true() {
    assert_eq!(
        run_pascal("program T; begin if 5 = 5 then WriteLn('y') else WriteLn('n'); end."),
        &["y"]
    );
}
#[test]
fn cmp_eq_false() {
    assert_eq!(
        run_pascal("program T; begin if 5 = 6 then WriteLn('y') else WriteLn('n'); end."),
        &["n"]
    );
}
#[test]
fn cmp_ne() {
    assert_eq!(
        run_pascal("program T; begin if 5 <> 6 then WriteLn('y') else WriteLn('n'); end."),
        &["y"]
    );
}
#[test]
fn cmp_lt() {
    assert_eq!(
        run_pascal("program T; begin if 3 < 5 then WriteLn('y') else WriteLn('n'); end."),
        &["y"]
    );
}
#[test]
fn cmp_gt() {
    assert_eq!(
        run_pascal("program T; begin if 5 > 3 then WriteLn('y') else WriteLn('n'); end."),
        &["y"]
    );
}
#[test]
fn cmp_le() {
    assert_eq!(
        run_pascal("program T; begin if 3 <= 3 then WriteLn('y') else WriteLn('n'); end."),
        &["y"]
    );
}
#[test]
fn cmp_ge() {
    assert_eq!(
        run_pascal("program T; begin if 5 >= 5 then WriteLn('y') else WriteLn('n'); end."),
        &["y"]
    );
}
#[test]
fn cmp_str_eq() {
    assert_eq!(
        run_pascal("program T; begin if 'abc' = 'abc' then WriteLn('y') else WriteLn('n'); end."),
        &["y"]
    );
}
#[test]
fn cmp_str_ne() {
    assert_eq!(
        run_pascal("program T; begin if 'abc' <> 'xyz' then WriteLn('y') else WriteLn('n'); end."),
        &["y"]
    );
}

// Boolean / Logical
#[test]
fn bool_and_tt() {
    assert_eq!(
        run_pascal("program T; begin if true and true then WriteLn('y') else WriteLn('n'); end."),
        &["y"]
    );
}
#[test]
fn bool_and_tf() {
    assert_eq!(
        run_pascal("program T; begin if true and false then WriteLn('y') else WriteLn('n'); end."),
        &["n"]
    );
}
#[test]
fn bool_or_tf() {
    assert_eq!(
        run_pascal("program T; begin if true or false then WriteLn('y') else WriteLn('n'); end."),
        &["y"]
    );
}
#[test]
fn bool_or_ff() {
    assert_eq!(
        run_pascal("program T; begin if false or false then WriteLn('y') else WriteLn('n'); end."),
        &["n"]
    );
}
#[test]
fn bool_not_t() {
    assert_eq!(
        run_pascal("program T; begin if not true then WriteLn('y') else WriteLn('n'); end."),
        &["n"]
    );
}
#[test]
fn bool_not_f() {
    assert_eq!(
        run_pascal("program T; begin if not false then WriteLn('y') else WriteLn('n'); end."),
        &["y"]
    );
}
#[test]
fn bool_short_and() {
    assert_eq!(
        run_pascal("program T; begin if false and true then WriteLn('y') else WriteLn('n'); end."),
        &["n"]
    );
}
#[test]
fn bool_short_or() {
    assert_eq!(
        run_pascal("program T; begin if true or false then WriteLn('y') else WriteLn('n'); end."),
        &["y"]
    );
}
#[test]
fn bool_compound() {
    assert_eq!(
        run_pascal("program T; begin if (3 > 2) and (5 > 4) then WriteLn('y'); end."),
        &["y"]
    );
}
#[test]
fn bool_complex() {
    assert_eq!(
        run_pascal("program T; begin if (1 < 2) or (10 < 5) then WriteLn('y'); end."),
        &["y"]
    );
}

#[test]
fn logical_xor_true_when_exactly_one_true() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(True xor False); end."#),
        &["true"]
    );
}

#[test]
fn logical_xor_false_when_both_same() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(True xor True); end."#),
        &["false"]
    );
}

#[test]
fn logical_not_inverts_false() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(not False); end."#),
        &["true"]
    );
}

#[test]
fn integer_bitwise_xor_operator() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(5 xor 3); end."#),
        &["6"]
    );
}

#[test]
fn integer_bitwise_or_operator() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(8 or 1); end."#),
        &["9"]
    );
}

#[test]
fn integer_bitwise_and_operator() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(7 and 3); end."#),
        &["3"]
    );
}

#[test]
fn integer_bitwise_not_operator() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(not 0); end."#),
        &["-1"]
    );
}

#[test]
fn shift_left_doubles_bits() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(1 shl 3); end."#),
        &["8"]
    );
}

#[test]
fn shift_right_halves_bits() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(16 shr 2); end."#),
        &["4"]
    );
}

#[test]
fn real_equality_exact_match() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(1.5 = 1.5); end."#),
        &["True"]
    );
}

#[test]
fn string_not_equal_operator() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn('a' <> 'b'); end."#),
        &["True"]
    );
}

#[test]
fn chained_range_check_with_and() {
    assert_eq!(
        run_pascal(
            r#"program T; var n: Integer; begin n := 5; WriteLn((n > 1) and (n < 10)); end."#
        ),
        &["true"]
    );
}
