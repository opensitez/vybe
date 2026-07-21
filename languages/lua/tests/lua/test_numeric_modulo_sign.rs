use super::helpers::run_lua_one;

#[test]
fn test_numeric_modulo_sign_pos_pos_small() {
    assert_eq!(run_lua_one("print(8 % 3)"), "2");
}

#[test]
fn test_numeric_modulo_sign_pos_pos_large() {
    assert_eq!(run_lua_one("print(20 % 7)"), "6");
}

#[test]
fn test_numeric_modulo_sign_pos_pos_eq() {
    assert_eq!(run_lua_one("print(12 % 6)"), "0");
}

#[test]
fn test_numeric_modulo_sign_pos_zero() {
    assert_eq!(run_lua_one("print(0 % 5)"), "0");
}

#[test]
fn test_numeric_modulo_sign_neg_pos() {
    assert_eq!(run_lua_one("print(-8 % 3)"), "1");
}

#[test]
fn test_numeric_modulo_sign_neg_pos_large() {
    assert_eq!(run_lua_one("print(-20 % 7)"), "1");
}

#[test]
fn test_numeric_modulo_sign_pos_neg() {
    assert_eq!(run_lua_one("print(8 % -3)"), "-1");
}

#[test]
fn test_numeric_modulo_sign_neg_neg() {
    assert_eq!(run_lua_one("print(-8 % -3)"), "-2");
}

#[test]
fn test_numeric_modulo_sign_neg_neg_large() {
    assert_eq!(run_lua_one("print(-20 % -7)"), "-6");
}

#[test]
fn test_numeric_modulo_sign_float_pos_pos() {
    assert_eq!(run_lua_one("print(8.5 % 3)"), "2.5");
}

#[test]
fn test_numeric_modulo_sign_float_neg_pos() {
    assert_eq!(run_lua_one("print(-8.5 % 3)"), "0.5");
}

#[test]
fn test_numeric_modulo_sign_float_pos_neg() {
    assert_eq!(run_lua_one("print(8.5 % -3)"), "-0.5");
}

#[test]
fn test_numeric_modulo_sign_float_neg_neg() {
    assert_eq!(run_lua_one("print(-8.5 % -3)"), "-2.5");
}

#[test]
fn test_numeric_modulo_sign_identity_rule() {
    assert_eq!(run_lua_one("print(5 % 9)"), "5");
}

#[test]
fn test_numeric_modulo_sign_identity_rule_negative() {
    assert_eq!(run_lua_one("print(-5 % 9)"), "4");
}

#[test]
fn test_numeric_modulo_sign_chain_inner() {
    assert_eq!(run_lua_one("print((20 % 7) % 3)"), "0");
}

#[test]
fn test_numeric_modulo_sign_chain_outer() {
    assert_eq!(run_lua_one("print(20 % (7 % 3))"), "0");
}

#[test]
fn test_numeric_modulo_sign_with_addition() {
    assert_eq!(run_lua_one("print((20 % 7) + 1)"), "7");
}

#[test]
fn test_numeric_modulo_sign_with_subtraction() {
    assert_eq!(run_lua_one("print((20 % 7) - 1)"), "5");
}

#[test]
fn test_numeric_modulo_sign_with_multiplication() {
    assert_eq!(run_lua_one("print((20 % 7) * 2)"), "12");
}

#[test]
fn test_numeric_modulo_sign_with_division() {
    assert_eq!(run_lua_one("print((20 % 7) / 2)"), "3.0");
}
