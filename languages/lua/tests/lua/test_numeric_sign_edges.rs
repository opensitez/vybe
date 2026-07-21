use super::helpers::run_lua_one;

#[test]
fn test_numeric_sign_edges_add_positive_with_zero() {
    assert_eq!(run_lua_one("print(0 + 10)"), "10");
}

#[test]
fn test_numeric_sign_edges_add_negative_with_zero() {
    assert_eq!(run_lua_one("print(0 + -10)"), "-10");
}

#[test]
fn test_numeric_sign_edges_positive_plus_negative() {
    assert_eq!(run_lua_one("print(15 + -7)"), "8");
}

#[test]
fn test_numeric_sign_edges_double_negation() {
    assert_eq!(run_lua_one("print(-(-12))"), "12");
}

#[test]
fn test_numeric_sign_edges_triple_negation() {
    assert_eq!(run_lua_one("print(-(-(-12)))"), "-12");
}

#[test]
fn test_numeric_sign_edges_negative_subtract_positive() {
    assert_eq!(run_lua_one("print(-9 - 4)"), "-13");
}

#[test]
fn test_numeric_sign_edges_subtract_negative() {
    assert_eq!(run_lua_one("print(9 - -4)"), "13");
}

#[test]
fn test_numeric_sign_edges_multiply_negative() {
    assert_eq!(run_lua_one("print(-9 * 4)"), "-36");
}

#[test]
fn test_numeric_sign_edges_multiply_negatives() {
    assert_eq!(run_lua_one("print(-9 * -4)"), "36");
}

#[test]
fn test_numeric_sign_edges_divide_positive() {
    assert_eq!(run_lua_one("print(12 / 3)"), "4.0");
}

#[test]
fn test_numeric_sign_edges_divide_negative_numerator() {
    assert_eq!(run_lua_one("print(-12 / 3)"), "-4.0");
}

#[test]
fn test_numeric_sign_edges_divide_negative_denominator() {
    assert_eq!(run_lua_one("print(12 / -3)"), "-4.0");
}

#[test]
fn test_numeric_sign_edges_divide_both_negative() {
    assert_eq!(run_lua_one("print(-12 / -3)"), "4.0");
}

#[test]
fn test_numeric_sign_edges_floor_division_positive() {
    assert_eq!(run_lua_one("print(17 // 5)"), "3");
}

#[test]
fn test_numeric_sign_edges_floor_division_neg_num() {
    assert_eq!(run_lua_one("print(-17 // 5)"), "-4");
}

#[test]
fn test_numeric_sign_edges_floor_division_neg_den() {
    assert_eq!(run_lua_one("print(17 // -5)"), "-4");
}

#[test]
fn test_numeric_sign_edges_floor_division_both_neg() {
    assert_eq!(run_lua_one("print(-17 // -5)"), "3");
}

#[test]
fn test_numeric_sign_edges_modulo_positive() {
    assert_eq!(run_lua_one("print(17 % 5)"), "2");
}

#[test]
fn test_numeric_sign_edges_modulo_negative_left() {
    assert_eq!(run_lua_one("print(-17 % 5)"), "3");
}

#[test]
fn test_numeric_sign_edges_modulo_negative_right() {
    assert_eq!(run_lua_one("print(17 % -5)"), "-3");
}

#[test]
fn test_numeric_sign_edges_modulo_both_negative() {
    assert_eq!(run_lua_one("print(-17 % -5)"), "-2");
}
