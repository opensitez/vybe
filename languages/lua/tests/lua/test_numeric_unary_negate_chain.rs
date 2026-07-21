use super::helpers::run_lua_one;

#[test]
fn test_numeric_unary_negate_chain_single_negation() {
    assert_eq!(run_lua_one("print(-8)"), "-8");
}

#[test]
fn test_numeric_unary_negate_chain_double_negation() {
    assert_eq!(run_lua_one("print(- -8)"), "8");
}

#[test]
fn test_numeric_unary_negate_chain_triple_negation() {
    assert_eq!(run_lua_one("print(- - -8)"), "-8");
}

#[test]
fn test_numeric_unary_negate_chain_quadruple_negation() {
    assert_eq!(run_lua_one("print(- - - -8)"), "8");
}

#[test]
fn test_numeric_unary_negate_chain_negate_variable() {
    assert_eq!(run_lua_one("local n = 12; print(-n)"), "-12");
}

#[test]
fn test_numeric_unary_negate_chain_negate_expression() {
    assert_eq!(run_lua_one("print(-(2 + 3))"), "-5");
}

#[test]
fn test_numeric_unary_negate_chain_negate_mul() {
    assert_eq!(run_lua_one("print(-(2 * 5))"), "-10");
}

#[test]
fn test_numeric_unary_negate_chain_negate_div() {
    assert_eq!(run_lua_one("print(-(10 / 2))"), "-5.0");
}

#[test]
fn test_numeric_unary_negate_chain_negate_float() {
    assert_eq!(run_lua_one("print(-0.25)"), "-0.25");
}

#[test]
fn test_numeric_unary_negate_chain_negate_negative() {
    assert_eq!(run_lua_one("print(-(-0.25))"), "0.25");
}

#[test]
fn test_numeric_unary_negate_chain_negate_function() {
    assert_eq!(run_lua_one("function v() return 9 end; print(-v())"), "-9");
}

#[test]
fn test_numeric_unary_negate_chain_nested_field() {
    assert_eq!(run_lua_one("local t = {v = 7}; print(-t.v)"), "-7");
}

#[test]
fn test_numeric_unary_negate_chain_nested_access() {
    assert_eq!(run_lua_one("local t = {u = {v = 7}}; print(-t.u.v)"), "-7");
}

#[test]
fn test_numeric_unary_negate_chain_add_before_neg() {
    assert_eq!(run_lua_one("print(-(4 + 5))"), "-9");
}

#[test]
fn test_numeric_unary_negate_chain_sub_before_neg() {
    assert_eq!(run_lua_one("print(-(10 - 3))"), "-7");
}

#[test]
fn test_numeric_unary_negate_chain_compare() {
    assert_eq!(run_lua_one("print(- (1 + 2) == -3)"), "true");
}

#[test]
fn test_numeric_unary_negate_chain_pow_wrap() {
    assert_eq!(run_lua_one("print(-(2 ^ 3))"), "-8.0");
}

#[test]
fn test_numeric_unary_negate_chain_or() {
    assert_eq!(run_lua_one("print((-(5) or 0))"), "-5");
}

#[test]
fn test_numeric_unary_negate_chain_mix_with_modulo() {
    assert_eq!(run_lua_one("print(-(10 % 4))"), "-2");
}

#[test]
fn test_numeric_unary_negate_chain_nested_and_composed() {
    assert_eq!(run_lua_one("print(-(-(-(-1))))"), "1");
}
