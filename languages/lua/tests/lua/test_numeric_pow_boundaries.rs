use super::helpers::run_lua_one;

#[test]
fn test_numeric_pow_boundaries_zero_power_zero() {
    assert_eq!(run_lua_one("print(0 ^ 0)"), "1");
}

#[test]
fn test_numeric_pow_boundaries_zero_power_positive() {
    assert_eq!(run_lua_one("print(0 ^ 8)"), "0");
}

#[test]
fn test_numeric_pow_boundaries_one_power_zero() {
    assert_eq!(run_lua_one("print(1 ^ 10)"), "1");
}

#[test]
fn test_numeric_pow_boundaries_one_power_negative() {
    assert_eq!(run_lua_one("print(1 ^ -10)"), "1");
}

#[test]
fn test_numeric_pow_boundaries_two_power_zero() {
    assert_eq!(run_lua_one("print(2 ^ 0)"), "1");
}

#[test]
fn test_numeric_pow_boundaries_two_power_one() {
    assert_eq!(run_lua_one("print(2 ^ 1)"), "2");
}

#[test]
fn test_numeric_pow_boundaries_two_power_five() {
    assert_eq!(run_lua_one("print(2 ^ 5)"), "32");
}

#[test]
fn test_numeric_pow_boundaries_two_power_negative() {
    assert_eq!(run_lua_one("print(2 ^ -2)"), "0.25");
}

#[test]
fn test_numeric_pow_boundaries_three_power_three() {
    assert_eq!(run_lua_one("print(3 ^ 3)"), "27");
}

#[test]
fn test_numeric_pow_boundaries_even_power() {
    assert_eq!(run_lua_one("print((-2) ^ 4)"), "16");
}

#[test]
fn test_numeric_pow_boundaries_odd_power() {
    assert_eq!(run_lua_one("print((-2) ^ 3)"), "-8");
}

#[test]
fn test_numeric_pow_boundaries_large_power_expression() {
    assert_eq!(run_lua_one("print((2 ^ 3) ^ 2)"), "64");
}

#[test]
fn test_numeric_pow_boundaries_chain_expression() {
    assert_eq!(run_lua_one("print(2 ^ (3 ^ 1))"), "8");
}

#[test]
fn test_numeric_pow_boundaries_mul_by_power() {
    assert_eq!(run_lua_one("print((2 ^ 3) * 2)"), "16");
}

#[test]
fn test_numeric_pow_boundaries_div_by_power() {
    assert_eq!(run_lua_one("print((2 ^ 3) / 2)"), "4");
}

#[test]
fn test_numeric_pow_boundaries_plus_power() {
    assert_eq!(run_lua_one("print((2 ^ 3) + 1)"), "9");
}

#[test]
fn test_numeric_pow_boundaries_minus_power() {
    assert_eq!(run_lua_one("print((2 ^ 4) - 1)"), "15");
}

#[test]
fn test_numeric_pow_boundaries_pow_mix() {
    assert_eq!(run_lua_one("print(2 ^ 3 % 3)"), "2");
}

#[test]
fn test_numeric_pow_boundaries_fractional_base() {
    assert_eq!(run_lua_one("print(4.5 ^ 2)"), "20.25");
}

#[test]
fn test_numeric_pow_boundaries_sqrt_via_power() {
    assert_eq!(run_lua_one("print(16 ^ 0.5)"), "4");
}
