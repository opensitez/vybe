use super::helpers::run_lua_one;

#[test]
fn test_numeric_floor_division_rules_integer_exact() {
    assert_eq!(run_lua_one("print(12 // 3)"), "4");
}

#[test]
fn test_numeric_floor_division_rules_integer_remainder() {
    assert_eq!(run_lua_one("print(13 // 3)"), "4");
}

#[test]
fn test_numeric_floor_division_rules_integer_negative_num() {
    assert_eq!(run_lua_one("print(-13 // 3)"), "-5");
}

#[test]
fn test_numeric_floor_division_rules_integer_negative_den() {
    assert_eq!(run_lua_one("print(13 // -3)"), "-5");
}

#[test]
fn test_numeric_floor_division_rules_integer_both_negative() {
    assert_eq!(run_lua_one("print(-13 // -3)"), "4");
}

#[test]
fn test_numeric_floor_division_rules_float_exact() {
    assert_eq!(run_lua_one("print(10.0 // 2.0)"), "5");
}

#[test]
fn test_numeric_floor_division_rules_float_fraction() {
    assert_eq!(run_lua_one("print(10.0 // 3.0)"), "3");
}

#[test]
fn test_numeric_floor_division_rules_float_negative_num() {
    assert_eq!(run_lua_one("print(-10.0 // 3.0)"), "-4");
}

#[test]
fn test_numeric_floor_division_rules_float_negative_den() {
    assert_eq!(run_lua_one("print(10.0 // -3.0)"), "-4");
}

#[test]
fn test_numeric_floor_division_rules_float_both_negative() {
    assert_eq!(run_lua_one("print(-10.0 // -3.0)"), "3");
}

#[test]
fn test_numeric_floor_division_rules_fractional_result() {
    assert_eq!(run_lua_one("print(7 // 3)"), "2");
}

#[test]
fn test_numeric_floor_division_rules_small_fraction() {
    assert_eq!(run_lua_one("print(1 // 5)"), "0");
}

#[test]
fn test_numeric_floor_division_rules_small_negative_fraction() {
    assert_eq!(run_lua_one("print(-1 // 5)"), "-1");
}

#[test]
fn test_numeric_floor_division_rules_neg_small_fraction() {
    assert_eq!(run_lua_one("print(1 // -5)"), "-1");
}

#[test]
fn test_numeric_floor_division_rules_zero_dividend() {
    assert_eq!(run_lua_one("print(0 // 5)"), "0");
}

#[test]
fn test_numeric_floor_division_rules_nested() {
    assert_eq!(run_lua_one("print((20 // 5) // 2)"), "2");
}

#[test]
fn test_numeric_floor_division_rules_nested_mixed_ops() {
    assert_eq!(run_lua_one("print((20 // 3) + 1)"), "7");
}

#[test]
fn test_numeric_floor_division_rules_chain() {
    assert_eq!(run_lua_one("print(100 // 10 // 2)"), "5");
}

#[test]
fn test_numeric_floor_division_rules_zero_divisor() {
    assert_eq!(run_lua_one("print(12 // 0)"), "inf");
}

#[test]
fn test_numeric_floor_division_rules_zero_divisor_negative() {
    assert_eq!(run_lua_one("print(12 // -0)"), "-inf");
}
