use super::helpers::run_lua_one;

#[test]
fn test_numeric_bitwise_or_xor_or_self() {
    assert_eq!(run_lua_one("print(12 | 12)"), "12");
}

#[test]
fn test_numeric_bitwise_or_xor_or_merge() {
    assert_eq!(run_lua_one("print(12 | 5)"), "13");
}

#[test]
fn test_numeric_bitwise_or_xor_or_overlap() {
    assert_eq!(run_lua_one("print(7 | 3)"), "7");
}

#[test]
fn test_numeric_bitwise_or_xor_and_self() {
    assert_eq!(run_lua_one("print(10 & 10)"), "10");
}

#[test]
fn test_numeric_bitwise_or_xor_and_clear() {
    assert_eq!(run_lua_one("print(12 & 10)"), "8");
}

#[test]
fn test_numeric_bitwise_or_xor_xor_self() {
    assert_eq!(run_lua_one("print(12 ~ 12)"), "0");
}

#[test]
fn test_numeric_bitwise_or_xor_xor_basic() {
    assert_eq!(run_lua_one("print(12 ~ 10)"), "6");
}

#[test]
fn test_numeric_bitwise_or_xor_xor_chain() {
    assert_eq!(run_lua_one("print((12 ~ 10) ~ 10)"), "12");
}

#[test]
fn test_numeric_bitwise_or_xor_left_shift() {
    assert_eq!(run_lua_one("print((1 | 4) << 1)"), "10");
}

#[test]
fn test_numeric_bitwise_or_xor_right_shift() {
    assert_eq!(run_lua_one("print((12 ~ 4) >> 1)"), "4");
}

#[test]
fn test_numeric_bitwise_or_xor_shift_and() {
    assert_eq!(run_lua_one("print(((1 << 3) | 4) & 4)"), "4");
}

#[test]
fn test_numeric_bitwise_or_xor_shift_xor() {
    assert_eq!(run_lua_one("print(((1 << 3) | 4) ~ 4)"), "8");
}

#[test]
fn test_numeric_bitwise_or_xor_and_xor_combo() {
    assert_eq!(run_lua_one("print((13 & 7) ~ 1)"), "4");
}

#[test]
fn test_numeric_bitwise_or_xor_or_xor_combo() {
    assert_eq!(run_lua_one("print((13 | 7) ~ 2)"), "13");
}

#[test]
fn test_numeric_bitwise_or_xor_xor_zero() {
    assert_eq!(run_lua_one("print(9 ~ 0)"), "9");
}

#[test]
fn test_numeric_bitwise_or_xor_and_zero() {
    assert_eq!(run_lua_one("print(9 & 0)"), "0");
}

#[test]
fn test_numeric_bitwise_or_xor_shift_chain() {
    assert_eq!(run_lua_one("print(((1 | 2) << 3) >> 1)"), "12");
}

#[test]
fn test_numeric_bitwise_or_xor_or_math() {
    assert_eq!(run_lua_one("print((1 | 2 | 4) + 1)"), "8");
}

#[test]
fn test_numeric_bitwise_or_xor_and_math() {
    assert_eq!(run_lua_one("print((12 & 10) + 2)"), "10");
}

#[test]
fn test_numeric_bitwise_or_xor_type_check() {
    assert_eq!(run_lua_one("print(type(12 | 10))"), "number");
}
