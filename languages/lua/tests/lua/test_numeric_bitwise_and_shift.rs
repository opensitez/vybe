use super::helpers::run_lua_one;

#[test]
fn test_numeric_bitwise_and_shift_and_basic() {
    assert_eq!(run_lua_one("print(6 & 3)"), "2");
}

#[test]
fn test_numeric_bitwise_and_shift_and_zero() {
    assert_eq!(run_lua_one("print(6 & 0)"), "0");
}

#[test]
fn test_numeric_bitwise_and_shift_and_self() {
    assert_eq!(run_lua_one("print(7 & 7)"), "7");
}

#[test]
fn test_numeric_bitwise_and_shift_or_basic() {
    assert_eq!(run_lua_one("print(6 | 3)"), "7");
}

#[test]
fn test_numeric_bitwise_and_shift_or_zero() {
    assert_eq!(run_lua_one("print(6 | 0)"), "6");
}

#[test]
fn test_numeric_bitwise_and_shift_or_self() {
    assert_eq!(run_lua_one("print(6 | 6)"), "6");
}

#[test]
fn test_numeric_bitwise_and_shift_xor_basic() {
    assert_eq!(run_lua_one("print(6 ~ 3)"), "5");
}

#[test]
fn test_numeric_bitwise_and_shift_xor_self() {
    assert_eq!(run_lua_one("print(6 ~ 6)"), "0");
}

#[test]
fn test_numeric_bitwise_and_shift_left_one() {
    assert_eq!(run_lua_one("print(3 << 1)"), "6");
}

#[test]
fn test_numeric_bitwise_and_shift_right_one() {
    assert_eq!(run_lua_one("print(8 >> 1)"), "4");
}

#[test]
fn test_numeric_bitwise_and_shift_left_two() {
    assert_eq!(run_lua_one("print(1 << 3)"), "8");
}

#[test]
fn test_numeric_bitwise_and_shift_right_two() {
    assert_eq!(run_lua_one("print(32 >> 2)"), "8");
}

#[test]
fn test_numeric_bitwise_and_shift_combined() {
    assert_eq!(run_lua_one("print((5 | 2) & 6)"), "6");
}

#[test]
fn test_numeric_bitwise_and_shift_combined_xor() {
    assert_eq!(run_lua_one("print((5 ~ 2) << 1)"), "12");
}

#[test]
fn test_numeric_bitwise_and_shift_shift_mix() {
    assert_eq!(run_lua_one("print((8 >> 2) << 1)"), "4");
}

#[test]
fn test_numeric_bitwise_and_shift_shift_mix_alt() {
    assert_eq!(run_lua_one("print((10 << 1) >> 1)"), "10");
}

#[test]
fn test_numeric_bitwise_and_shift_left_zero() {
    assert_eq!(run_lua_one("print(5 << 0)"), "5");
}

#[test]
fn test_numeric_bitwise_and_shift_right_zero() {
    assert_eq!(run_lua_one("print(5 >> 0)"), "5");
}

#[test]
fn test_numeric_bitwise_and_shift_ternary_like() {
    assert_eq!(run_lua_one("print((6 & 3) | 8)"), "10");
}

#[test]
fn test_numeric_bitwise_and_shift_nested_bits() {
    assert_eq!(run_lua_one("print(((3 << 2) & 12) | 1)"), "13");
}

#[test]
fn test_numeric_bitwise_and_shift_result_type() {
    assert_eq!(run_lua_one("print(type(3 << 2))"), "number");
}

#[test]
fn test_numeric_bitwise_and_shift_flag_logic() {
    assert_eq!(run_lua_one("print(((1 | 2 | 4) ~ 2) & 6)"), "4");
}
