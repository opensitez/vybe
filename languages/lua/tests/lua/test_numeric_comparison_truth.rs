use super::helpers::run_lua_one;

#[test]
fn test_numeric_comparison_truth_less_true() {
    assert_eq!(run_lua_one("print(1 < 2)"), "true");
}

#[test]
fn test_numeric_comparison_truth_less_false() {
    assert_eq!(run_lua_one("print(3 < 2)"), "false");
}

#[test]
fn test_numeric_comparison_truth_less_equal_true() {
    assert_eq!(run_lua_one("print(2 <= 2)"), "true");
}

#[test]
fn test_numeric_comparison_truth_less_equal_false() {
    assert_eq!(run_lua_one("print(3 <= 2)"), "false");
}

#[test]
fn test_numeric_comparison_truth_equal_true() {
    assert_eq!(run_lua_one("print(8 == 8)"), "true");
}

#[test]
fn test_numeric_comparison_truth_equal_false() {
    assert_eq!(run_lua_one("print(8 == 9)"), "false");
}

#[test]
fn test_numeric_comparison_truth_not_equal_true() {
    assert_eq!(run_lua_one("print(8 ~= 9)"), "true");
}

#[test]
fn test_numeric_comparison_truth_not_equal_false() {
    assert_eq!(run_lua_one("print(8 ~= 8)"), "false");
}

#[test]
fn test_numeric_comparison_truth_greater_true() {
    assert_eq!(run_lua_one("print(9 > 4)"), "true");
}

#[test]
fn test_numeric_comparison_truth_greater_false() {
    assert_eq!(run_lua_one("print(4 > 9)"), "false");
}

#[test]
fn test_numeric_comparison_truth_greater_equal_true() {
    assert_eq!(run_lua_one("print(9 >= 9)"), "true");
}

#[test]
fn test_numeric_comparison_truth_greater_equal_false() {
    assert_eq!(run_lua_one("print(8 >= 9)"), "false");
}

#[test]
fn test_numeric_comparison_truth_chain_all_true() {
    assert_eq!(run_lua_one("print(1 < 2 and 2 < 3 and 3 < 4)"), "true");
}

#[test]
fn test_numeric_comparison_truth_chain_with_false() {
    assert_eq!(run_lua_one("print(1 < 2 and 4 < 3 and 3 < 4)"), "false");
}

#[test]
fn test_numeric_comparison_truth_or_true() {
    assert_eq!(run_lua_one("print(1 > 2 or 3 > 2)"), "true");
}

#[test]
fn test_numeric_comparison_truth_or_false() {
    assert_eq!(run_lua_one("print(1 > 2 or 2 > 3)"), "false");
}

#[test]
fn test_numeric_comparison_truth_not_true() {
    assert_eq!(run_lua_one("print(not (2 == 2))"), "false");
}

#[test]
fn test_numeric_comparison_truth_not_false() {
    assert_eq!(run_lua_one("print(not (2 == 3))"), "true");
}

#[test]
fn test_numeric_comparison_truth_floor_vs_division() {
    assert_eq!(run_lua_one("print((10 // 3) < 4)"), "true");
}

#[test]
fn test_numeric_comparison_truth_power_cmp() {
    assert_eq!(run_lua_one("print((2 ^ 3) == 8)"), "true");
}

#[test]
fn test_numeric_comparison_truth_float_precision() {
    assert_eq!(run_lua_one("print(0.1 + 0.2 == 0.3)"), "true");
}

#[test]
fn test_numeric_comparison_truth_float_gap() {
    assert_eq!(run_lua_one("print(0.1 + 0.2 == 0.3001)"), "false");
}
