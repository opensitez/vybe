use super::helpers::run_lua_one;

#[test]
fn test_numeric_hex_and_exponent_hex_lower() {
    assert_eq!(run_lua_one("print(0x10)"), "16");
}

#[test]
fn test_numeric_hex_and_exponent_hex_upper() {
    assert_eq!(run_lua_one("print(0X10)"), "16");
}

#[test]
fn test_numeric_hex_and_exponent_hex_negative() {
    assert_eq!(run_lua_one("print(-0x8)"), "-8");
}

#[test]
fn test_numeric_hex_and_exponent_hex_addition() {
    assert_eq!(run_lua_one("print(0x10 + 1)"), "17");
}

#[test]
fn test_numeric_hex_and_exponent_string_hex() {
    assert_eq!(run_lua_one("print(tonumber('0x20'))"), "32");
}

#[test]
fn test_numeric_hex_and_exponent_string_hex_negative() {
    assert_eq!(run_lua_one("print(tonumber('-0x10'))"), "-16");
}

#[test]
fn test_numeric_hex_and_exponent_scientific_small() {
    assert_eq!(run_lua_one("print(1e-1)"), "0.1");
}

#[test]
fn test_numeric_hex_and_exponent_scientific_large() {
    assert_eq!(run_lua_one("print(3e3)"), "3000.0");
}

#[test]
fn test_numeric_hex_and_exponent_scientific_decimal() {
    assert_eq!(run_lua_one("print(1.5e2)"), "150.0");
}

#[test]
fn test_numeric_hex_and_exponent_scientific_negative_exponent() {
    assert_eq!(run_lua_one("print(5e-2)"), "0.05");
}

#[test]
fn test_numeric_hex_and_exponent_scientific_negative_value() {
    assert_eq!(run_lua_one("print(-2.5e1)"), "-25.0");
}

#[test]
fn test_numeric_hex_and_exponent_parse_scientific_string() {
    assert_eq!(run_lua_one("print(tonumber('1e2'))"), "100.0");
}

#[test]
fn test_numeric_hex_and_exponent_hex_to_scientific_sum() {
    assert_eq!(run_lua_one("print(0x10 + 1e1)"), "26");
}

#[test]
fn test_numeric_hex_and_exponent_hex_times_scientific() {
    assert_eq!(run_lua_one("print(0x10 * 1e1)"), "160.0");
}

#[test]
fn test_numeric_hex_and_exponent_exponent_floor_interaction() {
    assert_eq!(run_lua_one("print(2e1 // 5)"), "4.0");
}

#[test]
fn test_numeric_hex_and_exponent_mod_with_scientific() {
    assert_eq!(run_lua_one("print(25 % 1e1)"), "5.0");
}

#[test]
fn test_numeric_hex_and_exponent_power_with_scientific() {
    assert_eq!(run_lua_one("print(2 ^ 1e1)"), "1024.0");
}

#[test]
fn test_numeric_hex_and_exponent_fractional_hex_addition() {
    assert_eq!(run_lua_one("print(0x10 + 0.5)"), "16.5");
}

#[test]
fn test_numeric_hex_and_exponent_fractional_to_hex_sum() {
    assert_eq!(run_lua_one("print(1.5 + 0x10)"), "17.5");
}

#[test]
fn test_numeric_hex_and_exponent_string_hex_tostring() {
    assert_eq!(run_lua_one("print(tostring(tonumber('0xF') + 1))"), "16");
}
