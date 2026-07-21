use super::helpers::run_lua_one;

#[test]
fn test_numeric_type_mix_edges_decimal_string_to_number() {
    assert_eq!(run_lua_one("print(tonumber('21'))"), "21");
}

#[test]
fn test_numeric_type_mix_edges_hex_string_to_number() {
    assert_eq!(run_lua_one("print(tonumber('0x10'))"), "16");
}

#[test]
fn test_numeric_type_mix_edges_base_two_string() {
    assert_eq!(run_lua_one("print(tonumber('1011', 2))"), "11");
}

#[test]
fn test_numeric_type_mix_edges_base_eight_string() {
    assert_eq!(run_lua_one("print(tonumber('17', 8))"), "15");
}

#[test]
fn test_numeric_type_mix_edges_invalid_base_digit() {
    assert_eq!(run_lua_one("print(tonumber('19', 8))"), "nil");
}

#[test]
fn test_numeric_type_mix_edges_tostring_number() {
    assert_eq!(run_lua_one("print(tostring(42))"), "42");
}

#[test]
fn test_numeric_type_mix_edges_tostring_negative() {
    assert_eq!(run_lua_one("print(tostring(-11))"), "-11");
}

#[test]
fn test_numeric_type_mix_edges_tostring_float() {
    assert_eq!(run_lua_one("print(tostring(2.5))"), "2.5");
}

#[test]
fn test_numeric_type_mix_edges_tostring_boolean_true() {
    assert_eq!(run_lua_one("print(tostring(true))"), "true");
}

#[test]
fn test_numeric_type_mix_edges_tostring_boolean_false() {
    assert_eq!(run_lua_one("print(tostring(false))"), "false");
}

#[test]
fn test_numeric_type_mix_edges_tostring_nil() {
    assert_eq!(run_lua_one("print(tostring(nil))"), "nil");
}

#[test]
fn test_numeric_type_mix_edges_type_number() {
    assert_eq!(run_lua_one("print(type(1.5))"), "number");
}

#[test]
fn test_numeric_type_mix_edges_type_string_concat() {
    assert_eq!(run_lua_one("print(type('9' .. '8'))"), "string");
}

#[test]
fn test_numeric_type_mix_edges_integer_to_float() {
    assert_eq!(run_lua_one("print(7 + 0.5)"), "7.5");
}

#[test]
fn test_numeric_type_mix_edges_floor_mix() {
    assert_eq!(run_lua_one("print(math.floor(7.5))"), "7");
}

#[test]
fn test_numeric_type_mix_edges_ceil_mix() {
    assert_eq!(run_lua_one("print(math.ceil(7.1))"), "8");
}

#[test]
fn test_numeric_type_mix_edges_reject_parse_base_zero() {
    assert_eq!(
        run_lua_one("local ok = pcall(function() tonumber('010', 0) end); print(ok)"),
        "false"
    );
}

#[test]
fn test_numeric_type_mix_edges_parse_scientific_string() {
    assert_eq!(run_lua_one("print(tonumber('3e1'))"), "30.0");
}

#[test]
fn test_numeric_type_mix_edges_float_add_integer() {
    assert_eq!(run_lua_one("print(0.5 + 2)"), "2.5");
}

#[test]
fn test_numeric_type_mix_edges_float_times_integer() {
    assert_eq!(run_lua_one("print(0.5 * 8)"), "4.0");
}

#[test]
fn test_numeric_type_mix_edges_type_nil() {
    assert_eq!(run_lua_one("print(type(tonumber('bad')) )"), "nil");
}
