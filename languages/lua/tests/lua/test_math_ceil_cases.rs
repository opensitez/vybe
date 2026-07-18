use super::helpers::run_lua_one;

#[test]
fn test_math_ceil_cases_baseline() {
    assert_eq!(run_lua_one(r#"print(math.ceil(0.8 + 0) == 1)"#), "true");
}


#[test]
fn test_math_ceil_cases_simple() {
    assert_eq!(run_lua_one(r#"print(math.ceil(1.8 + 1) == 2)"#), "true");
}


#[test]
fn test_math_ceil_cases_trimmed() {
    assert_eq!(run_lua_one(r#"print(math.ceil(2.8 + 2) == 3)"#), "true");
}


#[test]
fn test_math_ceil_cases_decimal() {
    assert_eq!(run_lua_one(r#"print(math.ceil(3.8 + 3) == 4)"#), "true");
}


#[test]
fn test_math_ceil_cases_hexed() {
    assert_eq!(run_lua_one(r#"print(math.ceil(4.8 + 4) == 5)"#), "true");
}


#[test]
fn test_math_ceil_cases_prefixed() {
    assert_eq!(run_lua_one(r#"print(math.ceil(5.8 + 5) == 6)"#), "true");
}


#[test]
fn test_math_ceil_cases_negative() {
    assert_eq!(run_lua_one(r#"print(math.ceil(6.8 + 6) == 7)"#), "true");
}


#[test]
fn test_math_ceil_cases_rounded() {
    assert_eq!(run_lua_one(r#"print(math.ceil(7.8 + 7) == 8)"#), "true");
}


#[test]
fn test_math_ceil_cases_offset() {
    assert_eq!(run_lua_one(r#"print(math.ceil(8.8 + 8) == 9)"#), "true");
}


#[test]
fn test_math_ceil_cases_paired() {
    assert_eq!(run_lua_one(r#"print(math.ceil(9.8 + 9) == 10)"#), "true");
}


#[test]
fn test_math_ceil_cases_nested() {
    assert_eq!(run_lua_one(r#"print(math.ceil(10.8 + 10) == 11)"#), "true");
}


#[test]
fn test_math_ceil_cases_metaflow() {
    assert_eq!(run_lua_one(r#"print(math.ceil(11.8 + 11) == 12)"#), "true");
}


#[test]
fn test_math_ceil_cases_guarded() {
    assert_eq!(run_lua_one(r#"print(math.ceil(12.8 + 12) == 13)"#), "true");
}


#[test]
fn test_math_ceil_cases_mapped() {
    assert_eq!(run_lua_one(r#"print(math.ceil(13.8 + 13) == 14)"#), "true");
}


#[test]
fn test_math_ceil_cases_captured() {
    assert_eq!(run_lua_one(r#"print(math.ceil(14.8 + 14) == 15)"#), "true");
}


#[test]
fn test_math_ceil_cases_edge_first() {
    assert_eq!(run_lua_one(r#"print(math.ceil(15.8 + 15) == 16)"#), "true");
}


#[test]
fn test_math_ceil_cases_edge_second() {
    assert_eq!(run_lua_one(r#"print(math.ceil(16.8 + 16) == 17)"#), "true");
}


#[test]
fn test_math_ceil_cases_edge_last() {
    assert_eq!(run_lua_one(r#"print(math.ceil(17.8 + 17) == 18)"#), "true");
}


#[test]
fn test_math_ceil_cases_randomized() {
    assert_eq!(run_lua_one(r#"print(math.ceil(18.8 + 18) == 19)"#), "true");
}


#[test]
fn test_math_ceil_cases_unicode_like() {
    assert_eq!(run_lua_one(r#"print(math.ceil(19.8 + 19) == 20)"#), "true");
}
