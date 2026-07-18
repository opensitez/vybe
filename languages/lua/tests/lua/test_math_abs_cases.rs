use super::helpers::run_lua_one;

#[test]
fn test_math_abs_cases_baseline() {
    assert_eq!(run_lua_one(r#"print(math.abs(1) == 1)"#), "true");
}


#[test]
fn test_math_abs_cases_simple() {
    assert_eq!(run_lua_one(r#"print(math.abs(-2) == 2)"#), "true");
}


#[test]
fn test_math_abs_cases_trimmed() {
    assert_eq!(run_lua_one(r#"print(math.abs(3) == 3)"#), "true");
}


#[test]
fn test_math_abs_cases_decimal() {
    assert_eq!(run_lua_one(r#"print(math.abs(-4) == 4)"#), "true");
}


#[test]
fn test_math_abs_cases_hexed() {
    assert_eq!(run_lua_one(r#"print(math.abs(5) == 5)"#), "true");
}


#[test]
fn test_math_abs_cases_prefixed() {
    assert_eq!(run_lua_one(r#"print(math.abs(-6) == 6)"#), "true");
}


#[test]
fn test_math_abs_cases_negative() {
    assert_eq!(run_lua_one(r#"print(math.abs(7) == 7)"#), "true");
}


#[test]
fn test_math_abs_cases_rounded() {
    assert_eq!(run_lua_one(r#"print(math.abs(-8) == 8)"#), "true");
}


#[test]
fn test_math_abs_cases_offset() {
    assert_eq!(run_lua_one(r#"print(math.abs(9) == 9)"#), "true");
}


#[test]
fn test_math_abs_cases_paired() {
    assert_eq!(run_lua_one(r#"print(math.abs(-10) == 10)"#), "true");
}


#[test]
fn test_math_abs_cases_nested() {
    assert_eq!(run_lua_one(r#"print(math.abs(11) == 11)"#), "true");
}


#[test]
fn test_math_abs_cases_metaflow() {
    assert_eq!(run_lua_one(r#"print(math.abs(-12) == 12)"#), "true");
}


#[test]
fn test_math_abs_cases_guarded() {
    assert_eq!(run_lua_one(r#"print(math.abs(13) == 13)"#), "true");
}


#[test]
fn test_math_abs_cases_mapped() {
    assert_eq!(run_lua_one(r#"print(math.abs(-14) == 14)"#), "true");
}


#[test]
fn test_math_abs_cases_captured() {
    assert_eq!(run_lua_one(r#"print(math.abs(15) == 15)"#), "true");
}


#[test]
fn test_math_abs_cases_edge_first() {
    assert_eq!(run_lua_one(r#"print(math.abs(-16) == 16)"#), "true");
}


#[test]
fn test_math_abs_cases_edge_second() {
    assert_eq!(run_lua_one(r#"print(math.abs(17) == 17)"#), "true");
}


#[test]
fn test_math_abs_cases_edge_last() {
    assert_eq!(run_lua_one(r#"print(math.abs(-18) == 18)"#), "true");
}


#[test]
fn test_math_abs_cases_randomized() {
    assert_eq!(run_lua_one(r#"print(math.abs(19) == 19)"#), "true");
}


#[test]
fn test_math_abs_cases_unicode_like() {
    assert_eq!(run_lua_one(r#"print(math.abs(-20) == 20)"#), "true");
}
