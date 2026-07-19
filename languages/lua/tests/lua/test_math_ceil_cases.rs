use super::helpers::run_lua_one;

#[test]
fn test_math_ceil_cases_baseline() {
    assert_eq!(run_lua_one(r#"print(math.ceil(0.8 + 0) == 1)"#), "true");
}


#[test]
fn test_math_ceil_cases_simple() {
    assert_eq!(run_lua_one(r#"print(math.ceil(1.8 + 1) == 3)"#), "true");
}


#[test]
fn test_math_ceil_cases_trimmed() {
    assert_eq!(run_lua_one(r#"print(math.ceil(2.8 + 2) == 5)"#), "true");
}


#[test]
fn test_math_ceil_cases_decimal() {
    assert_eq!(run_lua_one(r#"print(math.ceil(3.8 + 3) == 7)"#), "true");
}


#[test]
fn test_math_ceil_cases_hexed() {
    assert_eq!(run_lua_one(r#"print(math.ceil(4.8 + 4) == 9)"#), "true");
}


#[test]
fn test_math_ceil_cases_prefixed() {
    assert_eq!(run_lua_one(r#"print(math.ceil(5.8 + 5) == 11)"#), "true");
}


#[test]
fn test_math_ceil_cases_negative() {
    assert_eq!(run_lua_one(r#"print(math.ceil(6.8 + 6) == 13)"#), "true");
}


#[test]
fn test_math_ceil_cases_rounded() {
    assert_eq!(run_lua_one(r#"print(math.ceil(7.8 + 7) == 15)"#), "true");
}


#[test]
fn test_math_ceil_cases_offset() {
    assert_eq!(run_lua_one(r#"print(math.ceil(8.8 + 8) == 17)"#), "true");
}


#[test]
fn test_math_ceil_cases_paired() {
    assert_eq!(run_lua_one(r#"print(math.ceil(9.8 + 9) == 19)"#), "true");
}


#[test]
fn test_math_ceil_cases_nested() {
    assert_eq!(run_lua_one(r#"print(math.ceil(10.8 + 10) == 21)"#), "true");
}


#[test]
fn test_math_ceil_cases_metaflow() {
    assert_eq!(run_lua_one(r#"print(math.ceil(11.8 + 11) == 23)"#), "true");
}


#[test]
fn test_math_ceil_cases_guarded() {
    assert_eq!(run_lua_one(r#"print(math.ceil(12.8 + 12) == 25)"#), "true");
}


#[test]
fn test_math_ceil_cases_mapped() {
    assert_eq!(run_lua_one(r#"print(math.ceil(13.8 + 13) == 27)"#), "true");
}


#[test]
fn test_math_ceil_cases_captured() {
    assert_eq!(run_lua_one(r#"print(math.ceil(14.8 + 14) == 29)"#), "true");
}


#[test]
fn test_math_ceil_cases_edge_first() {
    assert_eq!(run_lua_one(r#"print(math.ceil(15.8 + 15) == 31)"#), "true");
}


#[test]
fn test_math_ceil_cases_edge_second() {
    assert_eq!(run_lua_one(r#"print(math.ceil(16.8 + 16) == 33)"#), "true");
}


#[test]
fn test_math_ceil_cases_edge_last() {
    assert_eq!(run_lua_one(r#"print(math.ceil(17.8 + 17) == 35)"#), "true");
}


#[test]
fn test_math_ceil_cases_randomized() {
    assert_eq!(run_lua_one(r#"print(math.ceil(18.8 + 18) == 37)"#), "true");
}


#[test]
fn test_math_ceil_cases_unicode_like() {
    assert_eq!(run_lua_one(r#"print(math.ceil(19.8 + 19) == 39)"#), "true");
}
