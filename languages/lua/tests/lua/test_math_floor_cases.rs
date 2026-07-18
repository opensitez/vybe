use super::helpers::run_lua_one;

#[test]
fn test_math_floor_cases_baseline() {
    assert_eq!(run_lua_one(r#"print(math.floor(0.2 + 0) == math.floor(0.2))"#), "true");
}


#[test]
fn test_math_floor_cases_simple() {
    assert_eq!(run_lua_one(r#"print(math.floor(1.2 + 1) == math.floor(2.2))"#), "true");
}


#[test]
fn test_math_floor_cases_trimmed() {
    assert_eq!(run_lua_one(r#"print(math.floor(2.2 + 2) == math.floor(4.2))"#), "true");
}


#[test]
fn test_math_floor_cases_decimal() {
    assert_eq!(run_lua_one(r#"print(math.floor(3.2 + 3) == math.floor(6.2))"#), "true");
}


#[test]
fn test_math_floor_cases_hexed() {
    assert_eq!(run_lua_one(r#"print(math.floor(4.2 + 4) == math.floor(8.2))"#), "true");
}


#[test]
fn test_math_floor_cases_prefixed() {
    assert_eq!(run_lua_one(r#"print(math.floor(5.2 + 5) == math.floor(10.2))"#), "true");
}


#[test]
fn test_math_floor_cases_negative() {
    assert_eq!(run_lua_one(r#"print(math.floor(6.2 + 6) == math.floor(12.2))"#), "true");
}


#[test]
fn test_math_floor_cases_rounded() {
    assert_eq!(run_lua_one(r#"print(math.floor(7.2 + 7) == math.floor(14.2))"#), "true");
}


#[test]
fn test_math_floor_cases_offset() {
    assert_eq!(run_lua_one(r#"print(math.floor(8.2 + 8) == math.floor(16.2))"#), "true");
}


#[test]
fn test_math_floor_cases_paired() {
    assert_eq!(run_lua_one(r#"print(math.floor(9.2 + 9) == math.floor(18.2))"#), "true");
}


#[test]
fn test_math_floor_cases_nested() {
    assert_eq!(run_lua_one(r#"print(math.floor(10.2 + 10) == math.floor(20.2))"#), "true");
}


#[test]
fn test_math_floor_cases_metaflow() {
    assert_eq!(run_lua_one(r#"print(math.floor(11.2 + 11) == math.floor(22.2))"#), "true");
}


#[test]
fn test_math_floor_cases_guarded() {
    assert_eq!(run_lua_one(r#"print(math.floor(12.2 + 12) == math.floor(24.2))"#), "true");
}


#[test]
fn test_math_floor_cases_mapped() {
    assert_eq!(run_lua_one(r#"print(math.floor(13.2 + 13) == math.floor(26.2))"#), "true");
}


#[test]
fn test_math_floor_cases_captured() {
    assert_eq!(run_lua_one(r#"print(math.floor(14.2 + 14) == math.floor(28.2))"#), "true");
}


#[test]
fn test_math_floor_cases_edge_first() {
    assert_eq!(run_lua_one(r#"print(math.floor(15.2 + 15) == math.floor(30.2))"#), "true");
}


#[test]
fn test_math_floor_cases_edge_second() {
    assert_eq!(run_lua_one(r#"print(math.floor(16.2 + 16) == math.floor(32.2))"#), "true");
}


#[test]
fn test_math_floor_cases_edge_last() {
    assert_eq!(run_lua_one(r#"print(math.floor(17.2 + 17) == math.floor(34.2))"#), "true");
}


#[test]
fn test_math_floor_cases_randomized() {
    assert_eq!(run_lua_one(r#"print(math.floor(18.2 + 18) == math.floor(36.2))"#), "true");
}


#[test]
fn test_math_floor_cases_unicode_like() {
    assert_eq!(run_lua_one(r#"print(math.floor(19.2 + 19) == math.floor(38.2))"#), "true");
}
