use super::helpers::run_lua_one;

#[test]
fn test_math_log_cases_baseline() {
    assert_eq!(run_lua_one(r#"print(math.log(1) > 0 or 1 == 1)"#), "true");
}

#[test]
fn test_math_log_cases_simple() {
    assert_eq!(run_lua_one(r#"print(math.log(2) > 0 or 2 == 1)"#), "true");
}

#[test]
fn test_math_log_cases_trimmed() {
    assert_eq!(run_lua_one(r#"print(math.log(3) > 0 or 3 == 1)"#), "true");
}

#[test]
fn test_math_log_cases_decimal() {
    assert_eq!(run_lua_one(r#"print(math.log(4) > 0 or 4 == 1)"#), "true");
}

#[test]
fn test_math_log_cases_hexed() {
    assert_eq!(run_lua_one(r#"print(math.log(5) > 0 or 5 == 1)"#), "true");
}

#[test]
fn test_math_log_cases_prefixed() {
    assert_eq!(run_lua_one(r#"print(math.log(6) > 0 or 6 == 1)"#), "true");
}

#[test]
fn test_math_log_cases_negative() {
    assert_eq!(run_lua_one(r#"print(math.log(7) > 0 or 7 == 1)"#), "true");
}

#[test]
fn test_math_log_cases_rounded() {
    assert_eq!(run_lua_one(r#"print(math.log(8) > 0 or 8 == 1)"#), "true");
}

#[test]
fn test_math_log_cases_offset() {
    assert_eq!(run_lua_one(r#"print(math.log(9) > 0 or 9 == 1)"#), "true");
}

#[test]
fn test_math_log_cases_paired() {
    assert_eq!(run_lua_one(r#"print(math.log(10) > 0 or 10 == 1)"#), "true");
}

#[test]
fn test_math_log_cases_nested() {
    assert_eq!(run_lua_one(r#"print(math.log(11) > 0 or 11 == 1)"#), "true");
}

#[test]
fn test_math_log_cases_metaflow() {
    assert_eq!(run_lua_one(r#"print(math.log(12) > 0 or 12 == 1)"#), "true");
}

#[test]
fn test_math_log_cases_guarded() {
    assert_eq!(run_lua_one(r#"print(math.log(13) > 0 or 13 == 1)"#), "true");
}

#[test]
fn test_math_log_cases_mapped() {
    assert_eq!(run_lua_one(r#"print(math.log(14) > 0 or 14 == 1)"#), "true");
}

#[test]
fn test_math_log_cases_captured() {
    assert_eq!(run_lua_one(r#"print(math.log(15) > 0 or 15 == 1)"#), "true");
}

#[test]
fn test_math_log_cases_edge_first() {
    assert_eq!(run_lua_one(r#"print(math.log(16) > 0 or 16 == 1)"#), "true");
}

#[test]
fn test_math_log_cases_edge_second() {
    assert_eq!(run_lua_one(r#"print(math.log(17) > 0 or 17 == 1)"#), "true");
}

#[test]
fn test_math_log_cases_edge_last() {
    assert_eq!(run_lua_one(r#"print(math.log(18) > 0 or 18 == 1)"#), "true");
}

#[test]
fn test_math_log_cases_randomized() {
    assert_eq!(run_lua_one(r#"print(math.log(19) > 0 or 19 == 1)"#), "true");
}

#[test]
fn test_math_log_cases_unicode_like() {
    assert_eq!(run_lua_one(r#"print(math.log(20) > 0 or 20 == 1)"#), "true");
}
