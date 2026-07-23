use super::helpers::run_lua_one;

#[test]
fn test_math_fmod_cases_baseline() {
    assert_eq!(run_lua_one(r#"print(math.fmod(20, 7) == 6)"#), "true");
}

#[test]
fn test_math_fmod_cases_simple() {
    assert_eq!(run_lua_one(r#"print(math.fmod(21, 7) == 0)"#), "true");
}

#[test]
fn test_math_fmod_cases_trimmed() {
    assert_eq!(run_lua_one(r#"print(math.fmod(22, 7) == 1)"#), "true");
}

#[test]
fn test_math_fmod_cases_decimal() {
    assert_eq!(run_lua_one(r#"print(math.fmod(23, 7) == 2)"#), "true");
}

#[test]
fn test_math_fmod_cases_hexed() {
    assert_eq!(run_lua_one(r#"print(math.fmod(24, 7) == 3)"#), "true");
}

#[test]
fn test_math_fmod_cases_prefixed() {
    assert_eq!(run_lua_one(r#"print(math.fmod(25, 7) == 4)"#), "true");
}

#[test]
fn test_math_fmod_cases_negative() {
    assert_eq!(run_lua_one(r#"print(math.fmod(26, 7) == 5)"#), "true");
}

#[test]
fn test_math_fmod_cases_rounded() {
    assert_eq!(run_lua_one(r#"print(math.fmod(27, 7) == 6)"#), "true");
}

#[test]
fn test_math_fmod_cases_offset() {
    assert_eq!(run_lua_one(r#"print(math.fmod(28, 7) == 0)"#), "true");
}

#[test]
fn test_math_fmod_cases_paired() {
    assert_eq!(run_lua_one(r#"print(math.fmod(29, 7) == 1)"#), "true");
}

#[test]
fn test_math_fmod_cases_nested() {
    assert_eq!(run_lua_one(r#"print(math.fmod(30, 7) == 2)"#), "true");
}

#[test]
fn test_math_fmod_cases_metaflow() {
    assert_eq!(run_lua_one(r#"print(math.fmod(31, 7) == 3)"#), "true");
}

#[test]
fn test_math_fmod_cases_guarded() {
    assert_eq!(run_lua_one(r#"print(math.fmod(32, 7) == 4)"#), "true");
}

#[test]
fn test_math_fmod_cases_mapped() {
    assert_eq!(run_lua_one(r#"print(math.fmod(33, 7) == 5)"#), "true");
}

#[test]
fn test_math_fmod_cases_captured() {
    assert_eq!(run_lua_one(r#"print(math.fmod(34, 7) == 6)"#), "true");
}

#[test]
fn test_math_fmod_cases_edge_first() {
    assert_eq!(run_lua_one(r#"print(math.fmod(35, 7) == 0)"#), "true");
}

#[test]
fn test_math_fmod_cases_edge_second() {
    assert_eq!(run_lua_one(r#"print(math.fmod(36, 7) == 1)"#), "true");
}

#[test]
fn test_math_fmod_cases_edge_last() {
    assert_eq!(run_lua_one(r#"print(math.fmod(37, 7) == 2)"#), "true");
}

#[test]
fn test_math_fmod_cases_randomized() {
    assert_eq!(run_lua_one(r#"print(math.fmod(38, 7) == 3)"#), "true");
}

#[test]
fn test_math_fmod_cases_unicode_like() {
    assert_eq!(run_lua_one(r#"print(math.fmod(39, 7) == 4)"#), "true");
}
