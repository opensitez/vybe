use super::helpers::run_lua_one;

#[test]
fn test_math_trig_cases_baseline() {
    assert_eq!(
        run_lua_one(r#"print(math.cos(0.1) <= 1 and math.cos(0.1) >= -1)"#),
        "true"
    );
}

#[test]
fn test_math_trig_cases_simple() {
    assert_eq!(
        run_lua_one(r#"print(math.cos(0.2) <= 1 and math.cos(0.2) >= -1)"#),
        "true"
    );
}

#[test]
fn test_math_trig_cases_trimmed() {
    assert_eq!(
        run_lua_one(r#"print(math.cos(0.3) <= 1 and math.cos(0.3) >= -1)"#),
        "true"
    );
}

#[test]
fn test_math_trig_cases_decimal() {
    assert_eq!(
        run_lua_one(r#"print(math.cos(0.4) <= 1 and math.cos(0.4) >= -1)"#),
        "true"
    );
}

#[test]
fn test_math_trig_cases_hexed() {
    assert_eq!(
        run_lua_one(r#"print(math.cos(0.5) <= 1 and math.cos(0.5) >= -1)"#),
        "true"
    );
}

#[test]
fn test_math_trig_cases_prefixed() {
    assert_eq!(
        run_lua_one(r#"print(math.cos(0.6) <= 1 and math.cos(0.6) >= -1)"#),
        "true"
    );
}

#[test]
fn test_math_trig_cases_negative() {
    assert_eq!(
        run_lua_one(r#"print(math.cos(0.7) <= 1 and math.cos(0.7) >= -1)"#),
        "true"
    );
}

#[test]
fn test_math_trig_cases_rounded() {
    assert_eq!(
        run_lua_one(r#"print(math.cos(0.8) <= 1 and math.cos(0.8) >= -1)"#),
        "true"
    );
}

#[test]
fn test_math_trig_cases_offset() {
    assert_eq!(
        run_lua_one(r#"print(math.cos(0.9) <= 1 and math.cos(0.9) >= -1)"#),
        "true"
    );
}

#[test]
fn test_math_trig_cases_paired() {
    assert_eq!(
        run_lua_one(r#"print(math.cos(1) <= 1 and math.cos(1) >= -1)"#),
        "true"
    );
}

#[test]
fn test_math_trig_cases_nested() {
    assert_eq!(
        run_lua_one(r#"print(math.cos(1.1) <= 1 and math.cos(1.1) >= -1)"#),
        "true"
    );
}

#[test]
fn test_math_trig_cases_metaflow() {
    assert_eq!(
        run_lua_one(r#"print(math.cos(1.2) <= 1 and math.cos(1.2) >= -1)"#),
        "true"
    );
}

#[test]
fn test_math_trig_cases_guarded() {
    assert_eq!(
        run_lua_one(r#"print(math.cos(1.3) <= 1 and math.cos(1.3) >= -1)"#),
        "true"
    );
}

#[test]
fn test_math_trig_cases_mapped() {
    assert_eq!(
        run_lua_one(r#"print(math.cos(1.4) <= 1 and math.cos(1.4) >= -1)"#),
        "true"
    );
}

#[test]
fn test_math_trig_cases_captured() {
    assert_eq!(
        run_lua_one(r#"print(math.cos(1.5) <= 1 and math.cos(1.5) >= -1)"#),
        "true"
    );
}

#[test]
fn test_math_trig_cases_edge_first() {
    assert_eq!(
        run_lua_one(r#"print(math.cos(1.6) <= 1 and math.cos(1.6) >= -1)"#),
        "true"
    );
}

#[test]
fn test_math_trig_cases_edge_second() {
    assert_eq!(
        run_lua_one(r#"print(math.cos(1.7) <= 1 and math.cos(1.7) >= -1)"#),
        "true"
    );
}

#[test]
fn test_math_trig_cases_edge_last() {
    assert_eq!(
        run_lua_one(r#"print(math.cos(1.8) <= 1 and math.cos(1.8) >= -1)"#),
        "true"
    );
}

#[test]
fn test_math_trig_cases_randomized() {
    assert_eq!(
        run_lua_one(r#"print(math.cos(1.9) <= 1 and math.cos(1.9) >= -1)"#),
        "true"
    );
}

#[test]
fn test_math_trig_cases_unicode_like() {
    assert_eq!(
        run_lua_one(r#"print(math.cos(2) <= 1 and math.cos(2) >= -1)"#),
        "true"
    );
}
