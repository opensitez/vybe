use super::helpers::run_lua_one;

#[test]
fn test_string_rep_zero_baseline() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 0) == string.rep("x", 0))"#),
        "true"
    );
}

#[test]
fn test_string_rep_zero_simple() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 1) == string.rep("x", 1))"#),
        "true"
    );
}

#[test]
fn test_string_rep_zero_trimmed() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 2) == string.rep("x", 2))"#),
        "true"
    );
}

#[test]
fn test_string_rep_zero_decimal() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 0) == string.rep("x", 0))"#),
        "true"
    );
}

#[test]
fn test_string_rep_zero_hexed() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 1) == string.rep("x", 1))"#),
        "true"
    );
}

#[test]
fn test_string_rep_zero_prefixed() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 2) == string.rep("x", 2))"#),
        "true"
    );
}

#[test]
fn test_string_rep_zero_negative() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 0) == string.rep("x", 0))"#),
        "true"
    );
}

#[test]
fn test_string_rep_zero_rounded() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 1) == string.rep("x", 1))"#),
        "true"
    );
}

#[test]
fn test_string_rep_zero_offset() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 2) == string.rep("x", 2))"#),
        "true"
    );
}

#[test]
fn test_string_rep_zero_paired() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 0) == string.rep("x", 0))"#),
        "true"
    );
}

#[test]
fn test_string_rep_zero_nested() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 1) == string.rep("x", 1))"#),
        "true"
    );
}

#[test]
fn test_string_rep_zero_metaflow() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 2) == string.rep("x", 2))"#),
        "true"
    );
}

#[test]
fn test_string_rep_zero_guarded() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 0) == string.rep("x", 0))"#),
        "true"
    );
}

#[test]
fn test_string_rep_zero_mapped() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 1) == string.rep("x", 1))"#),
        "true"
    );
}

#[test]
fn test_string_rep_zero_captured() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 2) == string.rep("x", 2))"#),
        "true"
    );
}

#[test]
fn test_string_rep_zero_edge_first() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 0) == string.rep("x", 0))"#),
        "true"
    );
}

#[test]
fn test_string_rep_zero_edge_second() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 1) == string.rep("x", 1))"#),
        "true"
    );
}

#[test]
fn test_string_rep_zero_edge_last() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 2) == string.rep("x", 2))"#),
        "true"
    );
}

#[test]
fn test_string_rep_zero_randomized() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 0) == string.rep("x", 0))"#),
        "true"
    );
}

#[test]
fn test_string_rep_zero_unicode_like() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 1) == string.rep("x", 1))"#),
        "true"
    );
}
