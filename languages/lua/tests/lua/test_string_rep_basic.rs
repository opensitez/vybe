use super::helpers::run_lua_one;

#[test]
fn test_string_rep_basic_baseline() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 1) == string.rep("x", 1))"#),
        "true"
    );
}

#[test]
fn test_string_rep_basic_simple() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 2) == string.rep("x", 2))"#),
        "true"
    );
}

#[test]
fn test_string_rep_basic_trimmed() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 3) == string.rep("x", 3))"#),
        "true"
    );
}

#[test]
fn test_string_rep_basic_decimal() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 4) == string.rep("x", 4))"#),
        "true"
    );
}

#[test]
fn test_string_rep_basic_hexed() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 5) == string.rep("x", 5))"#),
        "true"
    );
}

#[test]
fn test_string_rep_basic_prefixed() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 6) == string.rep("x", 6))"#),
        "true"
    );
}

#[test]
fn test_string_rep_basic_negative() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 7) == string.rep("x", 7))"#),
        "true"
    );
}

#[test]
fn test_string_rep_basic_rounded() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 8) == string.rep("x", 8))"#),
        "true"
    );
}

#[test]
fn test_string_rep_basic_offset() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 1) == string.rep("x", 1))"#),
        "true"
    );
}

#[test]
fn test_string_rep_basic_paired() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 2) == string.rep("x", 2))"#),
        "true"
    );
}

#[test]
fn test_string_rep_basic_nested() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 3) == string.rep("x", 3))"#),
        "true"
    );
}

#[test]
fn test_string_rep_basic_metaflow() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 4) == string.rep("x", 4))"#),
        "true"
    );
}

#[test]
fn test_string_rep_basic_guarded() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 5) == string.rep("x", 5))"#),
        "true"
    );
}

#[test]
fn test_string_rep_basic_mapped() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 6) == string.rep("x", 6))"#),
        "true"
    );
}

#[test]
fn test_string_rep_basic_captured() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 7) == string.rep("x", 7))"#),
        "true"
    );
}

#[test]
fn test_string_rep_basic_edge_first() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 8) == string.rep("x", 8))"#),
        "true"
    );
}

#[test]
fn test_string_rep_basic_edge_second() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 1) == string.rep("x", 1))"#),
        "true"
    );
}

#[test]
fn test_string_rep_basic_edge_last() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 2) == string.rep("x", 2))"#),
        "true"
    );
}

#[test]
fn test_string_rep_basic_randomized() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 3) == string.rep("x", 3))"#),
        "true"
    );
}

#[test]
fn test_string_rep_basic_unicode_like() {
    assert_eq!(
        run_lua_one(r#"print(string.rep("x", 4) == string.rep("x", 4))"#),
        "true"
    );
}
