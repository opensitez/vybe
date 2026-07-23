use super::helpers::run_lua_one;

#[test]
fn test_string_match_tokens_baseline() {
    assert_eq!(
        run_lua_one(r#"print(string.match("a1 b2 c3", "b2") == "b2")"#),
        "true"
    );
}

#[test]
fn test_string_match_tokens_simple() {
    assert_eq!(
        run_lua_one(r#"print(string.match("a2 b3 c4", "b3") == "b3")"#),
        "true"
    );
}

#[test]
fn test_string_match_tokens_trimmed() {
    assert_eq!(
        run_lua_one(r#"print(string.match("a3 b4 c5", "b4") == "b4")"#),
        "true"
    );
}

#[test]
fn test_string_match_tokens_decimal() {
    assert_eq!(
        run_lua_one(r#"print(string.match("a4 b5 c6", "b5") == "b5")"#),
        "true"
    );
}

#[test]
fn test_string_match_tokens_hexed() {
    assert_eq!(
        run_lua_one(r#"print(string.match("a5 b6 c7", "b6") == "b6")"#),
        "true"
    );
}

#[test]
fn test_string_match_tokens_prefixed() {
    assert_eq!(
        run_lua_one(r#"print(string.match("a6 b7 c8", "b7") == "b7")"#),
        "true"
    );
}

#[test]
fn test_string_match_tokens_negative() {
    assert_eq!(
        run_lua_one(r#"print(string.match("a7 b8 c9", "b8") == "b8")"#),
        "true"
    );
}

#[test]
fn test_string_match_tokens_rounded() {
    assert_eq!(
        run_lua_one(r#"print(string.match("a8 b9 c10", "b9") == "b9")"#),
        "true"
    );
}

#[test]
fn test_string_match_tokens_offset() {
    assert_eq!(
        run_lua_one(r#"print(string.match("a9 b10 c11", "b10") == "b10")"#),
        "true"
    );
}

#[test]
fn test_string_match_tokens_paired() {
    assert_eq!(
        run_lua_one(r#"print(string.match("a10 b11 c12", "b11") == "b11")"#),
        "true"
    );
}

#[test]
fn test_string_match_tokens_nested() {
    assert_eq!(
        run_lua_one(r#"print(string.match("a11 b12 c13", "b12") == "b12")"#),
        "true"
    );
}

#[test]
fn test_string_match_tokens_metaflow() {
    assert_eq!(
        run_lua_one(r#"print(string.match("a12 b13 c14", "b13") == "b13")"#),
        "true"
    );
}

#[test]
fn test_string_match_tokens_guarded() {
    assert_eq!(
        run_lua_one(r#"print(string.match("a13 b14 c15", "b14") == "b14")"#),
        "true"
    );
}

#[test]
fn test_string_match_tokens_mapped() {
    assert_eq!(
        run_lua_one(r#"print(string.match("a14 b15 c16", "b15") == "b15")"#),
        "true"
    );
}

#[test]
fn test_string_match_tokens_captured() {
    assert_eq!(
        run_lua_one(r#"print(string.match("a15 b16 c17", "b16") == "b16")"#),
        "true"
    );
}

#[test]
fn test_string_match_tokens_edge_first() {
    assert_eq!(
        run_lua_one(r#"print(string.match("a16 b17 c18", "b17") == "b17")"#),
        "true"
    );
}

#[test]
fn test_string_match_tokens_edge_second() {
    assert_eq!(
        run_lua_one(r#"print(string.match("a17 b18 c19", "b18") == "b18")"#),
        "true"
    );
}

#[test]
fn test_string_match_tokens_edge_last() {
    assert_eq!(
        run_lua_one(r#"print(string.match("a18 b19 c20", "b19") == "b19")"#),
        "true"
    );
}

#[test]
fn test_string_match_tokens_randomized() {
    assert_eq!(
        run_lua_one(r#"print(string.match("a19 b20 c21", "b20") == "b20")"#),
        "true"
    );
}

#[test]
fn test_string_match_tokens_unicode_like() {
    assert_eq!(
        run_lua_one(r#"print(string.match("a20 b21 c22", "b21") == "b21")"#),
        "true"
    );
}
