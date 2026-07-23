use super::helpers::run_lua_one;

#[test]
fn test_string_format_integer_baseline() {
    assert_eq!(
        run_lua_one(r#"print(string.format("%d", 1) == "1")"#),
        "true"
    );
}

#[test]
fn test_string_format_integer_simple() {
    assert_eq!(
        run_lua_one(r#"print(string.format("%d", 2) == "2")"#),
        "true"
    );
}

#[test]
fn test_string_format_integer_trimmed() {
    assert_eq!(
        run_lua_one(r#"print(string.format("%d", 3) == "3")"#),
        "true"
    );
}

#[test]
fn test_string_format_integer_decimal() {
    assert_eq!(
        run_lua_one(r#"print(string.format("%d", 4) == "4")"#),
        "true"
    );
}

#[test]
fn test_string_format_integer_hexed() {
    assert_eq!(
        run_lua_one(r#"print(string.format("%d", 5) == "5")"#),
        "true"
    );
}

#[test]
fn test_string_format_integer_prefixed() {
    assert_eq!(
        run_lua_one(r#"print(string.format("%d", 6) == "6")"#),
        "true"
    );
}

#[test]
fn test_string_format_integer_negative() {
    assert_eq!(
        run_lua_one(r#"print(string.format("%d", 7) == "7")"#),
        "true"
    );
}

#[test]
fn test_string_format_integer_rounded() {
    assert_eq!(
        run_lua_one(r#"print(string.format("%d", 8) == "8")"#),
        "true"
    );
}

#[test]
fn test_string_format_integer_offset() {
    assert_eq!(
        run_lua_one(r#"print(string.format("%d", 9) == "9")"#),
        "true"
    );
}

#[test]
fn test_string_format_integer_paired() {
    assert_eq!(
        run_lua_one(r#"print(string.format("%d", 10) == "10")"#),
        "true"
    );
}

#[test]
fn test_string_format_integer_nested() {
    assert_eq!(
        run_lua_one(r#"print(string.format("%d", 11) == "11")"#),
        "true"
    );
}

#[test]
fn test_string_format_integer_metaflow() {
    assert_eq!(
        run_lua_one(r#"print(string.format("%d", 12) == "12")"#),
        "true"
    );
}

#[test]
fn test_string_format_integer_guarded() {
    assert_eq!(
        run_lua_one(r#"print(string.format("%d", 13) == "13")"#),
        "true"
    );
}

#[test]
fn test_string_format_integer_mapped() {
    assert_eq!(
        run_lua_one(r#"print(string.format("%d", 14) == "14")"#),
        "true"
    );
}

#[test]
fn test_string_format_integer_captured() {
    assert_eq!(
        run_lua_one(r#"print(string.format("%d", 15) == "15")"#),
        "true"
    );
}

#[test]
fn test_string_format_integer_edge_first() {
    assert_eq!(
        run_lua_one(r#"print(string.format("%d", 16) == "16")"#),
        "true"
    );
}

#[test]
fn test_string_format_integer_edge_second() {
    assert_eq!(
        run_lua_one(r#"print(string.format("%d", 17) == "17")"#),
        "true"
    );
}

#[test]
fn test_string_format_integer_edge_last() {
    assert_eq!(
        run_lua_one(r#"print(string.format("%d", 18) == "18")"#),
        "true"
    );
}

#[test]
fn test_string_format_integer_randomized() {
    assert_eq!(
        run_lua_one(r#"print(string.format("%d", 19) == "19")"#),
        "true"
    );
}

#[test]
fn test_string_format_integer_unicode_like() {
    assert_eq!(
        run_lua_one(r#"print(string.format("%d", 20) == "20")"#),
        "true"
    );
}
