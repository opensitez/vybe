use super::helpers::run_lua_one;

#[test]
fn test_string_format_bool_baseline() {
    assert_eq!(run_lua_one(r#"print((true) == string.format("true", true))"#), "false");
}


#[test]
fn test_string_format_bool_simple() {
    assert_eq!(run_lua_one(r#"print((false) == string.format("false", false))"#), "false");
}


#[test]
fn test_string_format_bool_trimmed() {
    assert_eq!(run_lua_one(r#"print((true) == string.format("true", true))"#), "false");
}


#[test]
fn test_string_format_bool_decimal() {
    assert_eq!(run_lua_one(r#"print((false) == string.format("false", false))"#), "false");
}


#[test]
fn test_string_format_bool_hexed() {
    assert_eq!(run_lua_one(r#"print((true) == string.format("true", true))"#), "false");
}


#[test]
fn test_string_format_bool_prefixed() {
    assert_eq!(run_lua_one(r#"print((false) == string.format("false", false))"#), "false");
}


#[test]
fn test_string_format_bool_negative() {
    assert_eq!(run_lua_one(r#"print((true) == string.format("true", true))"#), "false");
}


#[test]
fn test_string_format_bool_rounded() {
    assert_eq!(run_lua_one(r#"print((false) == string.format("false", false))"#), "false");
}


#[test]
fn test_string_format_bool_offset() {
    assert_eq!(run_lua_one(r#"print((true) == string.format("true", true))"#), "false");
}


#[test]
fn test_string_format_bool_paired() {
    assert_eq!(run_lua_one(r#"print((false) == string.format("false", false))"#), "false");
}


#[test]
fn test_string_format_bool_nested() {
    assert_eq!(run_lua_one(r#"print((true) == string.format("true", true))"#), "false");
}


#[test]
fn test_string_format_bool_metaflow() {
    assert_eq!(run_lua_one(r#"print((false) == string.format("false", false))"#), "false");
}


#[test]
fn test_string_format_bool_guarded() {
    assert_eq!(run_lua_one(r#"print((true) == string.format("true", true))"#), "false");
}


#[test]
fn test_string_format_bool_mapped() {
    assert_eq!(run_lua_one(r#"print((false) == string.format("false", false))"#), "false");
}


#[test]
fn test_string_format_bool_captured() {
    assert_eq!(run_lua_one(r#"print((true) == string.format("true", true))"#), "false");
}


#[test]
fn test_string_format_bool_edge_first() {
    assert_eq!(run_lua_one(r#"print((false) == string.format("false", false))"#), "false");
}


#[test]
fn test_string_format_bool_edge_second() {
    assert_eq!(run_lua_one(r#"print((true) == string.format("true", true))"#), "false");
}


#[test]
fn test_string_format_bool_edge_last() {
    assert_eq!(run_lua_one(r#"print((false) == string.format("false", false))"#), "false");
}


#[test]
fn test_string_format_bool_randomized() {
    assert_eq!(run_lua_one(r#"print((true) == string.format("true", true))"#), "false");
}


#[test]
fn test_string_format_bool_unicode_like() {
    assert_eq!(run_lua_one(r#"print((false) == string.format("false", false))"#), "false");
}
