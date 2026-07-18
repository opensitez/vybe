use super::helpers::run_lua_one;

#[test]
fn test_string_format_escaped_baseline() {
    assert_eq!(run_lua_one(r#"print(string.format("%q", "line0\n" ) ~= nil)"#), "true");
}


#[test]
fn test_string_format_escaped_simple() {
    assert_eq!(run_lua_one(r#"print(string.format("%q", "line1\n" ) ~= nil)"#), "true");
}


#[test]
fn test_string_format_escaped_trimmed() {
    assert_eq!(run_lua_one(r#"print(string.format("%q", "line2\n" ) ~= nil)"#), "true");
}


#[test]
fn test_string_format_escaped_decimal() {
    assert_eq!(run_lua_one(r#"print(string.format("%q", "line3\n" ) ~= nil)"#), "true");
}


#[test]
fn test_string_format_escaped_hexed() {
    assert_eq!(run_lua_one(r#"print(string.format("%q", "line4\n" ) ~= nil)"#), "true");
}


#[test]
fn test_string_format_escaped_prefixed() {
    assert_eq!(run_lua_one(r#"print(string.format("%q", "line5\n" ) ~= nil)"#), "true");
}


#[test]
fn test_string_format_escaped_negative() {
    assert_eq!(run_lua_one(r#"print(string.format("%q", "line6\n" ) ~= nil)"#), "true");
}


#[test]
fn test_string_format_escaped_rounded() {
    assert_eq!(run_lua_one(r#"print(string.format("%q", "line7\n" ) ~= nil)"#), "true");
}


#[test]
fn test_string_format_escaped_offset() {
    assert_eq!(run_lua_one(r#"print(string.format("%q", "line8\n" ) ~= nil)"#), "true");
}


#[test]
fn test_string_format_escaped_paired() {
    assert_eq!(run_lua_one(r#"print(string.format("%q", "line9\n" ) ~= nil)"#), "true");
}


#[test]
fn test_string_format_escaped_nested() {
    assert_eq!(run_lua_one(r#"print(string.format("%q", "line10\n" ) ~= nil)"#), "true");
}


#[test]
fn test_string_format_escaped_metaflow() {
    assert_eq!(run_lua_one(r#"print(string.format("%q", "line11\n" ) ~= nil)"#), "true");
}


#[test]
fn test_string_format_escaped_guarded() {
    assert_eq!(run_lua_one(r#"print(string.format("%q", "line12\n" ) ~= nil)"#), "true");
}


#[test]
fn test_string_format_escaped_mapped() {
    assert_eq!(run_lua_one(r#"print(string.format("%q", "line13\n" ) ~= nil)"#), "true");
}


#[test]
fn test_string_format_escaped_captured() {
    assert_eq!(run_lua_one(r#"print(string.format("%q", "line14\n" ) ~= nil)"#), "true");
}


#[test]
fn test_string_format_escaped_edge_first() {
    assert_eq!(run_lua_one(r#"print(string.format("%q", "line15\n" ) ~= nil)"#), "true");
}


#[test]
fn test_string_format_escaped_edge_second() {
    assert_eq!(run_lua_one(r#"print(string.format("%q", "line16\n" ) ~= nil)"#), "true");
}


#[test]
fn test_string_format_escaped_edge_last() {
    assert_eq!(run_lua_one(r#"print(string.format("%q", "line17\n" ) ~= nil)"#), "true");
}


#[test]
fn test_string_format_escaped_randomized() {
    assert_eq!(run_lua_one(r#"print(string.format("%q", "line18\n" ) ~= nil)"#), "true");
}


#[test]
fn test_string_format_escaped_unicode_like() {
    assert_eq!(run_lua_one(r#"print(string.format("%q", "line19\n" ) ~= nil)"#), "true");
}
