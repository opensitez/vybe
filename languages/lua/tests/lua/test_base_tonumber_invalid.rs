use super::helpers::run_lua_one;

#[test]
fn test_tonumber_invalid_baseline() {
    assert_eq!(run_lua_one(r#"print(tonumber("") == nil)"#), "true");
}


#[test]
fn test_tonumber_invalid_simple() {
    assert_eq!(run_lua_one(r#"print(tonumber("  ") == nil)"#), "true");
}


#[test]
fn test_tonumber_invalid_trimmed() {
    assert_eq!(run_lua_one(r#"print(tonumber("abc") == nil)"#), "true");
}


#[test]
fn test_tonumber_invalid_decimal() {
    assert_eq!(run_lua_one(r#"print(tonumber("10x") == nil)"#), "true");
}


#[test]
fn test_tonumber_invalid_hexed() {
    assert_eq!(run_lua_one(r#"print(tonumber("0x") == nil)"#), "true");
}


#[test]
fn test_tonumber_invalid_prefixed() {
    assert_eq!(run_lua_one(r#"print(tonumber("0b2") == nil)"#), "true");
}


#[test]
fn test_tonumber_invalid_negative() {
    assert_eq!(run_lua_one(r#"print(tonumber("1..2") == nil)"#), "true");
}


#[test]
fn test_tonumber_invalid_rounded() {
    assert_eq!(run_lua_one(r#"print(tonumber("a1") == nil)"#), "true");
}


#[test]
fn test_tonumber_invalid_offset() {
    assert_eq!(run_lua_one(r#"print(tonumber("one") == nil)"#), "true");
}


#[test]
fn test_tonumber_invalid_paired() {
    assert_eq!(run_lua_one(r#"print(tonumber("-") == nil)"#), "true");
}


#[test]
fn test_tonumber_invalid_nested() {
    assert_eq!(run_lua_one(r#"print(tonumber("--2") == nil)"#), "true");
}


#[test]
fn test_tonumber_invalid_metaflow() {
    assert_eq!(run_lua_one(r#"print(tonumber("nan%") == nil)"#), "true");
}


#[test]
fn test_tonumber_invalid_guarded() {
    assert_eq!(run_lua_one(r#"print(tonumber("1.2.3") == nil)"#), "true");
}


#[test]
fn test_tonumber_invalid_mapped() {
    assert_eq!(run_lua_one(r#"print(tonumber("[]") == nil)"#), "true");
}


#[test]
fn test_tonumber_invalid_captured() {
    assert_eq!(run_lua_one(r#"print(tonumber("{}") == nil)"#), "true");
}


#[test]
fn test_tonumber_invalid_edge_first() {
    assert_eq!(run_lua_one(r#"print(tonumber("nil") == nil)"#), "true");
}


#[test]
fn test_tonumber_invalid_edge_second() {
    assert_eq!(run_lua_one(r#"print(tonumber("true") == nil)"#), "true");
}


#[test]
fn test_tonumber_invalid_edge_last() {
    assert_eq!(run_lua_one(r#"print(tonumber("++1") == nil)"#), "true");
}


#[test]
fn test_tonumber_invalid_randomized() {
    assert_eq!(run_lua_one(r#"print(tonumber("--") == nil)"#), "true");
}


#[test]
fn test_tonumber_invalid_unicode_like() {
    assert_eq!(run_lua_one(r#"print(tonumber("1e") == nil)"#), "true");
}
