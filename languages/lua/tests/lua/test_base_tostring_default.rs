use super::helpers::run_lua_one;

#[test]
fn test_tostring_default_baseline() {
    assert_eq!(run_lua_one(r#"print(tostring(0) == "0")"#), "true");
}


#[test]
fn test_tostring_default_simple() {
    assert_eq!(run_lua_one(r#"print(tostring(1) == "1")"#), "true");
}


#[test]
fn test_tostring_default_trimmed() {
    assert_eq!(run_lua_one(r#"print(tostring(-2) == "-2")"#), "true");
}


#[test]
fn test_tostring_default_decimal() {
    assert_eq!(run_lua_one(r#"print(tostring(3.5) == "3.5")"#), "true");
}


#[test]
fn test_tostring_default_hexed() {
    assert_eq!(run_lua_one(r#"print(tostring(true) == "true")"#), "true");
}


#[test]
fn test_tostring_default_prefixed() {
    assert_eq!(run_lua_one(r#"print(tostring(false) == "false")"#), "true");
}


#[test]
fn test_tostring_default_negative() {
    assert_eq!(run_lua_one(r#"print(tostring(nil) == "nil")"#), "true");
}


#[test]
fn test_tostring_default_rounded() {
    assert_eq!(run_lua_one(r#"print(tostring("x") == "x")"#), "true");
}


#[test]
fn test_tostring_default_offset() {
    assert_eq!(run_lua_one(r#"print(tostring(" spaced ") == " spaced ")"#), "true");
}


#[test]
fn test_tostring_default_paired() {
    assert_eq!(run_lua_one(r#"print(tostring("A\nB") == "A\nB")"#), "true");
}


#[test]
fn test_tostring_default_nested() {
    assert_eq!(run_lua_one(r#"print(tostring("-") == "-")"#), "true");
}


#[test]
fn test_tostring_default_metaflow() {
    assert_eq!(run_lua_one(r#"print(tostring("") == "")"#), "true");
}


#[test]
fn test_tostring_default_guarded() {
    assert_eq!(run_lua_one(r#"print(tostring("42") == "42")"#), "true");
}


#[test]
fn test_tostring_default_mapped() {
    assert_eq!(run_lua_one(r#"print(tostring("0x1") == "0x1")"#), "true");
}


#[test]
fn test_tostring_default_captured() {
    assert_eq!(run_lua_one(r#"print(tostring("table") == "table")"#), "true");
}


#[test]
fn test_tostring_default_edge_first() {
    assert_eq!(run_lua_one(r#"print(tostring("function") == "function")"#), "true");
}


#[test]
fn test_tostring_default_edge_second() {
    assert_eq!(run_lua_one(r#"print(tostring("123") == "123")"#), "true");
}


#[test]
fn test_tostring_default_edge_last() {
    assert_eq!(run_lua_one(r#"print(tostring("alpha") == "alpha")"#), "true");
}


#[test]
fn test_tostring_default_randomized() {
    assert_eq!(run_lua_one(r#"print(tostring("beta") == "beta")"#), "true");
}


#[test]
fn test_tostring_default_unicode_like() {
    assert_eq!(run_lua_one(r#"print(tostring("gamma") == "gamma")"#), "true");
}
