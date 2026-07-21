use super::helpers::run_lua_one;

#[test]
fn test_string_reverse_unicode_baseline() {
    assert_eq!(run_lua_one(r#"local s = "a0b"
print(string.reverse(s) == "b0a")"#), "true");
}


#[test]
fn test_string_reverse_unicode_simple() {
    assert_eq!(run_lua_one(r#"local s = "a1b"
print(string.reverse(s) == "b1a")"#), "true");
}


#[test]
fn test_string_reverse_unicode_trimmed() {
    assert_eq!(run_lua_one(r#"local s = "a2b"
print(string.reverse(s) == "b2a")"#), "true");
}


#[test]
fn test_string_reverse_unicode_decimal() {
    assert_eq!(run_lua_one(r#"local s = "a3b"
print(string.reverse(s) == "b3a")"#), "true");
}


#[test]
fn test_string_reverse_unicode_hexed() {
    assert_eq!(run_lua_one(r#"local s = "a4b"
print(string.reverse(s) == "b4a")"#), "true");
}


#[test]
fn test_string_reverse_unicode_prefixed() {
    assert_eq!(run_lua_one(r#"local s = "a5b"
print(string.reverse(s) == "b5a")"#), "true");
}


#[test]
fn test_string_reverse_unicode_negative() {
    assert_eq!(run_lua_one(r#"local s = "a6b"
print(string.reverse(s) == "b6a")"#), "true");
}


#[test]
fn test_string_reverse_unicode_rounded() {
    assert_eq!(run_lua_one(r#"local s = "a7b"
print(string.reverse(s) == "b7a")"#), "true");
}


#[test]
fn test_string_reverse_unicode_offset() {
    assert_eq!(run_lua_one(r#"local s = "a8b"
print(string.reverse(s) == "b8a")"#), "true");
}


#[test]
fn test_string_reverse_unicode_paired() {
    assert_eq!(run_lua_one(r#"local s = "a9b"
print(string.reverse(s) == "b9a")"#), "true");
}


#[test]
fn test_string_reverse_unicode_nested() {
    assert_eq!(run_lua_one(r#"local s = "a10b"
print(string.reverse(s) == "b01a")"#), "true");
}


#[test]
fn test_string_reverse_unicode_metaflow() {
    assert_eq!(run_lua_one(r#"local s = "a11b"
print(string.reverse(s) == "b11a")"#), "true");
}


#[test]
fn test_string_reverse_unicode_guarded() {
    assert_eq!(run_lua_one(r#"local s = "a12b"
print(string.reverse(s) == "b21a")"#), "true");
}


#[test]
fn test_string_reverse_unicode_mapped() {
    assert_eq!(run_lua_one(r#"local s = "a13b"
print(string.reverse(s) == "b31a")"#), "true");
}


#[test]
fn test_string_reverse_unicode_captured() {
    assert_eq!(run_lua_one(r#"local s = "a14b"
print(string.reverse(s) == "b41a")"#), "true");
}


#[test]
fn test_string_reverse_unicode_edge_first() {
    assert_eq!(run_lua_one(r#"local s = "a15b"
print(string.reverse(s) == "b51a")"#), "true");
}


#[test]
fn test_string_reverse_unicode_edge_second() {
    assert_eq!(run_lua_one(r#"local s = "a16b"
print(string.reverse(s) == "b61a")"#), "true");
}


#[test]
fn test_string_reverse_unicode_edge_last() {
    assert_eq!(run_lua_one(r#"local s = "a17b"
print(string.reverse(s) == "b71a")"#), "true");
}


#[test]
fn test_string_reverse_unicode_randomized() {
    assert_eq!(run_lua_one(r#"local s = "a18b"
print(string.reverse(s) == "b81a")"#), "true");
}


#[test]
fn test_string_reverse_unicode_unicode_like() {
    assert_eq!(run_lua_one(r#"local s = "a19b"
print(string.reverse(s) == "b91a")"#), "true");
}
