use super::helpers::run_lua_one;

#[test]
fn test_rawlen_array_string_baseline() {
    assert_eq!(run_lua_one(r#"local s = ""; print(rawlen(s) == 0)"#), "true");
}


#[test]
fn test_rawlen_array_string_simple() {
    assert_eq!(run_lua_one(r#"local s = "hello"; print(rawlen(s) == 5)"#), "true");
}


#[test]
fn test_rawlen_array_string_trimmed() {
    assert_eq!(run_lua_one(r#"local s = " spaced "; print(rawlen(s) == 8)"#), "true");
}


#[test]
fn test_rawlen_array_string_decimal() {
    assert_eq!(run_lua_one(r#"local t = {1}; print(rawlen(t) == 1)"#), "true");
}


#[test]
fn test_rawlen_array_string_hexed() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; print(rawlen(t) == 3)"#), "true");
}


#[test]
fn test_rawlen_array_string_prefixed() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3,4,5}; print(rawlen(t) == 5)"#), "true");
}


#[test]
fn test_rawlen_array_string_negative() {
    assert_eq!(run_lua_one(r#"local t = {}; t[10] = 1; print(rawlen(t) == 10)"#), "true");
}


#[test]
fn test_rawlen_array_string_rounded() {
    assert_eq!(run_lua_one(r#"local t = {a = 1}; print(rawlen(t) == 0)"#), "true");
}


#[test]
fn test_rawlen_array_string_offset() {
    assert_eq!(run_lua_one(r#"local t = { [1] = "x", [2] = "y", [4] = "z"}; print(rawlen(t) == 4)"#), "true");
}


#[test]
fn test_rawlen_array_string_paired() {
    assert_eq!(run_lua_one(r#"local t = {}; rawset(t, 2, 2); print(rawlen(t) == 2)"#), "true");
}


#[test]
fn test_rawlen_array_string_nested() {
    assert_eq!(run_lua_one(r#"local s = "abcde"; print(rawlen(s) == #s)"#), "true");
}


#[test]
fn test_rawlen_array_string_metaflow() {
    assert_eq!(run_lua_one(r#"local s = "12345"; print(rawlen(s) == 5)"#), "true");
}


#[test]
fn test_rawlen_array_string_guarded() {
    assert_eq!(run_lua_one(r#"local s = "a	b"; print(rawlen(s) == 3)"#), "true");
}


#[test]
fn test_rawlen_array_string_mapped() {
    assert_eq!(run_lua_one(r#"local t = {1,2}; table.insert(t, 3); print(rawlen(t) == 3)"#), "true");
}


#[test]
fn test_rawlen_array_string_captured() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}; table.remove(t); print(rawlen(t) == 2)"#), "true");
}


#[test]
fn test_rawlen_array_string_edge_first() {
    assert_eq!(run_lua_one(r#"local t = {}; t[1] = "x"; print(rawlen(t) == 1)"#), "true");
}


#[test]
fn test_rawlen_array_string_edge_second() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3,4}; print(rawlen(t) == 4)"#), "true");
}


#[test]
fn test_rawlen_array_string_edge_last() {
    assert_eq!(run_lua_one(r#"local t = {1,2}; rawset(t, 3, 3); print(rawlen(t) == 3)"#), "true");
}


#[test]
fn test_rawlen_array_string_randomized() {
    assert_eq!(run_lua_one(r#"local s = "edge"; print(rawlen(s) == string.len(s))"#), "true");
}


#[test]
fn test_rawlen_array_string_unicode_like() {
    assert_eq!(run_lua_one(r#"local t = {n = 1}; print(type(rawlen(t)) == "number")"#), "true");
}
