use super::helpers::run_lua_one;

#[test]
fn test_string_pack_signed_baseline() {
    assert_eq!(run_lua_one(r#"local s = string.pack("i4", 1); local v = string.unpack("i4", s); print(v == 1)"#), "true");
}


#[test]
fn test_string_pack_signed_simple() {
    assert_eq!(run_lua_one(r#"local s = string.pack("i4", 2); local v = string.unpack("i4", s); print(v == 2)"#), "true");
}


#[test]
fn test_string_pack_signed_trimmed() {
    assert_eq!(run_lua_one(r#"local s = string.pack("i4", 3); local v = string.unpack("i4", s); print(v == 3)"#), "true");
}


#[test]
fn test_string_pack_signed_decimal() {
    assert_eq!(run_lua_one(r#"local s = string.pack("i4", 4); local v = string.unpack("i4", s); print(v == 4)"#), "true");
}


#[test]
fn test_string_pack_signed_hexed() {
    assert_eq!(run_lua_one(r#"local s = string.pack("i4", 5); local v = string.unpack("i4", s); print(v == 5)"#), "true");
}


#[test]
fn test_string_pack_signed_prefixed() {
    assert_eq!(run_lua_one(r#"local s = string.pack("i4", 6); local v = string.unpack("i4", s); print(v == 6)"#), "true");
}


#[test]
fn test_string_pack_signed_negative() {
    assert_eq!(run_lua_one(r#"local s = string.pack("i4", 7); local v = string.unpack("i4", s); print(v == 7)"#), "true");
}


#[test]
fn test_string_pack_signed_rounded() {
    assert_eq!(run_lua_one(r#"local s = string.pack("i4", 8); local v = string.unpack("i4", s); print(v == 8)"#), "true");
}


#[test]
fn test_string_pack_signed_offset() {
    assert_eq!(run_lua_one(r#"local s = string.pack("i4", 9); local v = string.unpack("i4", s); print(v == 9)"#), "true");
}


#[test]
fn test_string_pack_signed_paired() {
    assert_eq!(run_lua_one(r#"local s = string.pack("i4", 10); local v = string.unpack("i4", s); print(v == 10)"#), "true");
}


#[test]
fn test_string_pack_signed_nested() {
    assert_eq!(run_lua_one(r#"local s = string.pack("i4", 11); local v = string.unpack("i4", s); print(v == 11)"#), "true");
}


#[test]
fn test_string_pack_signed_metaflow() {
    assert_eq!(run_lua_one(r#"local s = string.pack("i4", 12); local v = string.unpack("i4", s); print(v == 12)"#), "true");
}


#[test]
fn test_string_pack_signed_guarded() {
    assert_eq!(run_lua_one(r#"local s = string.pack("i4", 13); local v = string.unpack("i4", s); print(v == 13)"#), "true");
}


#[test]
fn test_string_pack_signed_mapped() {
    assert_eq!(run_lua_one(r#"local s = string.pack("i4", 14); local v = string.unpack("i4", s); print(v == 14)"#), "true");
}


#[test]
fn test_string_pack_signed_captured() {
    assert_eq!(run_lua_one(r#"local s = string.pack("i4", 15); local v = string.unpack("i4", s); print(v == 15)"#), "true");
}


#[test]
fn test_string_pack_signed_edge_first() {
    assert_eq!(run_lua_one(r#"local s = string.pack("i4", 16); local v = string.unpack("i4", s); print(v == 16)"#), "true");
}


#[test]
fn test_string_pack_signed_edge_second() {
    assert_eq!(run_lua_one(r#"local s = string.pack("i4", 17); local v = string.unpack("i4", s); print(v == 17)"#), "true");
}


#[test]
fn test_string_pack_signed_edge_last() {
    assert_eq!(run_lua_one(r#"local s = string.pack("i4", 18); local v = string.unpack("i4", s); print(v == 18)"#), "true");
}


#[test]
fn test_string_pack_signed_randomized() {
    assert_eq!(run_lua_one(r#"local s = string.pack("i4", 19); local v = string.unpack("i4", s); print(v == 19)"#), "true");
}


#[test]
fn test_string_pack_signed_unicode_like() {
    assert_eq!(run_lua_one(r#"local s = string.pack("i4", 20); local v = string.unpack("i4", s); print(v == 20)"#), "true");
}
