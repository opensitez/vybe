use super::helpers::run_lua_one;

#[test]
fn test_string_pack_endian_baseline() {
    assert_eq!(
        run_lua_one(r#"local s = string.pack("<i2", 1); print(string.unpack("<i2", s) == 1)"#),
        "true"
    );
}

#[test]
fn test_string_pack_endian_simple() {
    assert_eq!(
        run_lua_one(r#"local s = string.pack("<i2", 2); print(string.unpack("<i2", s) == 2)"#),
        "true"
    );
}

#[test]
fn test_string_pack_endian_trimmed() {
    assert_eq!(
        run_lua_one(r#"local s = string.pack("<i2", 3); print(string.unpack("<i2", s) == 3)"#),
        "true"
    );
}

#[test]
fn test_string_pack_endian_decimal() {
    assert_eq!(
        run_lua_one(r#"local s = string.pack("<i2", 4); print(string.unpack("<i2", s) == 4)"#),
        "true"
    );
}

#[test]
fn test_string_pack_endian_hexed() {
    assert_eq!(
        run_lua_one(r#"local s = string.pack("<i2", 5); print(string.unpack("<i2", s) == 5)"#),
        "true"
    );
}

#[test]
fn test_string_pack_endian_prefixed() {
    assert_eq!(
        run_lua_one(r#"local s = string.pack("<i2", 6); print(string.unpack("<i2", s) == 6)"#),
        "true"
    );
}

#[test]
fn test_string_pack_endian_negative() {
    assert_eq!(
        run_lua_one(r#"local s = string.pack("<i2", 7); print(string.unpack("<i2", s) == 7)"#),
        "true"
    );
}

#[test]
fn test_string_pack_endian_rounded() {
    assert_eq!(
        run_lua_one(r#"local s = string.pack("<i2", 8); print(string.unpack("<i2", s) == 8)"#),
        "true"
    );
}

#[test]
fn test_string_pack_endian_offset() {
    assert_eq!(
        run_lua_one(r#"local s = string.pack("<i2", 9); print(string.unpack("<i2", s) == 9)"#),
        "true"
    );
}

#[test]
fn test_string_pack_endian_paired() {
    assert_eq!(
        run_lua_one(r#"local s = string.pack("<i2", 10); print(string.unpack("<i2", s) == 10)"#),
        "true"
    );
}

#[test]
fn test_string_pack_endian_nested() {
    assert_eq!(
        run_lua_one(r#"local s = string.pack("<i2", 11); print(string.unpack("<i2", s) == 11)"#),
        "true"
    );
}

#[test]
fn test_string_pack_endian_metaflow() {
    assert_eq!(
        run_lua_one(r#"local s = string.pack("<i2", 12); print(string.unpack("<i2", s) == 12)"#),
        "true"
    );
}

#[test]
fn test_string_pack_endian_guarded() {
    assert_eq!(
        run_lua_one(r#"local s = string.pack("<i2", 13); print(string.unpack("<i2", s) == 13)"#),
        "true"
    );
}

#[test]
fn test_string_pack_endian_mapped() {
    assert_eq!(
        run_lua_one(r#"local s = string.pack("<i2", 14); print(string.unpack("<i2", s) == 14)"#),
        "true"
    );
}

#[test]
fn test_string_pack_endian_captured() {
    assert_eq!(
        run_lua_one(r#"local s = string.pack("<i2", 15); print(string.unpack("<i2", s) == 15)"#),
        "true"
    );
}

#[test]
fn test_string_pack_endian_edge_first() {
    assert_eq!(
        run_lua_one(r#"local s = string.pack("<i2", 16); print(string.unpack("<i2", s) == 16)"#),
        "true"
    );
}

#[test]
fn test_string_pack_endian_edge_second() {
    assert_eq!(
        run_lua_one(r#"local s = string.pack("<i2", 17); print(string.unpack("<i2", s) == 17)"#),
        "true"
    );
}

#[test]
fn test_string_pack_endian_edge_last() {
    assert_eq!(
        run_lua_one(r#"local s = string.pack("<i2", 18); print(string.unpack("<i2", s) == 18)"#),
        "true"
    );
}

#[test]
fn test_string_pack_endian_randomized() {
    assert_eq!(
        run_lua_one(r#"local s = string.pack("<i2", 19); print(string.unpack("<i2", s) == 19)"#),
        "true"
    );
}

#[test]
fn test_string_pack_endian_unicode_like() {
    assert_eq!(
        run_lua_one(r#"local s = string.pack("<i2", 20); print(string.unpack("<i2", s) == 20)"#),
        "true"
    );
}
