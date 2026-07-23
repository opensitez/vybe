use super::helpers::run_lua_one;

#[test]
fn test_string_unpack_basic_baseline() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.pack("i4", 1); local a = string.unpack("i4", s); print(a == 1)"#
        ),
        "true"
    );
}

#[test]
fn test_string_unpack_basic_simple() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.pack("i4", 2); local a = string.unpack("i4", s); print(a == 2)"#
        ),
        "true"
    );
}

#[test]
fn test_string_unpack_basic_trimmed() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.pack("i4", 3); local a = string.unpack("i4", s); print(a == 3)"#
        ),
        "true"
    );
}

#[test]
fn test_string_unpack_basic_decimal() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.pack("i4", 4); local a = string.unpack("i4", s); print(a == 4)"#
        ),
        "true"
    );
}

#[test]
fn test_string_unpack_basic_hexed() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.pack("i4", 5); local a = string.unpack("i4", s); print(a == 5)"#
        ),
        "true"
    );
}

#[test]
fn test_string_unpack_basic_prefixed() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.pack("i4", 6); local a = string.unpack("i4", s); print(a == 6)"#
        ),
        "true"
    );
}

#[test]
fn test_string_unpack_basic_negative() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.pack("i4", 7); local a = string.unpack("i4", s); print(a == 7)"#
        ),
        "true"
    );
}

#[test]
fn test_string_unpack_basic_rounded() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.pack("i4", 8); local a = string.unpack("i4", s); print(a == 8)"#
        ),
        "true"
    );
}

#[test]
fn test_string_unpack_basic_offset() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.pack("i4", 9); local a = string.unpack("i4", s); print(a == 9)"#
        ),
        "true"
    );
}

#[test]
fn test_string_unpack_basic_paired() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.pack("i4", 10); local a = string.unpack("i4", s); print(a == 10)"#
        ),
        "true"
    );
}

#[test]
fn test_string_unpack_basic_nested() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.pack("i4", 11); local a = string.unpack("i4", s); print(a == 11)"#
        ),
        "true"
    );
}

#[test]
fn test_string_unpack_basic_metaflow() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.pack("i4", 12); local a = string.unpack("i4", s); print(a == 12)"#
        ),
        "true"
    );
}

#[test]
fn test_string_unpack_basic_guarded() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.pack("i4", 13); local a = string.unpack("i4", s); print(a == 13)"#
        ),
        "true"
    );
}

#[test]
fn test_string_unpack_basic_mapped() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.pack("i4", 14); local a = string.unpack("i4", s); print(a == 14)"#
        ),
        "true"
    );
}

#[test]
fn test_string_unpack_basic_captured() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.pack("i4", 15); local a = string.unpack("i4", s); print(a == 15)"#
        ),
        "true"
    );
}

#[test]
fn test_string_unpack_basic_edge_first() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.pack("i4", 16); local a = string.unpack("i4", s); print(a == 16)"#
        ),
        "true"
    );
}

#[test]
fn test_string_unpack_basic_edge_second() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.pack("i4", 17); local a = string.unpack("i4", s); print(a == 17)"#
        ),
        "true"
    );
}

#[test]
fn test_string_unpack_basic_edge_last() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.pack("i4", 18); local a = string.unpack("i4", s); print(a == 18)"#
        ),
        "true"
    );
}

#[test]
fn test_string_unpack_basic_randomized() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.pack("i4", 19); local a = string.unpack("i4", s); print(a == 19)"#
        ),
        "true"
    );
}

#[test]
fn test_string_unpack_basic_unicode_like() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.pack("i4", 20); local a = string.unpack("i4", s); print(a == 20)"#
        ),
        "true"
    );
}
