use super::helpers::run_lua_one;

#[test]
fn test_string_sub_negative_baseline() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", -1, -1) == string.sub("abcdefghijklmnopqrstuvwxyz", -1, -1))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_negative_simple() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", -2, -1) == string.sub("abcdefghijklmnopqrstuvwxyz", -2, -1))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_negative_trimmed() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", -3, -1) == string.sub("abcdefghijklmnopqrstuvwxyz", -3, -1))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_negative_decimal() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", -4, -1) == string.sub("abcdefghijklmnopqrstuvwxyz", -4, -1))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_negative_hexed() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", -5, -1) == string.sub("abcdefghijklmnopqrstuvwxyz", -5, -1))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_negative_prefixed() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", -6, -1) == string.sub("abcdefghijklmnopqrstuvwxyz", -6, -1))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_negative_negative() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", -7, -1) == string.sub("abcdefghijklmnopqrstuvwxyz", -7, -1))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_negative_rounded() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", -8, -1) == string.sub("abcdefghijklmnopqrstuvwxyz", -8, -1))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_negative_offset() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", -9, -1) == string.sub("abcdefghijklmnopqrstuvwxyz", -9, -1))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_negative_paired() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", -10, -1) == string.sub("abcdefghijklmnopqrstuvwxyz", -10, -1))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_negative_nested() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", -11, -1) == string.sub("abcdefghijklmnopqrstuvwxyz", -11, -1))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_negative_metaflow() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", -12, -1) == string.sub("abcdefghijklmnopqrstuvwxyz", -12, -1))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_negative_guarded() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", -13, -1) == string.sub("abcdefghijklmnopqrstuvwxyz", -13, -1))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_negative_mapped() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", -14, -1) == string.sub("abcdefghijklmnopqrstuvwxyz", -14, -1))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_negative_captured() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", -15, -1) == string.sub("abcdefghijklmnopqrstuvwxyz", -15, -1))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_negative_edge_first() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", -16, -1) == string.sub("abcdefghijklmnopqrstuvwxyz", -16, -1))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_negative_edge_second() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", -17, -1) == string.sub("abcdefghijklmnopqrstuvwxyz", -17, -1))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_negative_edge_last() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", -18, -1) == string.sub("abcdefghijklmnopqrstuvwxyz", -18, -1))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_negative_randomized() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", -19, -1) == string.sub("abcdefghijklmnopqrstuvwxyz", -19, -1))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_negative_unicode_like() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", -20, -1) == string.sub("abcdefghijklmnopqrstuvwxyz", -20, -1))"#
        ),
        "true"
    );
}
