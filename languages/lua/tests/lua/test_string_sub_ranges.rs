use super::helpers::run_lua_one;

#[test]
fn test_string_sub_ranges_baseline() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", 1, 3) == string.sub("abcdefghijklmnopqrstuvwxyz", 1, 3))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_ranges_simple() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", 1, 4) == string.sub("abcdefghijklmnopqrstuvwxyz", 1, 4))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_ranges_trimmed() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", 1, 5) == string.sub("abcdefghijklmnopqrstuvwxyz", 1, 5))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_ranges_decimal() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", 1, 6) == string.sub("abcdefghijklmnopqrstuvwxyz", 1, 6))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_ranges_hexed() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", 1, 7) == string.sub("abcdefghijklmnopqrstuvwxyz", 1, 7))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_ranges_prefixed() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", 1, 8) == string.sub("abcdefghijklmnopqrstuvwxyz", 1, 8))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_ranges_negative() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", 1, 9) == string.sub("abcdefghijklmnopqrstuvwxyz", 1, 9))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_ranges_rounded() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", 1, 10) == string.sub("abcdefghijklmnopqrstuvwxyz", 1, 10))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_ranges_offset() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", 1, 11) == string.sub("abcdefghijklmnopqrstuvwxyz", 1, 11))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_ranges_paired() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", 1, 12) == string.sub("abcdefghijklmnopqrstuvwxyz", 1, 12))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_ranges_nested() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", 1, 13) == string.sub("abcdefghijklmnopqrstuvwxyz", 1, 13))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_ranges_metaflow() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", 1, 14) == string.sub("abcdefghijklmnopqrstuvwxyz", 1, 14))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_ranges_guarded() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", 1, 15) == string.sub("abcdefghijklmnopqrstuvwxyz", 1, 15))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_ranges_mapped() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", 1, 16) == string.sub("abcdefghijklmnopqrstuvwxyz", 1, 16))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_ranges_captured() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", 1, 17) == string.sub("abcdefghijklmnopqrstuvwxyz", 1, 17))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_ranges_edge_first() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", 1, 18) == string.sub("abcdefghijklmnopqrstuvwxyz", 1, 18))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_ranges_edge_second() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", 1, 19) == string.sub("abcdefghijklmnopqrstuvwxyz", 1, 19))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_ranges_edge_last() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", 1, 20) == string.sub("abcdefghijklmnopqrstuvwxyz", 1, 20))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_ranges_randomized() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", 1, 21) == string.sub("abcdefghijklmnopqrstuvwxyz", 1, 21))"#
        ),
        "true"
    );
}

#[test]
fn test_string_sub_ranges_unicode_like() {
    assert_eq!(
        run_lua_one(
            r#"print(string.sub("abcdefghijklmnopqrstuvwxyz", 1, 22) == string.sub("abcdefghijklmnopqrstuvwxyz", 1, 22))"#
        ),
        "true"
    );
}
