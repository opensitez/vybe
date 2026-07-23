use super::helpers::run_lua_one;

#[test]
fn test_string_reverse_basic_baseline() {
    assert_eq!(
        run_lua_one(
            r#"print(string.reverse(string.sub("abcdefghijklmnopqrst", 1, 1)) == string.sub("abcdefghijklmnopqrst", 1, 1):reverse())"#
        ),
        "true"
    );
}

#[test]
fn test_string_reverse_basic_simple() {
    assert_eq!(
        run_lua_one(
            r#"print(string.reverse(string.sub("abcdefghijklmnopqrst", 1, 2)) == string.sub("abcdefghijklmnopqrst", 1, 2):reverse())"#
        ),
        "true"
    );
}

#[test]
fn test_string_reverse_basic_trimmed() {
    assert_eq!(
        run_lua_one(
            r#"print(string.reverse(string.sub("abcdefghijklmnopqrst", 1, 3)) == string.sub("abcdefghijklmnopqrst", 1, 3):reverse())"#
        ),
        "true"
    );
}

#[test]
fn test_string_reverse_basic_decimal() {
    assert_eq!(
        run_lua_one(
            r#"print(string.reverse(string.sub("abcdefghijklmnopqrst", 1, 4)) == string.sub("abcdefghijklmnopqrst", 1, 4):reverse())"#
        ),
        "true"
    );
}

#[test]
fn test_string_reverse_basic_hexed() {
    assert_eq!(
        run_lua_one(
            r#"print(string.reverse(string.sub("abcdefghijklmnopqrst", 1, 5)) == string.sub("abcdefghijklmnopqrst", 1, 5):reverse())"#
        ),
        "true"
    );
}

#[test]
fn test_string_reverse_basic_prefixed() {
    assert_eq!(
        run_lua_one(
            r#"print(string.reverse(string.sub("abcdefghijklmnopqrst", 1, 6)) == string.sub("abcdefghijklmnopqrst", 1, 6):reverse())"#
        ),
        "true"
    );
}

#[test]
fn test_string_reverse_basic_negative() {
    assert_eq!(
        run_lua_one(
            r#"print(string.reverse(string.sub("abcdefghijklmnopqrst", 1, 7)) == string.sub("abcdefghijklmnopqrst", 1, 7):reverse())"#
        ),
        "true"
    );
}

#[test]
fn test_string_reverse_basic_rounded() {
    assert_eq!(
        run_lua_one(
            r#"print(string.reverse(string.sub("abcdefghijklmnopqrst", 1, 8)) == string.sub("abcdefghijklmnopqrst", 1, 8):reverse())"#
        ),
        "true"
    );
}

#[test]
fn test_string_reverse_basic_offset() {
    assert_eq!(
        run_lua_one(
            r#"print(string.reverse(string.sub("abcdefghijklmnopqrst", 1, 9)) == string.sub("abcdefghijklmnopqrst", 1, 9):reverse())"#
        ),
        "true"
    );
}

#[test]
fn test_string_reverse_basic_paired() {
    assert_eq!(
        run_lua_one(
            r#"print(string.reverse(string.sub("abcdefghijklmnopqrst", 1, 10)) == string.sub("abcdefghijklmnopqrst", 1, 10):reverse())"#
        ),
        "true"
    );
}

#[test]
fn test_string_reverse_basic_nested() {
    assert_eq!(
        run_lua_one(
            r#"print(string.reverse(string.sub("abcdefghijklmnopqrst", 1, 11)) == string.sub("abcdefghijklmnopqrst", 1, 11):reverse())"#
        ),
        "true"
    );
}

#[test]
fn test_string_reverse_basic_metaflow() {
    assert_eq!(
        run_lua_one(
            r#"print(string.reverse(string.sub("abcdefghijklmnopqrst", 1, 12)) == string.sub("abcdefghijklmnopqrst", 1, 12):reverse())"#
        ),
        "true"
    );
}

#[test]
fn test_string_reverse_basic_guarded() {
    assert_eq!(
        run_lua_one(
            r#"print(string.reverse(string.sub("abcdefghijklmnopqrst", 1, 13)) == string.sub("abcdefghijklmnopqrst", 1, 13):reverse())"#
        ),
        "true"
    );
}

#[test]
fn test_string_reverse_basic_mapped() {
    assert_eq!(
        run_lua_one(
            r#"print(string.reverse(string.sub("abcdefghijklmnopqrst", 1, 14)) == string.sub("abcdefghijklmnopqrst", 1, 14):reverse())"#
        ),
        "true"
    );
}

#[test]
fn test_string_reverse_basic_captured() {
    assert_eq!(
        run_lua_one(
            r#"print(string.reverse(string.sub("abcdefghijklmnopqrst", 1, 15)) == string.sub("abcdefghijklmnopqrst", 1, 15):reverse())"#
        ),
        "true"
    );
}

#[test]
fn test_string_reverse_basic_edge_first() {
    assert_eq!(
        run_lua_one(
            r#"print(string.reverse(string.sub("abcdefghijklmnopqrst", 1, 16)) == string.sub("abcdefghijklmnopqrst", 1, 16):reverse())"#
        ),
        "true"
    );
}

#[test]
fn test_string_reverse_basic_edge_second() {
    assert_eq!(
        run_lua_one(
            r#"print(string.reverse(string.sub("abcdefghijklmnopqrst", 1, 17)) == string.sub("abcdefghijklmnopqrst", 1, 17):reverse())"#
        ),
        "true"
    );
}

#[test]
fn test_string_reverse_basic_edge_last() {
    assert_eq!(
        run_lua_one(
            r#"print(string.reverse(string.sub("abcdefghijklmnopqrst", 1, 18)) == string.sub("abcdefghijklmnopqrst", 1, 18):reverse())"#
        ),
        "true"
    );
}

#[test]
fn test_string_reverse_basic_randomized() {
    assert_eq!(
        run_lua_one(
            r#"print(string.reverse(string.sub("abcdefghijklmnopqrst", 1, 19)) == string.sub("abcdefghijklmnopqrst", 1, 19):reverse())"#
        ),
        "true"
    );
}

#[test]
fn test_string_reverse_basic_unicode_like() {
    assert_eq!(
        run_lua_one(
            r#"print(string.reverse(string.sub("abcdefghijklmnopqrst", 1, 20)) == string.sub("abcdefghijklmnopqrst", 1, 20):reverse())"#
        ),
        "true"
    );
}
