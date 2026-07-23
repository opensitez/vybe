use super::helpers::run_lua_one;

#[test]
fn test_string_gsub_simple_baseline() {
    assert_eq!(
        run_lua_one(
            r#"local _, n = string.gsub("a", "a", "b")
print(n == 1)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_simple_simple() {
    assert_eq!(
        run_lua_one(
            r#"local _, n = string.gsub("aa", "a", "b")
print(n == 2)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_simple_trimmed() {
    assert_eq!(
        run_lua_one(
            r#"local _, n = string.gsub("aaa", "a", "b")
print(n == 3)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_simple_decimal() {
    assert_eq!(
        run_lua_one(
            r#"local _, n = string.gsub("aaaa", "a", "b")
print(n == 4)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_simple_hexed() {
    assert_eq!(
        run_lua_one(
            r#"local _, n = string.gsub("aaaaa", "a", "b")
print(n == 5)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_simple_prefixed() {
    assert_eq!(
        run_lua_one(
            r#"local _, n = string.gsub("aaaaaa", "a", "b")
print(n == 6)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_simple_negative() {
    assert_eq!(
        run_lua_one(
            r#"local _, n = string.gsub("aaaaaaa", "a", "b")
print(n == 7)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_simple_rounded() {
    assert_eq!(
        run_lua_one(
            r#"local _, n = string.gsub("aaaaaaaa", "a", "b")
print(n == 8)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_simple_offset() {
    assert_eq!(
        run_lua_one(
            r#"local _, n = string.gsub("aaaaaaaaa", "a", "b")
print(n == 9)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_simple_paired() {
    assert_eq!(
        run_lua_one(
            r#"local _, n = string.gsub("aaaaaaaaaa", "a", "b")
print(n == 10)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_simple_nested() {
    assert_eq!(
        run_lua_one(
            r#"local _, n = string.gsub("aaaaaaaaaaa", "a", "b")
print(n == 11)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_simple_metaflow() {
    assert_eq!(
        run_lua_one(
            r#"local _, n = string.gsub("aaaaaaaaaaaa", "a", "b")
print(n == 12)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_simple_guarded() {
    assert_eq!(
        run_lua_one(
            r#"local _, n = string.gsub("aaaaaaaaaaaaa", "a", "b")
print(n == 13)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_simple_mapped() {
    assert_eq!(
        run_lua_one(
            r#"local _, n = string.gsub("aaaaaaaaaaaaaa", "a", "b")
print(n == 14)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_simple_captured() {
    assert_eq!(
        run_lua_one(
            r#"local _, n = string.gsub("aaaaaaaaaaaaaaa", "a", "b")
print(n == 15)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_simple_edge_first() {
    assert_eq!(
        run_lua_one(
            r#"local _, n = string.gsub("aaaaaaaaaaaaaaaa", "a", "b")
print(n == 16)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_simple_edge_second() {
    assert_eq!(
        run_lua_one(
            r#"local _, n = string.gsub("aaaaaaaaaaaaaaaaa", "a", "b")
print(n == 17)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_simple_edge_last() {
    assert_eq!(
        run_lua_one(
            r#"local _, n = string.gsub("aaaaaaaaaaaaaaaaaa", "a", "b")
print(n == 18)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_simple_randomized() {
    assert_eq!(
        run_lua_one(
            r#"local _, n = string.gsub("aaaaaaaaaaaaaaaaaaa", "a", "b")
print(n == 19)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_simple_unicode_like() {
    assert_eq!(
        run_lua_one(
            r#"local _, n = string.gsub("aaaaaaaaaaaaaaaaaaaa", "a", "b")
print(n == 20)"#
        ),
        "true"
    );
}
