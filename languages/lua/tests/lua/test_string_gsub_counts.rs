use super::helpers::run_lua_one;

#[test]
fn test_string_gsub_counts_baseline() {
    assert_eq!(
        run_lua_one(
            r#"local _, c = string.gsub(string.rep("a", 1), "a", "a")
print(c == 1)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_counts_simple() {
    assert_eq!(
        run_lua_one(
            r#"local _, c = string.gsub(string.rep("a", 2), "a", "a")
print(c == 2)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_counts_trimmed() {
    assert_eq!(
        run_lua_one(
            r#"local _, c = string.gsub(string.rep("a", 3), "a", "a")
print(c == 3)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_counts_decimal() {
    assert_eq!(
        run_lua_one(
            r#"local _, c = string.gsub(string.rep("a", 4), "a", "a")
print(c == 4)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_counts_hexed() {
    assert_eq!(
        run_lua_one(
            r#"local _, c = string.gsub(string.rep("a", 5), "a", "a")
print(c == 5)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_counts_prefixed() {
    assert_eq!(
        run_lua_one(
            r#"local _, c = string.gsub(string.rep("a", 6), "a", "a")
print(c == 6)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_counts_negative() {
    assert_eq!(
        run_lua_one(
            r#"local _, c = string.gsub(string.rep("a", 7), "a", "a")
print(c == 7)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_counts_rounded() {
    assert_eq!(
        run_lua_one(
            r#"local _, c = string.gsub(string.rep("a", 8), "a", "a")
print(c == 8)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_counts_offset() {
    assert_eq!(
        run_lua_one(
            r#"local _, c = string.gsub(string.rep("a", 9), "a", "a")
print(c == 9)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_counts_paired() {
    assert_eq!(
        run_lua_one(
            r#"local _, c = string.gsub(string.rep("a", 10), "a", "a")
print(c == 10)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_counts_nested() {
    assert_eq!(
        run_lua_one(
            r#"local _, c = string.gsub(string.rep("a", 11), "a", "a")
print(c == 11)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_counts_metaflow() {
    assert_eq!(
        run_lua_one(
            r#"local _, c = string.gsub(string.rep("a", 12), "a", "a")
print(c == 12)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_counts_guarded() {
    assert_eq!(
        run_lua_one(
            r#"local _, c = string.gsub(string.rep("a", 13), "a", "a")
print(c == 13)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_counts_mapped() {
    assert_eq!(
        run_lua_one(
            r#"local _, c = string.gsub(string.rep("a", 14), "a", "a")
print(c == 14)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_counts_captured() {
    assert_eq!(
        run_lua_one(
            r#"local _, c = string.gsub(string.rep("a", 15), "a", "a")
print(c == 15)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_counts_edge_first() {
    assert_eq!(
        run_lua_one(
            r#"local _, c = string.gsub(string.rep("a", 16), "a", "a")
print(c == 16)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_counts_edge_second() {
    assert_eq!(
        run_lua_one(
            r#"local _, c = string.gsub(string.rep("a", 17), "a", "a")
print(c == 17)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_counts_edge_last() {
    assert_eq!(
        run_lua_one(
            r#"local _, c = string.gsub(string.rep("a", 18), "a", "a")
print(c == 18)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_counts_randomized() {
    assert_eq!(
        run_lua_one(
            r#"local _, c = string.gsub(string.rep("a", 19), "a", "a")
print(c == 19)"#
        ),
        "true"
    );
}

#[test]
fn test_string_gsub_counts_unicode_like() {
    assert_eq!(
        run_lua_one(
            r#"local _, c = string.gsub(string.rep("a", 20), "a", "a")
print(c == 20)"#
        ),
        "true"
    );
}
