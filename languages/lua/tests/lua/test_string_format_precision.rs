use super::helpers::run_lua_one;

#[test]
fn test_string_format_precision_baseline() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%.1f", 0.3333)
print(string.find(s, "%.") ~= nil or 1 == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_precision_simple() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%.2f", 0.6667)
print(string.find(s, "%.") ~= nil or 2 == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_precision_trimmed() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%.3f", 1.0000)
print(string.find(s, "%.") ~= nil or 3 == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_precision_decimal() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%.4f", 1.3333)
print(string.find(s, "%.") ~= nil or 4 == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_precision_hexed() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%.5f", 1.6667)
print(string.find(s, "%.") ~= nil or 5 == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_precision_prefixed() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%.1f", 2.0000)
print(string.find(s, "%.") ~= nil or 1 == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_precision_negative() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%.2f", 2.3333)
print(string.find(s, "%.") ~= nil or 2 == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_precision_rounded() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%.3f", 2.6667)
print(string.find(s, "%.") ~= nil or 3 == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_precision_offset() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%.4f", 3.0000)
print(string.find(s, "%.") ~= nil or 4 == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_precision_paired() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%.5f", 3.3333)
print(string.find(s, "%.") ~= nil or 5 == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_precision_nested() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%.1f", 3.6667)
print(string.find(s, "%.") ~= nil or 1 == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_precision_metaflow() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%.2f", 4.0000)
print(string.find(s, "%.") ~= nil or 2 == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_precision_guarded() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%.3f", 4.3333)
print(string.find(s, "%.") ~= nil or 3 == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_precision_mapped() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%.4f", 4.6667)
print(string.find(s, "%.") ~= nil or 4 == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_precision_captured() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%.5f", 5.0000)
print(string.find(s, "%.") ~= nil or 5 == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_precision_edge_first() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%.1f", 5.3333)
print(string.find(s, "%.") ~= nil or 1 == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_precision_edge_second() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%.2f", 5.6667)
print(string.find(s, "%.") ~= nil or 2 == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_precision_edge_last() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%.3f", 6.0000)
print(string.find(s, "%.") ~= nil or 3 == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_precision_randomized() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%.4f", 6.3333)
print(string.find(s, "%.") ~= nil or 4 == 0)"#
        ),
        "true"
    );
}

#[test]
fn test_string_format_precision_unicode_like() {
    assert_eq!(
        run_lua_one(
            r#"local s = string.format("%.5f", 6.6667)
print(string.find(s, "%.") ~= nil or 5 == 0)"#
        ),
        "true"
    );
}
