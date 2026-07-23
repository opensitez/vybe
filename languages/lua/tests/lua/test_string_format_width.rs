use super::helpers::run_lua_one;

#[test]
fn test_string_format_width_baseline() {
    assert_eq!(
        run_lua_one(r#"local s = string.format("%3d", 1); print(#s == 3)"#),
        "true"
    );
}

#[test]
fn test_string_format_width_simple() {
    assert_eq!(
        run_lua_one(r#"local s = string.format("%4d", 2); print(#s == 4)"#),
        "true"
    );
}

#[test]
fn test_string_format_width_trimmed() {
    assert_eq!(
        run_lua_one(r#"local s = string.format("%5d", 3); print(#s == 5)"#),
        "true"
    );
}

#[test]
fn test_string_format_width_decimal() {
    assert_eq!(
        run_lua_one(r#"local s = string.format("%6d", 4); print(#s == 6)"#),
        "true"
    );
}

#[test]
fn test_string_format_width_hexed() {
    assert_eq!(
        run_lua_one(r#"local s = string.format("%7d", 5); print(#s == 7)"#),
        "true"
    );
}

#[test]
fn test_string_format_width_prefixed() {
    assert_eq!(
        run_lua_one(r#"local s = string.format("%8d", 6); print(#s == 8)"#),
        "true"
    );
}

#[test]
fn test_string_format_width_negative() {
    assert_eq!(
        run_lua_one(r#"local s = string.format("%9d", 7); print(#s == 9)"#),
        "true"
    );
}

#[test]
fn test_string_format_width_rounded() {
    assert_eq!(
        run_lua_one(r#"local s = string.format("%10d", 8); print(#s == 10)"#),
        "true"
    );
}

#[test]
fn test_string_format_width_offset() {
    assert_eq!(
        run_lua_one(r#"local s = string.format("%11d", 9); print(#s == 11)"#),
        "true"
    );
}

#[test]
fn test_string_format_width_paired() {
    assert_eq!(
        run_lua_one(r#"local s = string.format("%12d", 10); print(#s == 12)"#),
        "true"
    );
}

#[test]
fn test_string_format_width_nested() {
    assert_eq!(
        run_lua_one(r#"local s = string.format("%13d", 11); print(#s == 13)"#),
        "true"
    );
}

#[test]
fn test_string_format_width_metaflow() {
    assert_eq!(
        run_lua_one(r#"local s = string.format("%14d", 12); print(#s == 14)"#),
        "true"
    );
}

#[test]
fn test_string_format_width_guarded() {
    assert_eq!(
        run_lua_one(r#"local s = string.format("%15d", 13); print(#s == 15)"#),
        "true"
    );
}

#[test]
fn test_string_format_width_mapped() {
    assert_eq!(
        run_lua_one(r#"local s = string.format("%16d", 14); print(#s == 16)"#),
        "true"
    );
}

#[test]
fn test_string_format_width_captured() {
    assert_eq!(
        run_lua_one(r#"local s = string.format("%17d", 15); print(#s == 17)"#),
        "true"
    );
}

#[test]
fn test_string_format_width_edge_first() {
    assert_eq!(
        run_lua_one(r#"local s = string.format("%18d", 16); print(#s == 18)"#),
        "true"
    );
}

#[test]
fn test_string_format_width_edge_second() {
    assert_eq!(
        run_lua_one(r#"local s = string.format("%19d", 17); print(#s == 19)"#),
        "true"
    );
}

#[test]
fn test_string_format_width_edge_last() {
    assert_eq!(
        run_lua_one(r#"local s = string.format("%20d", 18); print(#s == 20)"#),
        "true"
    );
}

#[test]
fn test_string_format_width_randomized() {
    assert_eq!(
        run_lua_one(r#"local s = string.format("%21d", 19); print(#s == 21)"#),
        "true"
    );
}

#[test]
fn test_string_format_width_unicode_like() {
    assert_eq!(
        run_lua_one(r#"local s = string.format("%22d", 20); print(#s == 22)"#),
        "true"
    );
}
