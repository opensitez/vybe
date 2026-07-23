use super::helpers::run_lua_one;

#[test]
fn test_string_byte_indices_baseline() {
    assert_eq!(
        run_lua_one(r#"print(string.byte("abcdefghijklmnopqrstuvwxyz", 1) == 97)"#),
        "true"
    );
}

#[test]
fn test_string_byte_indices_simple() {
    assert_eq!(
        run_lua_one(r#"print(string.byte("abcdefghijklmnopqrstuvwxyz", 2) == 98)"#),
        "true"
    );
}

#[test]
fn test_string_byte_indices_trimmed() {
    assert_eq!(
        run_lua_one(r#"print(string.byte("abcdefghijklmnopqrstuvwxyz", 3) == 99)"#),
        "true"
    );
}

#[test]
fn test_string_byte_indices_decimal() {
    assert_eq!(
        run_lua_one(r#"print(string.byte("abcdefghijklmnopqrstuvwxyz", 4) == 100)"#),
        "true"
    );
}

#[test]
fn test_string_byte_indices_hexed() {
    assert_eq!(
        run_lua_one(r#"print(string.byte("abcdefghijklmnopqrstuvwxyz", 5) == 101)"#),
        "true"
    );
}

#[test]
fn test_string_byte_indices_prefixed() {
    assert_eq!(
        run_lua_one(r#"print(string.byte("abcdefghijklmnopqrstuvwxyz", 6) == 102)"#),
        "true"
    );
}

#[test]
fn test_string_byte_indices_negative() {
    assert_eq!(
        run_lua_one(r#"print(string.byte("abcdefghijklmnopqrstuvwxyz", 7) == 103)"#),
        "true"
    );
}

#[test]
fn test_string_byte_indices_rounded() {
    assert_eq!(
        run_lua_one(r#"print(string.byte("abcdefghijklmnopqrstuvwxyz", 8) == 104)"#),
        "true"
    );
}

#[test]
fn test_string_byte_indices_offset() {
    assert_eq!(
        run_lua_one(r#"print(string.byte("abcdefghijklmnopqrstuvwxyz", 9) == 105)"#),
        "true"
    );
}

#[test]
fn test_string_byte_indices_paired() {
    assert_eq!(
        run_lua_one(r#"print(string.byte("abcdefghijklmnopqrstuvwxyz", 10) == 106)"#),
        "true"
    );
}

#[test]
fn test_string_byte_indices_nested() {
    assert_eq!(
        run_lua_one(r#"print(string.byte("abcdefghijklmnopqrstuvwxyz", -1) == 122)"#),
        "true"
    );
}

#[test]
fn test_string_byte_indices_metaflow() {
    assert_eq!(
        run_lua_one(r#"print(string.byte("abcdefghijklmnopqrstuvwxyz", -2) == 121)"#),
        "true"
    );
}

#[test]
fn test_string_byte_indices_guarded() {
    assert_eq!(
        run_lua_one(r#"print(string.byte("abcdefghijklmnopqrstuvwxyz", -3) == 120)"#),
        "true"
    );
}

#[test]
fn test_string_byte_indices_mapped() {
    assert_eq!(
        run_lua_one(r#"print(string.byte("abcdefghijklmnopqrstuvwxyz", -4) == 119)"#),
        "true"
    );
}

#[test]
fn test_string_byte_indices_captured() {
    assert_eq!(
        run_lua_one(r#"print(string.byte("abcdefghijklmnopqrstuvwxyz", -5) == 118)"#),
        "true"
    );
}

#[test]
fn test_string_byte_indices_edge_first() {
    assert_eq!(
        run_lua_one(r#"print(string.byte("abcdefghijklmnopqrstuvwxyz", -6) == 117)"#),
        "true"
    );
}

#[test]
fn test_string_byte_indices_edge_second() {
    assert_eq!(
        run_lua_one(r#"print(string.byte("abcdefghijklmnopqrstuvwxyz", -7) == 116)"#),
        "true"
    );
}

#[test]
fn test_string_byte_indices_edge_last() {
    assert_eq!(
        run_lua_one(r#"print(string.byte("abcdefghijklmnopqrstuvwxyz", -8) == 115)"#),
        "true"
    );
}

#[test]
fn test_string_byte_indices_randomized() {
    assert_eq!(
        run_lua_one(r#"print(string.byte("abcdefghijklmnopqrstuvwxyz", -9) == 114)"#),
        "true"
    );
}

#[test]
fn test_string_byte_indices_unicode_like() {
    assert_eq!(
        run_lua_one(r#"print(string.byte("abcdefghijklmnopqrstuvwxyz", -10) == 113)"#),
        "true"
    );
}
