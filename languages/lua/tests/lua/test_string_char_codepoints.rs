use super::helpers::run_lua_one;

#[test]
fn test_string_char_codepoints_baseline() {
    assert_eq!(
        run_lua_one(r#"local c = string.char(65); print(string.byte(c) == 65)"#),
        "true"
    );
}

#[test]
fn test_string_char_codepoints_simple() {
    assert_eq!(
        run_lua_one(r#"local c = string.char(66); print(string.byte(c) == 66)"#),
        "true"
    );
}

#[test]
fn test_string_char_codepoints_trimmed() {
    assert_eq!(
        run_lua_one(r#"local c = string.char(67); print(string.byte(c) == 67)"#),
        "true"
    );
}

#[test]
fn test_string_char_codepoints_decimal() {
    assert_eq!(
        run_lua_one(r#"local c = string.char(68); print(string.byte(c) == 68)"#),
        "true"
    );
}

#[test]
fn test_string_char_codepoints_hexed() {
    assert_eq!(
        run_lua_one(r#"local c = string.char(69); print(string.byte(c) == 69)"#),
        "true"
    );
}

#[test]
fn test_string_char_codepoints_prefixed() {
    assert_eq!(
        run_lua_one(r#"local c = string.char(70); print(string.byte(c) == 70)"#),
        "true"
    );
}

#[test]
fn test_string_char_codepoints_negative() {
    assert_eq!(
        run_lua_one(r#"local c = string.char(71); print(string.byte(c) == 71)"#),
        "true"
    );
}

#[test]
fn test_string_char_codepoints_rounded() {
    assert_eq!(
        run_lua_one(r#"local c = string.char(72); print(string.byte(c) == 72)"#),
        "true"
    );
}

#[test]
fn test_string_char_codepoints_offset() {
    assert_eq!(
        run_lua_one(r#"local c = string.char(73); print(string.byte(c) == 73)"#),
        "true"
    );
}

#[test]
fn test_string_char_codepoints_paired() {
    assert_eq!(
        run_lua_one(r#"local c = string.char(74); print(string.byte(c) == 74)"#),
        "true"
    );
}

#[test]
fn test_string_char_codepoints_nested() {
    assert_eq!(
        run_lua_one(r#"local c = string.char(75); print(string.byte(c) == 75)"#),
        "true"
    );
}

#[test]
fn test_string_char_codepoints_metaflow() {
    assert_eq!(
        run_lua_one(r#"local c = string.char(76); print(string.byte(c) == 76)"#),
        "true"
    );
}

#[test]
fn test_string_char_codepoints_guarded() {
    assert_eq!(
        run_lua_one(r#"local c = string.char(77); print(string.byte(c) == 77)"#),
        "true"
    );
}

#[test]
fn test_string_char_codepoints_mapped() {
    assert_eq!(
        run_lua_one(r#"local c = string.char(78); print(string.byte(c) == 78)"#),
        "true"
    );
}

#[test]
fn test_string_char_codepoints_captured() {
    assert_eq!(
        run_lua_one(r#"local c = string.char(79); print(string.byte(c) == 79)"#),
        "true"
    );
}

#[test]
fn test_string_char_codepoints_edge_first() {
    assert_eq!(
        run_lua_one(r#"local c = string.char(80); print(string.byte(c) == 80)"#),
        "true"
    );
}

#[test]
fn test_string_char_codepoints_edge_second() {
    assert_eq!(
        run_lua_one(r#"local c = string.char(81); print(string.byte(c) == 81)"#),
        "true"
    );
}

#[test]
fn test_string_char_codepoints_edge_last() {
    assert_eq!(
        run_lua_one(r#"local c = string.char(82); print(string.byte(c) == 82)"#),
        "true"
    );
}

#[test]
fn test_string_char_codepoints_randomized() {
    assert_eq!(
        run_lua_one(r#"local c = string.char(83); print(string.byte(c) == 83)"#),
        "true"
    );
}

#[test]
fn test_string_char_codepoints_unicode_like() {
    assert_eq!(
        run_lua_one(r#"local c = string.char(84); print(string.byte(c) == 84)"#),
        "true"
    );
}
