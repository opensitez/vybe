use super::helpers::run_lua_one;

#[test]
fn test_string_match_capture_baseline() {
    assert_eq!(
        run_lua_one(
            r#"local tag, num = string.match("id1:value2", "(id%d+):(value%d+)")
print(tag == "id1" and num == "value2")"#
        ),
        "true"
    );
}

#[test]
fn test_string_match_capture_simple() {
    assert_eq!(
        run_lua_one(
            r#"local tag, num = string.match("id2:value4", "(id%d+):(value%d+)")
print(tag == "id2" and num == "value4")"#
        ),
        "true"
    );
}

#[test]
fn test_string_match_capture_trimmed() {
    assert_eq!(
        run_lua_one(
            r#"local tag, num = string.match("id3:value6", "(id%d+):(value%d+)")
print(tag == "id3" and num == "value6")"#
        ),
        "true"
    );
}

#[test]
fn test_string_match_capture_decimal() {
    assert_eq!(
        run_lua_one(
            r#"local tag, num = string.match("id4:value8", "(id%d+):(value%d+)")
print(tag == "id4" and num == "value8")"#
        ),
        "true"
    );
}

#[test]
fn test_string_match_capture_hexed() {
    assert_eq!(
        run_lua_one(
            r#"local tag, num = string.match("id5:value10", "(id%d+):(value%d+)")
print(tag == "id5" and num == "value10")"#
        ),
        "true"
    );
}

#[test]
fn test_string_match_capture_prefixed() {
    assert_eq!(
        run_lua_one(
            r#"local tag, num = string.match("id6:value12", "(id%d+):(value%d+)")
print(tag == "id6" and num == "value12")"#
        ),
        "true"
    );
}

#[test]
fn test_string_match_capture_negative() {
    assert_eq!(
        run_lua_one(
            r#"local tag, num = string.match("id7:value14", "(id%d+):(value%d+)")
print(tag == "id7" and num == "value14")"#
        ),
        "true"
    );
}

#[test]
fn test_string_match_capture_rounded() {
    assert_eq!(
        run_lua_one(
            r#"local tag, num = string.match("id8:value16", "(id%d+):(value%d+)")
print(tag == "id8" and num == "value16")"#
        ),
        "true"
    );
}

#[test]
fn test_string_match_capture_offset() {
    assert_eq!(
        run_lua_one(
            r#"local tag, num = string.match("id9:value18", "(id%d+):(value%d+)")
print(tag == "id9" and num == "value18")"#
        ),
        "true"
    );
}

#[test]
fn test_string_match_capture_paired() {
    assert_eq!(
        run_lua_one(
            r#"local tag, num = string.match("id10:value20", "(id%d+):(value%d+)")
print(tag == "id10" and num == "value20")"#
        ),
        "true"
    );
}

#[test]
fn test_string_match_capture_nested() {
    assert_eq!(
        run_lua_one(
            r#"local tag, num = string.match("id11:value22", "(id%d+):(value%d+)")
print(tag == "id11" and num == "value22")"#
        ),
        "true"
    );
}

#[test]
fn test_string_match_capture_metaflow() {
    assert_eq!(
        run_lua_one(
            r#"local tag, num = string.match("id12:value24", "(id%d+):(value%d+)")
print(tag == "id12" and num == "value24")"#
        ),
        "true"
    );
}

#[test]
fn test_string_match_capture_guarded() {
    assert_eq!(
        run_lua_one(
            r#"local tag, num = string.match("id13:value26", "(id%d+):(value%d+)")
print(tag == "id13" and num == "value26")"#
        ),
        "true"
    );
}

#[test]
fn test_string_match_capture_mapped() {
    assert_eq!(
        run_lua_one(
            r#"local tag, num = string.match("id14:value28", "(id%d+):(value%d+)")
print(tag == "id14" and num == "value28")"#
        ),
        "true"
    );
}

#[test]
fn test_string_match_capture_captured() {
    assert_eq!(
        run_lua_one(
            r#"local tag, num = string.match("id15:value30", "(id%d+):(value%d+)")
print(tag == "id15" and num == "value30")"#
        ),
        "true"
    );
}

#[test]
fn test_string_match_capture_edge_first() {
    assert_eq!(
        run_lua_one(
            r#"local tag, num = string.match("id16:value32", "(id%d+):(value%d+)")
print(tag == "id16" and num == "value32")"#
        ),
        "true"
    );
}

#[test]
fn test_string_match_capture_edge_second() {
    assert_eq!(
        run_lua_one(
            r#"local tag, num = string.match("id17:value34", "(id%d+):(value%d+)")
print(tag == "id17" and num == "value34")"#
        ),
        "true"
    );
}

#[test]
fn test_string_match_capture_edge_last() {
    assert_eq!(
        run_lua_one(
            r#"local tag, num = string.match("id18:value36", "(id%d+):(value%d+)")
print(tag == "id18" and num == "value36")"#
        ),
        "true"
    );
}

#[test]
fn test_string_match_capture_randomized() {
    assert_eq!(
        run_lua_one(
            r#"local tag, num = string.match("id19:value38", "(id%d+):(value%d+)")
print(tag == "id19" and num == "value38")"#
        ),
        "true"
    );
}

#[test]
fn test_string_match_capture_unicode_like() {
    assert_eq!(
        run_lua_one(
            r#"local tag, num = string.match("id20:value40", "(id%d+):(value%d+)")
print(tag == "id20" and num == "value40")"#
        ),
        "true"
    );
}
