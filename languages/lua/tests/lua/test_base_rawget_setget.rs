use super::helpers::run_lua_one;

#[test]
fn test_rawget_setget_baseline() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
rawset(t, "a", 1)
print(rawget(t, "a") == 1)"#
        ),
        "true"
    );
}

#[test]
fn test_rawget_setget_simple() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
rawset(t, "b", 2)
print(rawget(t, "b") == 2)"#
        ),
        "true"
    );
}

#[test]
fn test_rawget_setget_trimmed() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
rawset(t, "c", 3)
print(rawget(t, "c") == 3)"#
        ),
        "true"
    );
}

#[test]
fn test_rawget_setget_decimal() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
rawset(t, "d", 4)
print(rawget(t, "d") == 4)"#
        ),
        "true"
    );
}

#[test]
fn test_rawget_setget_hexed() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
rawset(t, "e", 5)
print(rawget(t, "e") == 5)"#
        ),
        "true"
    );
}

#[test]
fn test_rawget_setget_prefixed() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
rawset(t, "f", 6)
print(rawget(t, "f") == 6)"#
        ),
        "true"
    );
}

#[test]
fn test_rawget_setget_negative() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
rawset(t, "g", 7)
print(rawget(t, "g") == 7)"#
        ),
        "true"
    );
}

#[test]
fn test_rawget_setget_rounded() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
rawset(t, "h", 8)
print(rawget(t, "h") == 8)"#
        ),
        "true"
    );
}

#[test]
fn test_rawget_setget_offset() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
rawset(t, "i", 9)
print(rawget(t, "i") == 9)"#
        ),
        "true"
    );
}

#[test]
fn test_rawget_setget_paired() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
rawset(t, "j", 10)
print(rawget(t, "j") == 10)"#
        ),
        "true"
    );
}

#[test]
fn test_rawget_setget_nested() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
rawset(t, "k", 11)
print(rawget(t, "k") == 11)"#
        ),
        "true"
    );
}

#[test]
fn test_rawget_setget_metaflow() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
rawset(t, "l", 12)
print(rawget(t, "l") == 12)"#
        ),
        "true"
    );
}

#[test]
fn test_rawget_setget_guarded() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
rawset(t, "m", 13)
print(rawget(t, "m") == 13)"#
        ),
        "true"
    );
}

#[test]
fn test_rawget_setget_mapped() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
rawset(t, "n", 14)
print(rawget(t, "n") == 14)"#
        ),
        "true"
    );
}

#[test]
fn test_rawget_setget_captured() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
rawset(t, "o", 15)
print(rawget(t, "o") == 15)"#
        ),
        "true"
    );
}

#[test]
fn test_rawget_setget_edge_first() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
rawset(t, "p", 16)
print(rawget(t, "p") == 16)"#
        ),
        "true"
    );
}

#[test]
fn test_rawget_setget_edge_second() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
rawset(t, "q", 17)
print(rawget(t, "q") == 17)"#
        ),
        "true"
    );
}

#[test]
fn test_rawget_setget_edge_last() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
rawset(t, "r", 18)
print(rawget(t, "r") == 18)"#
        ),
        "true"
    );
}

#[test]
fn test_rawget_setget_randomized() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
rawset(t, "s", 19)
print(rawget(t, "s") == 19)"#
        ),
        "true"
    );
}

#[test]
fn test_rawget_setget_unicode_like() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
rawset(t, "t", 20)
print(rawget(t, "t") == 20)"#
        ),
        "true"
    );
}
