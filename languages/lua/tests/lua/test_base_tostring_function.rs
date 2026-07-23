use super::helpers::run_lua_one;

#[test]
fn test_tostring_function_baseline() {
    assert_eq!(
        run_lua_one(
            r#"local f = function() return 0 end; local s = tostring(f); print(type(s) == "string" and string.sub(s, 1, 9) == "function:")"#
        ),
        "true"
    );
}

#[test]
fn test_tostring_function_simple() {
    assert_eq!(
        run_lua_one(
            r#"local f = function() return 1 + 1 end; local s = tostring(f); print(type(s) == "string" and string.sub(s, 1, 9) == "function:")"#
        ),
        "true"
    );
}

#[test]
fn test_tostring_function_trimmed() {
    assert_eq!(
        run_lua_one(
            r#"local f = function() return 2 end; local s = tostring(f); print(type(s) == "string" and string.sub(s, 1, 9) == "function:")"#
        ),
        "true"
    );
}

#[test]
fn test_tostring_function_decimal() {
    assert_eq!(
        run_lua_one(
            r#"local f = function() return 3 + 1 end; local s = tostring(f); print(type(s) == "string" and string.sub(s, 1, 9) == "function:")"#
        ),
        "true"
    );
}

#[test]
fn test_tostring_function_hexed() {
    assert_eq!(
        run_lua_one(
            r#"local f = function() return 4 end; local s = tostring(f); print(type(s) == "string" and string.sub(s, 1, 9) == "function:")"#
        ),
        "true"
    );
}

#[test]
fn test_tostring_function_prefixed() {
    assert_eq!(
        run_lua_one(
            r#"local f = function() return 5 + 1 end; local s = tostring(f); print(type(s) == "string" and string.sub(s, 1, 9) == "function:")"#
        ),
        "true"
    );
}

#[test]
fn test_tostring_function_negative() {
    assert_eq!(
        run_lua_one(
            r#"local f = function() return 6 end; local s = tostring(f); print(type(s) == "string" and string.sub(s, 1, 9) == "function:")"#
        ),
        "true"
    );
}

#[test]
fn test_tostring_function_rounded() {
    assert_eq!(
        run_lua_one(
            r#"local f = function() return 7 + 1 end; local s = tostring(f); print(type(s) == "string" and string.sub(s, 1, 9) == "function:")"#
        ),
        "true"
    );
}

#[test]
fn test_tostring_function_offset() {
    assert_eq!(
        run_lua_one(
            r#"local f = function() return 8 end; local s = tostring(f); print(type(s) == "string" and string.sub(s, 1, 9) == "function:")"#
        ),
        "true"
    );
}

#[test]
fn test_tostring_function_paired() {
    assert_eq!(
        run_lua_one(
            r#"local f = function() return 9 + 1 end; local s = tostring(f); print(type(s) == "string" and string.sub(s, 1, 9) == "function:")"#
        ),
        "true"
    );
}

#[test]
fn test_tostring_function_nested() {
    assert_eq!(
        run_lua_one(
            r#"local f = function() return 10 end; local s = tostring(f); print(type(s) == "string" and string.sub(s, 1, 9) == "function:")"#
        ),
        "true"
    );
}

#[test]
fn test_tostring_function_metaflow() {
    assert_eq!(
        run_lua_one(
            r#"local f = function() return 11 + 1 end; local s = tostring(f); print(type(s) == "string" and string.sub(s, 1, 9) == "function:")"#
        ),
        "true"
    );
}

#[test]
fn test_tostring_function_guarded() {
    assert_eq!(
        run_lua_one(
            r#"local f = function() return 12 end; local s = tostring(f); print(type(s) == "string" and string.sub(s, 1, 9) == "function:")"#
        ),
        "true"
    );
}

#[test]
fn test_tostring_function_mapped() {
    assert_eq!(
        run_lua_one(
            r#"local f = function() return 13 + 1 end; local s = tostring(f); print(type(s) == "string" and string.sub(s, 1, 9) == "function:")"#
        ),
        "true"
    );
}

#[test]
fn test_tostring_function_captured() {
    assert_eq!(
        run_lua_one(
            r#"local f = function() return 14 end; local s = tostring(f); print(type(s) == "string" and string.sub(s, 1, 9) == "function:")"#
        ),
        "true"
    );
}

#[test]
fn test_tostring_function_edge_first() {
    assert_eq!(
        run_lua_one(
            r#"local f = function() return 15 + 1 end; local s = tostring(f); print(type(s) == "string" and string.sub(s, 1, 9) == "function:")"#
        ),
        "true"
    );
}

#[test]
fn test_tostring_function_edge_second() {
    assert_eq!(
        run_lua_one(
            r#"local f = function() return 16 end; local s = tostring(f); print(type(s) == "string" and string.sub(s, 1, 9) == "function:")"#
        ),
        "true"
    );
}

#[test]
fn test_tostring_function_edge_last() {
    assert_eq!(
        run_lua_one(
            r#"local f = function() return 17 + 1 end; local s = tostring(f); print(type(s) == "string" and string.sub(s, 1, 9) == "function:")"#
        ),
        "true"
    );
}

#[test]
fn test_tostring_function_randomized() {
    assert_eq!(
        run_lua_one(
            r#"local f = function() return 18 end; local s = tostring(f); print(type(s) == "string" and string.sub(s, 1, 9) == "function:")"#
        ),
        "true"
    );
}

#[test]
fn test_tostring_function_unicode_like() {
    assert_eq!(
        run_lua_one(
            r#"local f = function() return 19 + 1 end; local s = tostring(f); print(type(s) == "string" and string.sub(s, 1, 9) == "function:")"#
        ),
        "true"
    );
}
