use super::helpers::run_lua_one;

#[test]
fn test_assert_varargs_baseline() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(1, 2, 3); print(v1 + v2 == 1 + 2)"#), "true");
}


#[test]
fn test_assert_varargs_simple() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(2, 3, 4); print(v1 + v2 == 2 + 3)"#), "true");
}


#[test]
fn test_assert_varargs_trimmed() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(3, 4, 5); print(v1 + v2 == 3 + 4)"#), "true");
}


#[test]
fn test_assert_varargs_decimal() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(4, 5, 6); print(v1 + v2 == 4 + 5)"#), "true");
}


#[test]
fn test_assert_varargs_hexed() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(5, 6, 7); print(v1 + v2 == 5 + 6)"#), "true");
}


#[test]
fn test_assert_varargs_prefixed() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(6, 7, 8); print(v1 + v2 == 6 + 7)"#), "true");
}


#[test]
fn test_assert_varargs_negative() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(7, 8, 9); print(v1 + v2 == 7 + 8)"#), "true");
}


#[test]
fn test_assert_varargs_rounded() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(8, 9, 10); print(v1 + v2 == 8 + 9)"#), "true");
}


#[test]
fn test_assert_varargs_offset() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(9, 10, 11); print(v1 + v2 == 9 + 10)"#), "true");
}


#[test]
fn test_assert_varargs_paired() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(10, 11, 12); print(v1 + v2 == 10 + 11)"#), "true");
}


#[test]
fn test_assert_varargs_nested() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(11, 12, 13); print(v1 + v2 == 11 + 12)"#), "true");
}


#[test]
fn test_assert_varargs_metaflow() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(12, 13, 14); print(v1 + v2 == 12 + 13)"#), "true");
}


#[test]
fn test_assert_varargs_guarded() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(13, 14, 15); print(v1 + v2 == 13 + 14)"#), "true");
}


#[test]
fn test_assert_varargs_mapped() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(14, 15, 16); print(v1 + v2 == 14 + 15)"#), "true");
}


#[test]
fn test_assert_varargs_captured() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(15, 16, 17); print(v1 + v2 == 15 + 16)"#), "true");
}


#[test]
fn test_assert_varargs_edge_first() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(16, 17, 18); print(v1 + v2 == 16 + 17)"#), "true");
}


#[test]
fn test_assert_varargs_edge_second() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(17, 18, 19); print(v1 + v2 == 17 + 18)"#), "true");
}


#[test]
fn test_assert_varargs_edge_last() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(18, 19, 20); print(v1 + v2 == 18 + 19)"#), "true");
}


#[test]
fn test_assert_varargs_randomized() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(19, 20, 21); print(v1 + v2 == 19 + 20)"#), "true");
}


#[test]
fn test_assert_varargs_unicode_like() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(20, 21, 22); print(v1 + v2 == 20 + 21)"#), "true");
}
