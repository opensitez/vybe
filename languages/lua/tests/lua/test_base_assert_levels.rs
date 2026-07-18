use super::helpers::run_lua_one;

#[test]
fn test_assert_level_baseline() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(2, 10, undefined); print(v1 + v2 == 2 + 10)"#), "true");
}


#[test]
fn test_assert_level_simple() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(3, 11, undefined); print(v1 + v2 == 3 + 11)"#), "true");
}


#[test]
fn test_assert_level_trimmed() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(4, 12, undefined); print(v1 + v2 == 4 + 12)"#), "true");
}


#[test]
fn test_assert_level_decimal() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(5, 13, undefined); print(v1 + v2 == 5 + 13)"#), "true");
}


#[test]
fn test_assert_level_hexed() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(6, 14, undefined); print(v1 + v2 == 6 + 14)"#), "true");
}


#[test]
fn test_assert_level_prefixed() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(7, 15, undefined); print(v1 + v2 == 7 + 15)"#), "true");
}


#[test]
fn test_assert_level_negative() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(8, 16, undefined); print(v1 + v2 == 8 + 16)"#), "true");
}


#[test]
fn test_assert_level_rounded() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(9, 17, undefined); print(v1 + v2 == 9 + 17)"#), "true");
}


#[test]
fn test_assert_level_offset() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(10, 18, undefined); print(v1 + v2 == 10 + 18)"#), "true");
}


#[test]
fn test_assert_level_paired() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(11, 19, undefined); print(v1 + v2 == 11 + 19)"#), "true");
}


#[test]
fn test_assert_level_nested() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(12, 20, undefined); print(v1 + v2 == 12 + 20)"#), "true");
}


#[test]
fn test_assert_level_metaflow() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(13, 21, undefined); print(v1 + v2 == 13 + 21)"#), "true");
}


#[test]
fn test_assert_level_guarded() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(14, 22, undefined); print(v1 + v2 == 14 + 22)"#), "true");
}


#[test]
fn test_assert_level_mapped() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(15, 23, undefined); print(v1 + v2 == 15 + 23)"#), "true");
}


#[test]
fn test_assert_level_captured() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(16, 24, undefined); print(v1 + v2 == 16 + 24)"#), "true");
}


#[test]
fn test_assert_level_edge_first() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(17, 25, undefined); print(v1 + v2 == 17 + 25)"#), "true");
}


#[test]
fn test_assert_level_edge_second() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(18, 26, undefined); print(v1 + v2 == 18 + 26)"#), "true");
}


#[test]
fn test_assert_level_edge_last() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(19, 27, undefined); print(v1 + v2 == 19 + 27)"#), "true");
}


#[test]
fn test_assert_level_randomized() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(20, 28, undefined); print(v1 + v2 == 20 + 28)"#), "true");
}


#[test]
fn test_assert_level_unicode_like() {
    assert_eq!(run_lua_one(r#"local v1, v2, v3 = assert(21, 29, undefined); print(v1 + v2 == 21 + 29)"#), "true");
}
