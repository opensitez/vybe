use super::helpers::run_lua_one;

#[test]
fn test_table_len_boundaries_baseline() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}
rawset(t, 4, 5)
print(#t >= 3)"#), "true");
}


#[test]
fn test_table_len_boundaries_simple() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}
rawset(t, 5, 6)
print(#t >= 3)"#), "true");
}


#[test]
fn test_table_len_boundaries_trimmed() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}
rawset(t, 6, 7)
print(#t >= 3)"#), "true");
}


#[test]
fn test_table_len_boundaries_decimal() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}
rawset(t, 7, 8)
print(#t >= 3)"#), "true");
}


#[test]
fn test_table_len_boundaries_hexed() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}
rawset(t, 8, 9)
print(#t >= 3)"#), "true");
}


#[test]
fn test_table_len_boundaries_prefixed() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}
rawset(t, 9, 10)
print(#t >= 3)"#), "true");
}


#[test]
fn test_table_len_boundaries_negative() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}
rawset(t, 10, 11)
print(#t >= 3)"#), "true");
}


#[test]
fn test_table_len_boundaries_rounded() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}
rawset(t, 11, 12)
print(#t >= 3)"#), "true");
}


#[test]
fn test_table_len_boundaries_offset() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}
rawset(t, 12, 13)
print(#t >= 3)"#), "true");
}


#[test]
fn test_table_len_boundaries_paired() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}
rawset(t, 13, 14)
print(#t >= 3)"#), "true");
}


#[test]
fn test_table_len_boundaries_nested() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}
rawset(t, 14, 15)
print(#t >= 3)"#), "true");
}


#[test]
fn test_table_len_boundaries_metaflow() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}
rawset(t, 15, 16)
print(#t >= 3)"#), "true");
}


#[test]
fn test_table_len_boundaries_guarded() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}
rawset(t, 16, 17)
print(#t >= 3)"#), "true");
}


#[test]
fn test_table_len_boundaries_mapped() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}
rawset(t, 17, 18)
print(#t >= 3)"#), "true");
}


#[test]
fn test_table_len_boundaries_captured() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}
rawset(t, 18, 19)
print(#t >= 3)"#), "true");
}


#[test]
fn test_table_len_boundaries_edge_first() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}
rawset(t, 19, 20)
print(#t >= 3)"#), "true");
}


#[test]
fn test_table_len_boundaries_edge_second() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}
rawset(t, 20, 21)
print(#t >= 3)"#), "true");
}


#[test]
fn test_table_len_boundaries_edge_last() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}
rawset(t, 21, 22)
print(#t >= 3)"#), "true");
}


#[test]
fn test_table_len_boundaries_randomized() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}
rawset(t, 22, 23)
print(#t >= 3)"#), "true");
}


#[test]
fn test_table_len_boundaries_unicode_like() {
    assert_eq!(run_lua_one(r#"local t = {1,2,3}
rawset(t, 23, 24)
print(#t >= 3)"#), "true");
}
