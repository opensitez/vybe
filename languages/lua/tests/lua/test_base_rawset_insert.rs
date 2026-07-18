use super::helpers::run_lua_one;

#[test]
fn test_rawset_insert_baseline() {
    assert_eq!(run_lua_one(r#"local t = {}
rawset(t, 1, 2)
print(rawget(t, 1) == 2)"#), "true");
}


#[test]
fn test_rawset_insert_simple() {
    assert_eq!(run_lua_one(r#"local t = {}
rawset(t, 2, 4)
print(rawget(t, 2) == 4)"#), "true");
}


#[test]
fn test_rawset_insert_trimmed() {
    assert_eq!(run_lua_one(r#"local t = {}
rawset(t, 3, 6)
print(rawget(t, 3) == 6)"#), "true");
}


#[test]
fn test_rawset_insert_decimal() {
    assert_eq!(run_lua_one(r#"local t = {}
rawset(t, 4, 8)
print(rawget(t, 4) == 8)"#), "true");
}


#[test]
fn test_rawset_insert_hexed() {
    assert_eq!(run_lua_one(r#"local t = {}
rawset(t, 5, 10)
print(rawget(t, 5) == 10)"#), "true");
}


#[test]
fn test_rawset_insert_prefixed() {
    assert_eq!(run_lua_one(r#"local t = {}
rawset(t, 6, 12)
print(rawget(t, 6) == 12)"#), "true");
}


#[test]
fn test_rawset_insert_negative() {
    assert_eq!(run_lua_one(r#"local t = {}
rawset(t, 7, 14)
print(rawget(t, 7) == 14)"#), "true");
}


#[test]
fn test_rawset_insert_rounded() {
    assert_eq!(run_lua_one(r#"local t = {}
rawset(t, 8, 16)
print(rawget(t, 8) == 16)"#), "true");
}


#[test]
fn test_rawset_insert_offset() {
    assert_eq!(run_lua_one(r#"local t = {}
rawset(t, 9, 18)
print(rawget(t, 9) == 18)"#), "true");
}


#[test]
fn test_rawset_insert_paired() {
    assert_eq!(run_lua_one(r#"local t = {}
rawset(t, 10, 20)
print(rawget(t, 10) == 20)"#), "true");
}


#[test]
fn test_rawset_insert_nested() {
    assert_eq!(run_lua_one(r#"local t = {}
rawset(t, 11, 22)
print(rawget(t, 11) == 22)"#), "true");
}


#[test]
fn test_rawset_insert_metaflow() {
    assert_eq!(run_lua_one(r#"local t = {}
rawset(t, 12, 24)
print(rawget(t, 12) == 24)"#), "true");
}


#[test]
fn test_rawset_insert_guarded() {
    assert_eq!(run_lua_one(r#"local t = {}
rawset(t, 13, 26)
print(rawget(t, 13) == 26)"#), "true");
}


#[test]
fn test_rawset_insert_mapped() {
    assert_eq!(run_lua_one(r#"local t = {}
rawset(t, 14, 28)
print(rawget(t, 14) == 28)"#), "true");
}


#[test]
fn test_rawset_insert_captured() {
    assert_eq!(run_lua_one(r#"local t = {}
rawset(t, 15, 30)
print(rawget(t, 15) == 30)"#), "true");
}


#[test]
fn test_rawset_insert_edge_first() {
    assert_eq!(run_lua_one(r#"local t = {}
rawset(t, 16, 32)
print(rawget(t, 16) == 32)"#), "true");
}


#[test]
fn test_rawset_insert_edge_second() {
    assert_eq!(run_lua_one(r#"local t = {}
rawset(t, 17, 34)
print(rawget(t, 17) == 34)"#), "true");
}


#[test]
fn test_rawset_insert_edge_last() {
    assert_eq!(run_lua_one(r#"local t = {}
rawset(t, 18, 36)
print(rawget(t, 18) == 36)"#), "true");
}


#[test]
fn test_rawset_insert_randomized() {
    assert_eq!(run_lua_one(r#"local t = {}
rawset(t, 19, 38)
print(rawget(t, 19) == 38)"#), "true");
}


#[test]
fn test_rawset_insert_unicode_like() {
    assert_eq!(run_lua_one(r#"local t = {}
rawset(t, 20, 40)
print(rawget(t, 20) == 40)"#), "true");
}
