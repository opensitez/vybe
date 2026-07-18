use super::helpers::run_lua_one;

#[test]
fn test_string_find_last_baseline() {
    assert_eq!(run_lua_one(r#"print(string.find("x1 x1 x1 x1 x1 x1 x1 x1 x1 x1 z1", "z1", 1, true) == 31)"#), "true");
}


#[test]
fn test_string_find_last_simple() {
    assert_eq!(run_lua_one(r#"print(string.find("x2 x2 x2 x2 x2 x2 x2 x2 x2 x2 z2", "z2", 1, true) == 31)"#), "true");
}


#[test]
fn test_string_find_last_trimmed() {
    assert_eq!(run_lua_one(r#"print(string.find("x3 x3 x3 x3 x3 x3 x3 x3 x3 x3 z3", "z3", 1, true) == 31)"#), "true");
}


#[test]
fn test_string_find_last_decimal() {
    assert_eq!(run_lua_one(r#"print(string.find("x4 x4 x4 x4 x4 x4 x4 x4 x4 x4 z4", "z4", 1, true) == 31)"#), "true");
}


#[test]
fn test_string_find_last_hexed() {
    assert_eq!(run_lua_one(r#"print(string.find("x5 x5 x5 x5 x5 x5 x5 x5 x5 x5 z5", "z5", 1, true) == 31)"#), "true");
}


#[test]
fn test_string_find_last_prefixed() {
    assert_eq!(run_lua_one(r#"print(string.find("x6 x6 x6 x6 x6 x6 x6 x6 x6 x6 z6", "z6", 1, true) == 31)"#), "true");
}


#[test]
fn test_string_find_last_negative() {
    assert_eq!(run_lua_one(r#"print(string.find("x7 x7 x7 x7 x7 x7 x7 x7 x7 x7 z7", "z7", 1, true) == 31)"#), "true");
}


#[test]
fn test_string_find_last_rounded() {
    assert_eq!(run_lua_one(r#"print(string.find("x8 x8 x8 x8 x8 x8 x8 x8 x8 x8 z8", "z8", 1, true) == 31)"#), "true");
}


#[test]
fn test_string_find_last_offset() {
    assert_eq!(run_lua_one(r#"print(string.find("x9 x9 x9 x9 x9 x9 x9 x9 x9 x9 z9", "z9", 1, true) == 31)"#), "true");
}


#[test]
fn test_string_find_last_paired() {
    assert_eq!(run_lua_one(r#"print(string.find("x10 x10 x10 x10 x10 x10 x10 x10 x10 x10 z10", "z10", 1, true) == 41)"#), "true");
}


#[test]
fn test_string_find_last_nested() {
    assert_eq!(run_lua_one(r#"print(string.find("x11 x11 x11 x11 x11 x11 x11 x11 x11 x11 z11", "z11", 1, true) == 41)"#), "true");
}


#[test]
fn test_string_find_last_metaflow() {
    assert_eq!(run_lua_one(r#"print(string.find("x12 x12 x12 x12 x12 x12 x12 x12 x12 x12 z12", "z12", 1, true) == 41)"#), "true");
}


#[test]
fn test_string_find_last_guarded() {
    assert_eq!(run_lua_one(r#"print(string.find("x13 x13 x13 x13 x13 x13 x13 x13 x13 x13 z13", "z13", 1, true) == 41)"#), "true");
}


#[test]
fn test_string_find_last_mapped() {
    assert_eq!(run_lua_one(r#"print(string.find("x14 x14 x14 x14 x14 x14 x14 x14 x14 x14 z14", "z14", 1, true) == 41)"#), "true");
}


#[test]
fn test_string_find_last_captured() {
    assert_eq!(run_lua_one(r#"print(string.find("x15 x15 x15 x15 x15 x15 x15 x15 x15 x15 z15", "z15", 1, true) == 41)"#), "true");
}


#[test]
fn test_string_find_last_edge_first() {
    assert_eq!(run_lua_one(r#"print(string.find("x16 x16 x16 x16 x16 x16 x16 x16 x16 x16 z16", "z16", 1, true) == 41)"#), "true");
}


#[test]
fn test_string_find_last_edge_second() {
    assert_eq!(run_lua_one(r#"print(string.find("x17 x17 x17 x17 x17 x17 x17 x17 x17 x17 z17", "z17", 1, true) == 41)"#), "true");
}


#[test]
fn test_string_find_last_edge_last() {
    assert_eq!(run_lua_one(r#"print(string.find("x18 x18 x18 x18 x18 x18 x18 x18 x18 x18 z18", "z18", 1, true) == 41)"#), "true");
}


#[test]
fn test_string_find_last_randomized() {
    assert_eq!(run_lua_one(r#"print(string.find("x19 x19 x19 x19 x19 x19 x19 x19 x19 x19 z19", "z19", 1, true) == 41)"#), "true");
}


#[test]
fn test_string_find_last_unicode_like() {
    assert_eq!(run_lua_one(r#"print(string.find("x20 x20 x20 x20 x20 x20 x20 x20 x20 x20 z20", "z20", 1, true) == 41)"#), "true");
}
