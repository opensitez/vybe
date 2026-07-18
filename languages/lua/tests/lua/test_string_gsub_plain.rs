use super::helpers::run_lua_one;

#[test]
fn test_string_gsub_plain_baseline() {
    assert_eq!(run_lua_one(r#"local s = string.rep("x+y", 1)
local _, replaced = string.gsub(s, "x+y", "z", 1)
print(replaced == 1)"#), "true");
}


#[test]
fn test_string_gsub_plain_simple() {
    assert_eq!(run_lua_one(r#"local s = string.rep("x+y", 2)
local _, replaced = string.gsub(s, "x+y", "z", 1)
print(replaced == 1)"#), "true");
}


#[test]
fn test_string_gsub_plain_trimmed() {
    assert_eq!(run_lua_one(r#"local s = string.rep("x+y", 3)
local _, replaced = string.gsub(s, "x+y", "z", 1)
print(replaced == 1)"#), "true");
}


#[test]
fn test_string_gsub_plain_decimal() {
    assert_eq!(run_lua_one(r#"local s = string.rep("x+y", 4)
local _, replaced = string.gsub(s, "x+y", "z", 1)
print(replaced == 1)"#), "true");
}


#[test]
fn test_string_gsub_plain_hexed() {
    assert_eq!(run_lua_one(r#"local s = string.rep("x+y", 5)
local _, replaced = string.gsub(s, "x+y", "z", 1)
print(replaced == 1)"#), "true");
}


#[test]
fn test_string_gsub_plain_prefixed() {
    assert_eq!(run_lua_one(r#"local s = string.rep("x+y", 6)
local _, replaced = string.gsub(s, "x+y", "z", 1)
print(replaced == 1)"#), "true");
}


#[test]
fn test_string_gsub_plain_negative() {
    assert_eq!(run_lua_one(r#"local s = string.rep("x+y", 7)
local _, replaced = string.gsub(s, "x+y", "z", 1)
print(replaced == 1)"#), "true");
}


#[test]
fn test_string_gsub_plain_rounded() {
    assert_eq!(run_lua_one(r#"local s = string.rep("x+y", 8)
local _, replaced = string.gsub(s, "x+y", "z", 1)
print(replaced == 1)"#), "true");
}


#[test]
fn test_string_gsub_plain_offset() {
    assert_eq!(run_lua_one(r#"local s = string.rep("x+y", 9)
local _, replaced = string.gsub(s, "x+y", "z", 1)
print(replaced == 1)"#), "true");
}


#[test]
fn test_string_gsub_plain_paired() {
    assert_eq!(run_lua_one(r#"local s = string.rep("x+y", 10)
local _, replaced = string.gsub(s, "x+y", "z", 1)
print(replaced == 1)"#), "true");
}


#[test]
fn test_string_gsub_plain_nested() {
    assert_eq!(run_lua_one(r#"local s = string.rep("x+y", 11)
local _, replaced = string.gsub(s, "x+y", "z", 1)
print(replaced == 1)"#), "true");
}


#[test]
fn test_string_gsub_plain_metaflow() {
    assert_eq!(run_lua_one(r#"local s = string.rep("x+y", 12)
local _, replaced = string.gsub(s, "x+y", "z", 1)
print(replaced == 1)"#), "true");
}


#[test]
fn test_string_gsub_plain_guarded() {
    assert_eq!(run_lua_one(r#"local s = string.rep("x+y", 13)
local _, replaced = string.gsub(s, "x+y", "z", 1)
print(replaced == 1)"#), "true");
}


#[test]
fn test_string_gsub_plain_mapped() {
    assert_eq!(run_lua_one(r#"local s = string.rep("x+y", 14)
local _, replaced = string.gsub(s, "x+y", "z", 1)
print(replaced == 1)"#), "true");
}


#[test]
fn test_string_gsub_plain_captured() {
    assert_eq!(run_lua_one(r#"local s = string.rep("x+y", 15)
local _, replaced = string.gsub(s, "x+y", "z", 1)
print(replaced == 1)"#), "true");
}


#[test]
fn test_string_gsub_plain_edge_first() {
    assert_eq!(run_lua_one(r#"local s = string.rep("x+y", 16)
local _, replaced = string.gsub(s, "x+y", "z", 1)
print(replaced == 1)"#), "true");
}


#[test]
fn test_string_gsub_plain_edge_second() {
    assert_eq!(run_lua_one(r#"local s = string.rep("x+y", 17)
local _, replaced = string.gsub(s, "x+y", "z", 1)
print(replaced == 1)"#), "true");
}


#[test]
fn test_string_gsub_plain_edge_last() {
    assert_eq!(run_lua_one(r#"local s = string.rep("x+y", 18)
local _, replaced = string.gsub(s, "x+y", "z", 1)
print(replaced == 1)"#), "true");
}


#[test]
fn test_string_gsub_plain_randomized() {
    assert_eq!(run_lua_one(r#"local s = string.rep("x+y", 19)
local _, replaced = string.gsub(s, "x+y", "z", 1)
print(replaced == 1)"#), "true");
}


#[test]
fn test_string_gsub_plain_unicode_like() {
    assert_eq!(run_lua_one(r#"local s = string.rep("x+y", 20)
local _, replaced = string.gsub(s, "x+y", "z", 1)
print(replaced == 1)"#), "true");
}
