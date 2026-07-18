use super::helpers::run_lua_one;

#[test]
fn test_math_random_range_baseline() {
    assert_eq!(run_lua_one(r#"math.randomseed(5)
local x = math.random(1, 3)
print(x >= 1 and x <= 3)"#), "true");
}


#[test]
fn test_math_random_range_simple() {
    assert_eq!(run_lua_one(r#"math.randomseed(6)
local x = math.random(1, 4)
print(x >= 1 and x <= 4)"#), "true");
}


#[test]
fn test_math_random_range_trimmed() {
    assert_eq!(run_lua_one(r#"math.randomseed(7)
local x = math.random(1, 5)
print(x >= 1 and x <= 5)"#), "true");
}


#[test]
fn test_math_random_range_decimal() {
    assert_eq!(run_lua_one(r#"math.randomseed(8)
local x = math.random(1, 6)
print(x >= 1 and x <= 6)"#), "true");
}


#[test]
fn test_math_random_range_hexed() {
    assert_eq!(run_lua_one(r#"math.randomseed(9)
local x = math.random(1, 7)
print(x >= 1 and x <= 7)"#), "true");
}


#[test]
fn test_math_random_range_prefixed() {
    assert_eq!(run_lua_one(r#"math.randomseed(10)
local x = math.random(1, 8)
print(x >= 1 and x <= 8)"#), "true");
}


#[test]
fn test_math_random_range_negative() {
    assert_eq!(run_lua_one(r#"math.randomseed(11)
local x = math.random(1, 9)
print(x >= 1 and x <= 9)"#), "true");
}


#[test]
fn test_math_random_range_rounded() {
    assert_eq!(run_lua_one(r#"math.randomseed(12)
local x = math.random(1, 10)
print(x >= 1 and x <= 10)"#), "true");
}


#[test]
fn test_math_random_range_offset() {
    assert_eq!(run_lua_one(r#"math.randomseed(13)
local x = math.random(1, 11)
print(x >= 1 and x <= 11)"#), "true");
}


#[test]
fn test_math_random_range_paired() {
    assert_eq!(run_lua_one(r#"math.randomseed(14)
local x = math.random(1, 12)
print(x >= 1 and x <= 12)"#), "true");
}


#[test]
fn test_math_random_range_nested() {
    assert_eq!(run_lua_one(r#"math.randomseed(15)
local x = math.random(1, 13)
print(x >= 1 and x <= 13)"#), "true");
}


#[test]
fn test_math_random_range_metaflow() {
    assert_eq!(run_lua_one(r#"math.randomseed(16)
local x = math.random(1, 14)
print(x >= 1 and x <= 14)"#), "true");
}


#[test]
fn test_math_random_range_guarded() {
    assert_eq!(run_lua_one(r#"math.randomseed(17)
local x = math.random(1, 15)
print(x >= 1 and x <= 15)"#), "true");
}


#[test]
fn test_math_random_range_mapped() {
    assert_eq!(run_lua_one(r#"math.randomseed(18)
local x = math.random(1, 16)
print(x >= 1 and x <= 16)"#), "true");
}


#[test]
fn test_math_random_range_captured() {
    assert_eq!(run_lua_one(r#"math.randomseed(19)
local x = math.random(1, 17)
print(x >= 1 and x <= 17)"#), "true");
}


#[test]
fn test_math_random_range_edge_first() {
    assert_eq!(run_lua_one(r#"math.randomseed(20)
local x = math.random(1, 18)
print(x >= 1 and x <= 18)"#), "true");
}


#[test]
fn test_math_random_range_edge_second() {
    assert_eq!(run_lua_one(r#"math.randomseed(21)
local x = math.random(1, 19)
print(x >= 1 and x <= 19)"#), "true");
}


#[test]
fn test_math_random_range_edge_last() {
    assert_eq!(run_lua_one(r#"math.randomseed(22)
local x = math.random(1, 20)
print(x >= 1 and x <= 20)"#), "true");
}


#[test]
fn test_math_random_range_randomized() {
    assert_eq!(run_lua_one(r#"math.randomseed(23)
local x = math.random(1, 21)
print(x >= 1 and x <= 21)"#), "true");
}


#[test]
fn test_math_random_range_unicode_like() {
    assert_eq!(run_lua_one(r#"math.randomseed(24)
local x = math.random(1, 22)
print(x >= 1 and x <= 22)"#), "true");
}
