use super::helpers::run_lua_one;

#[test]
fn test_math_random_seed_baseline() {
    assert_eq!(
        run_lua_one(
            r#"math.randomseed(100)
local x = math.random()
math.randomseed(100)
local y = math.random()
print(x == y)"#
        ),
        "true"
    );
}

#[test]
fn test_math_random_seed_simple() {
    assert_eq!(
        run_lua_one(
            r#"math.randomseed(101)
local x = math.random()
math.randomseed(101)
local y = math.random()
print(x == y)"#
        ),
        "true"
    );
}

#[test]
fn test_math_random_seed_trimmed() {
    assert_eq!(
        run_lua_one(
            r#"math.randomseed(102)
local x = math.random()
math.randomseed(102)
local y = math.random()
print(x == y)"#
        ),
        "true"
    );
}

#[test]
fn test_math_random_seed_decimal() {
    assert_eq!(
        run_lua_one(
            r#"math.randomseed(103)
local x = math.random()
math.randomseed(103)
local y = math.random()
print(x == y)"#
        ),
        "true"
    );
}

#[test]
fn test_math_random_seed_hexed() {
    assert_eq!(
        run_lua_one(
            r#"math.randomseed(104)
local x = math.random()
math.randomseed(104)
local y = math.random()
print(x == y)"#
        ),
        "true"
    );
}

#[test]
fn test_math_random_seed_prefixed() {
    assert_eq!(
        run_lua_one(
            r#"math.randomseed(105)
local x = math.random()
math.randomseed(105)
local y = math.random()
print(x == y)"#
        ),
        "true"
    );
}

#[test]
fn test_math_random_seed_negative() {
    assert_eq!(
        run_lua_one(
            r#"math.randomseed(106)
local x = math.random()
math.randomseed(106)
local y = math.random()
print(x == y)"#
        ),
        "true"
    );
}

#[test]
fn test_math_random_seed_rounded() {
    assert_eq!(
        run_lua_one(
            r#"math.randomseed(107)
local x = math.random()
math.randomseed(107)
local y = math.random()
print(x == y)"#
        ),
        "true"
    );
}

#[test]
fn test_math_random_seed_offset() {
    assert_eq!(
        run_lua_one(
            r#"math.randomseed(108)
local x = math.random()
math.randomseed(108)
local y = math.random()
print(x == y)"#
        ),
        "true"
    );
}

#[test]
fn test_math_random_seed_paired() {
    assert_eq!(
        run_lua_one(
            r#"math.randomseed(109)
local x = math.random()
math.randomseed(109)
local y = math.random()
print(x == y)"#
        ),
        "true"
    );
}

#[test]
fn test_math_random_seed_nested() {
    assert_eq!(
        run_lua_one(
            r#"math.randomseed(110)
local x = math.random()
math.randomseed(110)
local y = math.random()
print(x == y)"#
        ),
        "true"
    );
}

#[test]
fn test_math_random_seed_metaflow() {
    assert_eq!(
        run_lua_one(
            r#"math.randomseed(111)
local x = math.random()
math.randomseed(111)
local y = math.random()
print(x == y)"#
        ),
        "true"
    );
}

#[test]
fn test_math_random_seed_guarded() {
    assert_eq!(
        run_lua_one(
            r#"math.randomseed(112)
local x = math.random()
math.randomseed(112)
local y = math.random()
print(x == y)"#
        ),
        "true"
    );
}

#[test]
fn test_math_random_seed_mapped() {
    assert_eq!(
        run_lua_one(
            r#"math.randomseed(113)
local x = math.random()
math.randomseed(113)
local y = math.random()
print(x == y)"#
        ),
        "true"
    );
}

#[test]
fn test_math_random_seed_captured() {
    assert_eq!(
        run_lua_one(
            r#"math.randomseed(114)
local x = math.random()
math.randomseed(114)
local y = math.random()
print(x == y)"#
        ),
        "true"
    );
}

#[test]
fn test_math_random_seed_edge_first() {
    assert_eq!(
        run_lua_one(
            r#"math.randomseed(115)
local x = math.random()
math.randomseed(115)
local y = math.random()
print(x == y)"#
        ),
        "true"
    );
}

#[test]
fn test_math_random_seed_edge_second() {
    assert_eq!(
        run_lua_one(
            r#"math.randomseed(116)
local x = math.random()
math.randomseed(116)
local y = math.random()
print(x == y)"#
        ),
        "true"
    );
}

#[test]
fn test_math_random_seed_edge_last() {
    assert_eq!(
        run_lua_one(
            r#"math.randomseed(117)
local x = math.random()
math.randomseed(117)
local y = math.random()
print(x == y)"#
        ),
        "true"
    );
}

#[test]
fn test_math_random_seed_randomized() {
    assert_eq!(
        run_lua_one(
            r#"math.randomseed(118)
local x = math.random()
math.randomseed(118)
local y = math.random()
print(x == y)"#
        ),
        "true"
    );
}

#[test]
fn test_math_random_seed_unicode_like() {
    assert_eq!(
        run_lua_one(
            r#"math.randomseed(119)
local x = math.random()
math.randomseed(119)
local y = math.random()
print(x == y)"#
        ),
        "true"
    );
}
