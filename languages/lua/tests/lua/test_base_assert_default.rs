use super::helpers::run_lua_one;

#[test]
fn test_assert_default_baseline() {
    assert_eq!(run_lua_one(r#"print(select(1, pcall(function() assert(true) end)))"#), "true");
}


#[test]
fn test_assert_default_simple() {
    assert_eq!(run_lua_one(r#"print(select(1, pcall(function() assert(1 == 1) end)))"#), "true");
}


#[test]
fn test_assert_default_trimmed() {
    assert_eq!(run_lua_one(r#"print(select(1, pcall(function() assert(0 < 1) end)))"#), "true");
}


#[test]
fn test_assert_default_decimal() {
    assert_eq!(run_lua_one(r#"print(select(1, pcall(function() assert("x" ~= nil) end)))"#), "true");
}


#[test]
fn test_assert_default_hexed() {
    assert_eq!(run_lua_one(r#"print(select(1, pcall(function() assert(math.abs(-3) == 3) end)))"#), "true");
}


#[test]
fn test_assert_default_prefixed() {
    assert_eq!(run_lua_one(r#"print(select(1, pcall(function() assert(type({}) == "table") end)))"#), "true");
}


#[test]
fn test_assert_default_negative() {
    assert_eq!(run_lua_one(r#"print(select(1, pcall(function() assert(type(true) == "boolean") end)))"#), "true");
}


#[test]
fn test_assert_default_rounded() {
    assert_eq!(run_lua_one(r#"print(select(1, pcall(function() assert(tostring(2) == "2") end)))"#), "true");
}


#[test]
fn test_assert_default_offset() {
    assert_eq!(run_lua_one(r#"print(select(1, pcall(function() assert(string.len("abc") == 3) end)))"#), "true");
}


#[test]
fn test_assert_default_paired() {
    assert_eq!(run_lua_one(r#"print(select(1, pcall(function() assert(math.sin(0) == 0) end)))"#), "true");
}


#[test]
fn test_assert_default_nested() {
    assert_eq!(run_lua_one(r#"print(select(1, pcall(function() assert(1 + 1 == 2) end)))"#), "true");
}


#[test]
fn test_assert_default_metaflow() {
    assert_eq!(run_lua_one(r#"print(select(1, pcall(function() assert(#"x" == 1) end)))"#), "true");
}


#[test]
fn test_assert_default_guarded() {
    assert_eq!(run_lua_one(r#"print(select(1, pcall(function() assert(false or true) end)))"#), "true");
}


#[test]
fn test_assert_default_mapped() {
    assert_eq!(run_lua_one(r#"print(select(1, pcall(function() assert((1 + 2) == 3) end)))"#), "true");
}


#[test]
fn test_assert_default_captured() {
    assert_eq!(run_lua_one(r#"print(select(1, pcall(function() assert("hello" == "hello") end)))"#), "true");
}


#[test]
fn test_assert_default_edge_first() {
    assert_eq!(run_lua_one(r#"print(select(1, pcall(function() assert(type(nil) == "nil") end)))"#), "true");
}


#[test]
fn test_assert_default_edge_second() {
    assert_eq!(run_lua_one(r#"print(select(1, pcall(function() assert(9 >= 9) end)))"#), "true");
}


#[test]
fn test_assert_default_edge_last() {
    assert_eq!(run_lua_one(r#"print(select(1, pcall(function() assert(5 ~= 6) end)))"#), "true");
}


#[test]
fn test_assert_default_randomized() {
    assert_eq!(run_lua_one(r#"print(select(1, pcall(function() assert((not false)) end)))"#), "true");
}


#[test]
fn test_assert_default_unicode_like() {
    assert_eq!(run_lua_one(r#"print(select(1, pcall(function() assert(type({a = 1}) == "table") end)))"#), "true");
}
