use super::helpers::run_lua_one;

#[test]
fn test_print_nested_baseline() {
    assert_eq!(run_lua_one(r#"print((function() return (function(x) return (function(y) return x + y end)(1) end)(2) end)())"#), "true");
}


#[test]
fn test_print_nested_simple() {
    assert_eq!(run_lua_one(r#"print((function() return (function(x) return (function(y) return x + y end)(2) end)(3) end)())"#), "true");
}


#[test]
fn test_print_nested_trimmed() {
    assert_eq!(run_lua_one(r#"print((function() return (function(x) return (function(y) return x + y end)(3) end)(4) end)())"#), "true");
}


#[test]
fn test_print_nested_decimal() {
    assert_eq!(run_lua_one(r#"print((function() return (function(x) return (function(y) return x + y end)(4) end)(5) end)())"#), "true");
}


#[test]
fn test_print_nested_hexed() {
    assert_eq!(run_lua_one(r#"print((function() return (function(x) return (function(y) return x + y end)(5) end)(6) end)())"#), "true");
}


#[test]
fn test_print_nested_prefixed() {
    assert_eq!(run_lua_one(r#"print((function() return (function(x) return (function(y) return x + y end)(6) end)(7) end)())"#), "true");
}


#[test]
fn test_print_nested_negative() {
    assert_eq!(run_lua_one(r#"print((function() return (function(x) return (function(y) return x + y end)(7) end)(8) end)())"#), "true");
}


#[test]
fn test_print_nested_rounded() {
    assert_eq!(run_lua_one(r#"print((function() return (function(x) return (function(y) return x + y end)(8) end)(9) end)())"#), "true");
}


#[test]
fn test_print_nested_offset() {
    assert_eq!(run_lua_one(r#"print((function() return (function(x) return (function(y) return x + y end)(9) end)(10) end)())"#), "true");
}


#[test]
fn test_print_nested_paired() {
    assert_eq!(run_lua_one(r#"print((function() return (function(x) return (function(y) return x + y end)(10) end)(11) end)())"#), "true");
}


#[test]
fn test_print_nested_nested() {
    assert_eq!(run_lua_one(r#"print((function() return (function(x) return (function(y) return x + y end)(11) end)(12) end)())"#), "true");
}


#[test]
fn test_print_nested_metaflow() {
    assert_eq!(run_lua_one(r#"print((function() return (function(x) return (function(y) return x + y end)(12) end)(13) end)())"#), "true");
}


#[test]
fn test_print_nested_guarded() {
    assert_eq!(run_lua_one(r#"print((function() return (function(x) return (function(y) return x + y end)(13) end)(14) end)())"#), "true");
}


#[test]
fn test_print_nested_mapped() {
    assert_eq!(run_lua_one(r#"print((function() return (function(x) return (function(y) return x + y end)(14) end)(15) end)())"#), "true");
}


#[test]
fn test_print_nested_captured() {
    assert_eq!(run_lua_one(r#"print((function() return (function(x) return (function(y) return x + y end)(15) end)(16) end)())"#), "true");
}


#[test]
fn test_print_nested_edge_first() {
    assert_eq!(run_lua_one(r#"print((function() return (function(x) return (function(y) return x + y end)(16) end)(17) end)())"#), "true");
}


#[test]
fn test_print_nested_edge_second() {
    assert_eq!(run_lua_one(r#"print((function() return (function(x) return (function(y) return x + y end)(17) end)(18) end)())"#), "true");
}


#[test]
fn test_print_nested_edge_last() {
    assert_eq!(run_lua_one(r#"print((function() return (function(x) return (function(y) return x + y end)(18) end)(19) end)())"#), "true");
}


#[test]
fn test_print_nested_randomized() {
    assert_eq!(run_lua_one(r#"print((function() return (function(x) return (function(y) return x + y end)(19) end)(20) end)())"#), "true");
}


#[test]
fn test_print_nested_unicode_like() {
    assert_eq!(run_lua_one(r#"print((function() return (function(x) return (function(y) return x + y end)(20) end)(21) end)())"#), "true");
}
