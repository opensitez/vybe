use super::helpers::run_lua_one;

#[test]
fn test_error_level_0_message() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() error("lvl0", 0) end)
print(string.find(err, "lvl0") ~= nil)"#), "true");
}

#[test]
fn test_error_level_1_message() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() error("lvl1", 1) end)
print(string.find(err, "lvl1") ~= nil)"#), "true");
}

#[test]
fn test_error_level_2_message() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() error("lvl2", 2) end)
print(string.find(err, "lvl2") ~= nil)"#), "true");
}

#[test]
fn test_error_level_3_message() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() error("lvl3", 3) end)
print(string.find(err, "lvl3") ~= nil)"#), "true");
}

#[test]
fn test_error_level_4_message() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() error("lvl4", 4) end)
print(string.find(err, "lvl4") ~= nil)"#), "true");
}

#[test]
fn test_error_level_negative() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() error("neg", -1) end)
print(string.find(err, "neg") ~= nil)"#), "true");
}

#[test]
fn test_error_level_large_number() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() error("big", 99) end)
print(string.find(err, "big") ~= nil)"#), "true");
}

#[test]
fn test_error_level_with_assert() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() assert(false, "x") end)
print(type(err) == "string")"#), "true");
}

#[test]
fn test_error_level_float() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() error("f", 1.5) end)
print(type(err) == "string")"#), "true");
}

#[test]
fn test_error_level_string_level() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() error("s", "2") end)
print(type(err) == "string")"#), "true");
}

#[test]
fn test_error_level_true_level() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() error("t", true) end)
print(string.find(err, "t") ~= nil)"#), "true");
}

#[test]
fn test_error_level_nil_level() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() error("nil", nil) end)
print(string.find(err, "nil") ~= nil)"#), "true");
}

#[test]
fn test_error_level_in_function() {
    assert_eq!(run_lua_one(r#"local function f()
  error("inner", 2)
end
local ok, err = pcall(f)
print(string.find(err, "inner") ~= nil)"#), "true");
}

#[test]
fn test_error_level_nested_function() {
    assert_eq!(run_lua_one(r#"local function outer()
  local function inner() error("nest", 3) end
  inner()
end
local ok, err = pcall(outer)
print(string.find(err, "nest") ~= nil)"#), "true");
}

#[test]
fn test_error_level_with_recursion() {
    assert_eq!(run_lua_one(r#"local function walk(n)
  if n > 0 then return walk(n - 1) end
  error("deep", 4)
end
local ok, err = pcall(function() walk(1) end)
print(string.find(err, "deep") ~= nil)"#), "true");
}

#[test]
fn test_error_level_conditional() {
    assert_eq!(run_lua_one(r#"local function f(x) if x then error("cond", 2) else return true end end
local ok, err = pcall(function() f(true) end)
print(string.find(err, "cond") ~= nil)"#), "true");
}

#[test]
fn test_error_level_pcall_wrapper() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() pcall(function() error("wrap", 2) end) end)
print(ok == true and err == nil)"#), "true");
}

#[test]
fn test_error_level_table_payload() {
    assert_eq!(run_lua_one(r#"local t = {reason = "tbl"}
local ok, err = pcall(function() error(t, 2) end)
print(type(err) == "table")"#), "true");
}

#[test]
fn test_error_level_bool_payload() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() error(false, 3) end)
print(err == false)"#), "true");
}

#[test]
fn test_error_level_function_payload() {
    assert_eq!(run_lua_one(r#"local function k() end
local ok, err = pcall(function() error(k, 2) end)
print(type(err) == "function")"#), "true");
}

#[test]
fn test_error_level_function_in_pcall() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() error("x", 2) end)
print(ok == false)"#), "true");
}
