use super::helpers::run_lua_one;

#[test]
fn test_error_only_message() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() error("message") end)
print(err == "message")"#), "true");
}

#[test]
fn test_error_empty_message() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() error("") end)
print(err == "")"#), "true");
}

#[test]
fn test_error_number_message() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() error(15) end)
print(err == 15)"#), "true");
}

#[test]
fn test_error_boolean_message() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() error(false) end)
print(err == false)"#), "true");
}

#[test]
fn test_error_true_message() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() error(true) end)
print(err == true)"#), "true");
}

#[test]
fn test_error_table_message() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() error({label = "x"}) end)
print(type(err) == "table")"#), "true");
}

#[test]
fn test_error_function_message() {
    assert_eq!(run_lua_one(r#"local marker = function() return 1 end
local ok, err = pcall(function() error(marker) end)
print(type(err) == "function")"#), "true");
}

#[test]
fn test_error_thread_message() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() end)
local ok, err = pcall(function() error(t) end)
print(type(err) == "thread")"#), "true");
}

#[test]
fn test_error_table_field_accessible() {
    assert_eq!(run_lua_one(r#"local payload = {code = 500}
local ok, err = pcall(function() error(payload) end)
print(type(err) == "table" and err.code == 500)"#), "true");
}

#[test]
fn test_error_concat_with_message() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() error("a" .. "b") end)
print(err == "ab")"#), "true");
}

#[test]
fn test_error_nested_function() {
    assert_eq!(run_lua_one(r#"local function inner() error("inner") end
local function outer() inner() end
local ok, err = pcall(outer)
print(err == "inner")"#), "true");
}

#[test]
fn test_error_nested_return() {
    assert_eq!(run_lua_one(r#"local function f() error("ret") end
local ok, err = pcall(f)
print(err == "ret")"#), "true");
}

#[test]
fn test_error_with_conditional_true() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() if true then error("cond") end end)
print(err == "cond")"#), "true");
}

#[test]
fn test_error_with_conditional_false() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() if false then error("cond") else return true end end)
print(ok == true)"#), "true");
}

#[test]
fn test_error_from_assert() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() assert(false, "asserted") end)
print(type(err) == "string")"#), "true");
}

#[test]
fn test_error_from_pcall() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() pcall(function() error("x") end) end)
print(ok == true)"#), "true");
}

#[test]
fn test_error_recovery_after_error() {
    assert_eq!(run_lua_one(r#"local function run()
  local ok, err = pcall(function() error("fail") end)
  if ok then return "ok" end
  return err
end
print(run())"#), "fail");
}

#[test]
fn test_error_after_success() {
    assert_eq!(run_lua_one(r#"local function check()
  local ok, err = pcall(function() return 1 end)
  if ok then return 7 end
  return err
end
print(check())"#), "7");
}

#[test]
fn test_error_long_text() {
    assert_eq!(run_lua_one(r#"local text = string.rep("e", 16)
local ok, err = pcall(function() error(text) end)
print(string.len(err) == 16)"#), "true");
}

#[test]
fn test_error_with_newline() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() error("line1\\nline2") end)
print(string.find(err, "line1") ~= nil)"#), "true");
}

#[test]
fn test_error_in_inner_function() {
    assert_eq!(run_lua_one(r#"local fn = function() error("innerfn") end
local ok, err = pcall(fn)
print(err == "innerfn")"#), "true");
}

#[test]
fn test_error_to_boolean_false() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() error(false) end)
print(err == false)"#), "true");
}

#[test]
fn test_error_to_true() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() error(true) end)
print(err == true)"#), "true");
}
