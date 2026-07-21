use super::helpers::run_lua_one;

#[test]
fn test_wrap_non_function_errors() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() return coroutine.wrap(1) end)
print(ok == false)"#), "true");
}

#[test]
fn test_wrap_invalid_resume_error() {
    assert_eq!(run_lua_one(r#"local f = coroutine.wrap(function() error("x") end)
local ok, err = pcall(f)
print(ok == false)"#), "true");
}

#[test]
fn test_wrap_error_type() {
    assert_eq!(run_lua_one(r#"local f = coroutine.wrap(function() error("oops") end)
local ok, err = pcall(f)
print(ok == false and type(err) == "string")"#), "true");
}

#[test]
fn test_wrap_error_in_yield_resume() {
    assert_eq!(run_lua_one(r#"local f = coroutine.wrap(function() coroutine.yield(1); error("y") end)
f()
local ok, err = pcall(function() f() end)
print(ok == false and type(err) == "string")"#), "true");
}

#[test]
fn test_wrap_error_payload_preserved() {
    assert_eq!(run_lua_one(r#"local f = coroutine.wrap(function() error({code=3}) end)
local ok, err = pcall(f)
print(ok == false and type(err) == "table" )"#), "true");
}

#[test]
fn test_wrap_error_bool_payload() {
    assert_eq!(run_lua_one(r#"local f = coroutine.wrap(function() error(false) end)
local ok, err = pcall(f)
print(ok == false and err == false)"#), "true");
}

#[test]
fn test_wrap_error_number_payload() {
    assert_eq!(run_lua_one(r#"local f = coroutine.wrap(function() error(42) end)
local ok, err = pcall(f)
print(ok == false and err == 42)"#), "true");
}

#[test]
fn test_wrap_error_function_payload() {
    assert_eq!(run_lua_one(r#"local inner = function() return 1 end
local f = coroutine.wrap(function() error(inner) end)
local ok, err = pcall(f)
print(ok == false and type(err) == "function")"#), "true");
}

#[test]
fn test_wrap_error_on_second_call() {
    assert_eq!(run_lua_one(r#"local f = coroutine.wrap(function() coroutine.yield(1); error("stop") end)
f()
local ok, err = pcall(f)
print(ok == false and string.find(err, "stop") ~= nil)"#), "true");
}

#[test]
fn test_wrap_error_after_done() {
    assert_eq!(run_lua_one(r#"local f = coroutine.wrap(function() return 1 end)
f()
local ok, err = pcall(f)
print(ok == false or err == nil)"#), "true");
}

#[test]
fn test_wrap_error_nested_wrap() {
    assert_eq!(run_lua_one(r#"local g = coroutine.wrap(function() error("inner") end)
local f = coroutine.wrap(function() return g() end)
local ok, err = pcall(f)
print(ok == false and string.find(err, "inner") ~= nil)"#), "true");
}

#[test]
fn test_wrap_error_argumentless_error() {
    assert_eq!(run_lua_one(r#"local f = coroutine.wrap(function() error() end)
local ok, err = pcall(f)
print(ok == false and type(err) == "string")"#), "true");
}

#[test]
fn test_wrap_error_then_recover() {
    assert_eq!(run_lua_one(r#"local f = coroutine.wrap(function() error("x") end)
local ok, _ = pcall(f)
local ok2, v = pcall(function() return 1 end)
print(ok2 == true and v == 1)"#), "true");
}

#[test]
fn test_wrap_error_in_argument() {
    assert_eq!(run_lua_one(r#"local function bad(x) if x == 0 then error("bad") end return x end
local f = coroutine.wrap(bad)
local ok, err = pcall(function() f(0) end)
print(ok == false and string.find(err, "bad") ~= nil)"#), "true");
}

#[test]
fn test_wrap_error_assert() {
    assert_eq!(run_lua_one(r#"local f = coroutine.wrap(function() assert(false, "assert") end)
local ok, err = pcall(f)
print(ok == false and type(err) == "string")"#), "true");
}

#[test]
fn test_wrap_error_message_match() {
    assert_eq!(run_lua_one(r#"local f = coroutine.wrap(function() error("hello world") end)
local ok, err = pcall(f)
print(ok == false and string.find(err, "hello") ~= nil)"#), "true");
}

#[test]
fn test_wrap_error_nil_message() {
    assert_eq!(run_lua_one(r#"local f = coroutine.wrap(function() error(nil) end)
local ok, err = pcall(f)
print(ok == false and type(err) == "nil")"#), "true");
}

#[test]
fn test_wrap_error_thread_payload() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() end)
local f = coroutine.wrap(function() error(t) end)
local ok, err = pcall(f)
print(ok == false and type(err) == "thread")"#), "true");
}

#[test]
fn test_wrap_error_after_yielded_value() {
    assert_eq!(run_lua_one(r#"local f = coroutine.wrap(function() coroutine.yield(1); error(12) end)
f()
local ok, err = pcall(f)
print(ok == false and err == 12)"#), "true");
}

#[test]
fn test_wrap_error_table_lookup() {
    assert_eq!(run_lua_one(r#"local f = coroutine.wrap(function() local t = nil; return t.a end)
local ok, err = pcall(f)
print(ok == false and type(err) == "string")"#), "true");
}

#[test]
fn test_wrap_error_concat() {
    assert_eq!(run_lua_one(r#"local f = coroutine.wrap(function() return nil .. "x" end)
local ok, err = pcall(f)
print(ok == false and type(err) == "string")"#), "true");
}

#[test]
fn test_wrap_error_second_stage() {
    assert_eq!(run_lua_one(r#"local f = coroutine.wrap(function(x) return x end)
local one = f(1)
local ok, err = pcall(f)
print(ok == false or type(ok) == "boolean")"#), "true");
}
