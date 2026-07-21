use super::helpers::run_lua_one;

#[test]
fn test_create_basic_thread() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() return 1 end)
print(type(t))"#), "thread");
}

#[test]
fn test_create_and_resume_simple() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() return 5 end)
local ok, v = coroutine.resume(t)
print(ok and v == 5)"#), "true");
}

#[test]
fn test_create_and_resume_multiple_values() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() return 1,2,3 end)
local ok, v = coroutine.resume(t)
print(ok and v == 1)"#), "true");
}

#[test]
fn test_create_resume_false_function() {
    assert_eq!(run_lua_one(r#"local ok, err = coroutine.resume(1)
print(ok == false)"#), "true");
}

#[test]
fn test_create_resume_no_function() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() coroutine.create(1) end)
print(ok == false)"#), "true");
}

#[test]
fn test_create_with_upvalue_capture() {
    assert_eq!(run_lua_one(r#"local x = 7
local t = coroutine.create(function() return x end)
local ok, v = coroutine.resume(t)
print(ok and v == 7)"#), "true");
}

#[test]
fn test_create_local_parameter() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function(a) return a * 2 end)
local ok, v = coroutine.resume(t, 4)
print(ok and v == 8)"#), "true");
}

#[test]
fn test_create_two_calls_after_dead() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() return 1 end)
local ok1 = coroutine.resume(t)
local ok2 = coroutine.resume(t)
print(ok1 ~= nil and ok2 == false)"#), "true");
}

#[test]
fn test_create_nested_resume() {
    assert_eq!(run_lua_one(r#"local inner = coroutine.create(function() return 2 end)
local function outer()
  return coroutine.resume(inner)
end
local ok, r = pcall(outer)
print(ok and type(r) == "boolean")"#), "true");
}

#[test]
fn test_create_resume_arg_passthrough() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function(v) return v end)
local ok, v = coroutine.resume(t, 11)
print(ok and v == 11)"#), "true");
}

#[test]
fn test_create_resume_false_from_function_error() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() error("x") end)
local ok, err = coroutine.resume(t)
print(ok == false)"#), "true");
}

#[test]
fn test_create_resume_nested_error_payload() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() error("inner") end)
local ok, err = coroutine.resume(t)
print(ok == false and string.find(err, "inner") ~= nil)"#), "true");
}

#[test]
fn test_create_resume_status_running() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() coroutine.yield(true) end)
local ok, v = coroutine.resume(t)
print(type(v) == "boolean" and ok)
"#), "true");
}

#[test]
fn test_create_resume_yield_count() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() coroutine.yield(1); return 2 end)
local ok, v = coroutine.resume(t)
local ok2, v2 = coroutine.resume(t)
print(ok and ok2 and v == 1 and v2 == 2)"#), "true");
}

#[test]
fn test_create_resume_after_error_status() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() error("boom") end)
local ok, _ = coroutine.resume(t)
local status = coroutine.status(t)
print(status == "dead")"#), "true");
}

#[test]
fn test_create_yielding_returns_value() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() coroutine.yield(9); return 3 end)
coroutine.resume(t)
local ok, v = coroutine.resume(t)
print(ok and v == 3)"#), "true");
}

#[test]
fn test_create_resumed_with_input() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function(x) coroutine.yield(x); return x + 1 end)
coroutine.resume(t, 4)
local ok, v = coroutine.resume(t, 0)
print(ok and v == 5)"#), "true");
}

#[test]
fn test_create_thread_type() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() return 1 end)
print(type(t))"#), "thread");
}

#[test]
fn test_create_resume_return_boolean() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() return false end)
local ok, v = coroutine.resume(t)
print(ok and v == false)"#), "true");
}

#[test]
fn test_create_resume_nil() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() return nil end)
local ok, v = coroutine.resume(t)
print(ok and v == nil)"#), "true");
}

#[test]
fn test_create_resume_complex_expression() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() return 3 * 4 end)
local ok, v = coroutine.resume(t)
print(ok and v == 12)"#), "true");
}
