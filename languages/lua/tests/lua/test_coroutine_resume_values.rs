use super::helpers::run_lua_one;

#[test]
fn test_resume_value_sum() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() return 1 + 2 end)
local ok, v = coroutine.resume(t)
print(ok and v == 3)"#), "true");
}

#[test]
fn test_resume_value_string_concat() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() return \"a\" .. \"b\" end)
local ok, v = coroutine.resume(t)
print(ok and v == \"ab\")"#), "true");
}

#[test]
fn test_resume_value_boolean() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() return false end)
local ok, v = coroutine.resume(t)
print(ok and v == false)"#), "true");
}

#[test]
fn test_resume_value_table() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() return {a = 1} end)
local ok, v = coroutine.resume(t)
print(ok and type(v) == \"table\")"#), "true");
}

#[test]
fn test_resume_value_function() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() return function() end end)
local ok, v = coroutine.resume(t)
print(ok and type(v) == \"function\")"#), "true");
}

#[test]
fn test_resume_value_nil_and_more() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() return nil, 1, 2 end)
local ok, v = coroutine.resume(t)
print(ok and v == nil)"#), "true");
}

#[test]
fn test_resume_value_numeric_string_mix() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() return 1, \"x\" end)
local ok, v = coroutine.resume(t)
print(ok and v == 1)"#), "true");
}

#[test]
fn test_resume_value_math_abs() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() return math.abs(-10) end)
local ok, v = coroutine.resume(t)
print(ok and v == 10)"#), "true");
}

#[test]
fn test_resume_value_len() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() local s = \"abc\"; return #s end)
local ok, v = coroutine.resume(t)
print(ok and v == 3)"#), "true");
}

#[test]
fn test_resume_yield_value() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() coroutine.yield(5); return 8 end)
local ok1, v1 = coroutine.resume(t)
local ok2, v2 = coroutine.resume(t)
print(ok1 and v1 == 5 and ok2 and v2 == 8)"#), "true");
}

#[test]
fn test_resume_send_argument() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function(v) return v * 3 end)
local ok1, y = coroutine.resume(t, 3)
print(ok1 and y == 9)"#), "true");
}

#[test]
fn test_resume_chain_send_argument() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function(x)
  local y = coroutine.yield(x + 1)
  return y + 1
end)
coroutine.resume(t, 4)
local ok, v = coroutine.resume(t, 10)
print(ok and v == 11)"#), "true");
}

#[test]
fn test_resume_string_payload() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() return \"v\" end)
local ok, v = coroutine.resume(t)
print(ok and v == \"v\")"#), "true");
}

#[test]
fn test_resume_thread_payload() {
    assert_eq!(run_lua_one(r#"local other = coroutine.create(function() return 1 end)
local t = coroutine.create(function() return other end)
local ok, v = coroutine.resume(t)
print(ok and type(v) == \"thread\")"#), "true");
}

#[test]
fn test_resume_status_running_to_suspended() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() coroutine.yield(true) end)
local ok, v = coroutine.resume(t)
print(ok and v == true and coroutine.status(t) == \"suspended\")"#), "true");
}

#[test]
fn test_resume_value_table_index() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() local x = {a=7}; return x.a end)
local ok, v = coroutine.resume(t)
print(ok and v == 7)"#), "true");
}

#[test]
fn test_resume_value_nested_func() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() return (function() return 4 end)() end)
local ok, v = coroutine.resume(t)
print(ok and v == 4)"#), "true");
}

#[test]
fn test_resume_false_yield() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() coroutine.yield(false); return true end)
local ok1, v1 = coroutine.resume(t)
local ok2, v2 = coroutine.resume(t)
print(ok1 and v1 == false and ok2 and v2 == true)"#), "true");
}

#[test]
fn test_resume_conditional_return() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() local x = 1; if x > 0 then return 1 else return 0 end end)
local ok, v = coroutine.resume(t)
print(ok and v == 1)"#), "true");
}

#[test]
fn test_resume_compare_value() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() return 10 > 3 end)
local ok, v = coroutine.resume(t)
print(ok and v == true)"#), "true");
}
