use super::helpers::run_lua_one;

#[test]
fn test_initial_status_is_suspended() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() return 1 end)
print(coroutine.status(t) == \"suspended\")"#), "true");
}

#[test]
fn test_running_status_after_create() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() return 1 end)
print(coroutine.status(t) == \"suspended\")"#), "true");
}

#[test]
fn test_status_after_resume_dead_on_return() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() return 1 end)
coroutine.resume(t)
print(coroutine.status(t) == \"dead\")"#), "true");
}

#[test]
fn test_status_after_error_dead() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() error(\"x\") end)
coroutine.resume(t)
print(coroutine.status(t) == \"dead\")"#), "true");
}

#[test]
fn test_status_after_yield_running() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() coroutine.yield(1) end)
coroutine.resume(t)
print(coroutine.status(t) == \"suspended\")"#), "true");
}

#[test]
fn test_status_running_during_body() {
    assert_eq!(run_lua_one(r#"local active = false
local t = coroutine.create(function()
  active = (coroutine.running() ~= nil)
end)
coroutine.resume(t)
print(active)
"#), "true");
}

#[test]
fn test_status_running_returns_main_thread_in_main() {
    assert_eq!(run_lua_one(r#"local thread, ismain = coroutine.running()
print(ismain == true)"#), "true");
}

#[test]
fn test_status_yield_then_dead() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() coroutine.yield(1) end)
coroutine.resume(t)
coroutine.resume(t)
print(coroutine.status(t) == \"dead\")"#), "true");
}

#[test]
fn test_status_create_multiple_threads() {
    assert_eq!(run_lua_one(r#"local a = coroutine.create(function() return 1 end)
local b = coroutine.create(function() return 2 end)
print(coroutine.status(a) .. \"/\" .. coroutine.status(b))"#), "suspended/suspended");
}

#[test]
fn test_status_create_zero_arg() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function(a) return a end)
print(coroutine.status(t) == \"suspended\")"#), "true");
}

#[test]
fn test_status_resume_then_status() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function(x) return x end)
print((function()
  local ok, _ = coroutine.resume(t, 7)
  return coroutine.status(t)
end)() == \"dead\")"#), "true");
}

#[test]
fn test_status_yield_and_send() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function(v) coroutine.yield(v) return v end)
coroutine.resume(t, 5)
print(coroutine.status(t) == \"suspended\")"#), "true");
}

#[test]
fn test_status_alive_after_yield() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() coroutine.yield(1) end)
coroutine.resume(t)
print(coroutine.status(t) ~= \"dead\")"#), "true");
}

#[test]
fn test_status_after_extra_resume_failure() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() return 1 end)
coroutine.resume(t)
coroutine.resume(t)
print(coroutine.status(t) == \"dead\")"#), "true");
}

#[test]
fn test_status_with_nested_resume() {
    assert_eq!(run_lua_one(r#"local outer
outer = coroutine.create(function()
  local ok, innerStatus = pcall(function() return coroutine.status(inner) end)
  return ok and innerStatus == nil
end)
local inner = coroutine.create(function() return 1 end)
coroutine.resume(outer)
print(1)")#), "1");
}

#[test]
fn test_status_for_two_depth() {
    assert_eq!(run_lua_one(r#"local t1 = coroutine.create(function() return 1 end)
local t2 = coroutine.create(function() return coroutine.resume(t1) end)
coroutine.resume(t2)
print(coroutine.status(t2) == \"dead\")"#), "true");
}

#[test]
fn test_status_empty_resume_error() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() return 1/0 end)
local ok, _ = coroutine.resume(t)
print(coroutine.status(t) == \"dead\")"#), "true");
}

#[test]
fn test_status_nil_payload() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() return nil end)
coroutine.resume(t)
print(coroutine.status(t) == \"dead\")"#), "true");
}

#[test]
fn test_status_bool_payload() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() return false end)
coroutine.resume(t)
print(coroutine.status(t) == \"dead\")"#), "true");
}

#[test]
fn test_status_yield_payload() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function() coroutine.yield(0) end)
coroutine.resume(t)
print(coroutine.status(t) == \"suspended\")"#), "true");
}

#[test]
fn test_status_after_multiple_yields() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function()
  coroutine.yield(1)
  coroutine.yield(2)
end)
coroutine.resume(t)
coroutine.resume(t, 0)
print(coroutine.status(t) == \"dead\")"#), "true");
}

#[test]
fn test_status_after_argument_resume() {
    assert_eq!(run_lua_one(r#"local t = coroutine.create(function(a)
  local b = coroutine.yield(a)
  return b
end)
coroutine.resume(t, 1)
coroutine.resume(t, 2)
print(coroutine.status(t) == \"dead\")"#), "true");
}
