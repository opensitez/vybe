use super::helpers::run_lua_one;

#[test]
fn test_running_inside_body_is_mainthread_false() {
    assert_eq!(
        run_lua_one(
            r#"local threadState = { false }
local t = coroutine.create(function()
  local t, isMain = coroutine.running()
  threadState[1] = (isMain == false and t ~= nil)
end)
coroutine.resume(t)
print(threadState[1])"#
        ),
        "true"
    );
}

#[test]
fn test_running_main_thread_true() {
    assert_eq!(
        run_lua_one(
            r#"local thread, isMain = coroutine.running()
print(isMain == true)"#
        ),
        "true"
    );
}

#[test]
fn test_running_thread_return_type() {
    assert_eq!(
        run_lua_one(
            r#"local t = coroutine.create(function()
  local thr = coroutine.running()
  print(type(thr))
end)
print(coroutine.resume(t) == true)"#
        ),
        "true"
    );
}

#[test]
fn test_running_before_resume() {
    assert_eq!(
        run_lua_one(
            r#"local t = coroutine.create(function() end)
print(coroutine.running() ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_running_inside_two_threads() {
    assert_eq!(
        run_lua_one(
            r#"local inner_state
local t1 = coroutine.create(function()
  inner_state = coroutine.running() ~= nil
end)
local t2 = coroutine.create(function() coroutine.resume(t1) end)
coroutine.resume(t2)
print(inner_state and "ok" or "no")"#
        ),
        "ok"
    );
}

#[test]
fn test_running_returns_thread_not_main_when_in_coroutine() {
    assert_eq!(
        run_lua_one(
            r#"local stored
local t = coroutine.create(function()
  local th, main = coroutine.running()
  stored = (main == false)
  print(main)
end)
coroutine.resume(t)
print(stored == true)"#
        ),
        "true"
    );
}

#[test]
fn test_running_before_yield() {
    assert_eq!(
        run_lua_one(
            r#"local t = coroutine.create(function()
  local _, main = coroutine.running()
  coroutine.yield(main)
end)
local ok, isMain = coroutine.resume(t)
print(ok == true and isMain == false)"#
        ),
        "true"
    );
}

#[test]
fn test_running_status_during_dead() {
    assert_eq!(
        run_lua_one(
            r#"local th, main = coroutine.running()
print(main == true)"#
        ),
        "true"
    );
}

#[test]
fn test_running_after_resume() {
    assert_eq!(
        run_lua_one(
            r#"local seen
local t = coroutine.create(function()
  local _, main = coroutine.running()
  seen = main
  return 1
end)
coroutine.resume(t)
print(seen == false)"#
        ),
        "true"
    );
}

#[test]
fn test_running_nested_yield() {
    assert_eq!(
        run_lua_one(
            r#"local seen = false
local t = coroutine.create(function()
  local _, main = coroutine.running()
  coroutine.yield(main)
  seen = true
end)
coroutine.resume(t)
coroutine.resume(t)
print(seen)
"#
        ),
        "true"
    );
}

#[test]
fn test_running_coroutine_type() {
    assert_eq!(
        run_lua_one(
            r#"local t = coroutine.create(function()
  local thread = coroutine.running()
  print(type(thread))
end)
coroutine.resume(t)
print(coroutine.status(t) == "dead")"#
        ),
        "true"
    );
}

#[test]
fn test_running_from_function_returned() {
    assert_eq!(
        run_lua_one(
            r#"local function f()
  local _, main = coroutine.running()
  return main
end
local t = coroutine.create(f)
local ok, v = coroutine.resume(t)
print(ok and v == false)"#
        ),
        "true"
    );
}

#[test]
fn test_running_after_error() {
    assert_eq!(
        run_lua_one(
            r#"local state
local t = coroutine.create(function()
  state = coroutine.running()
  error("x")
end)
coroutine.resume(t)
print(state ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_running_two_runs_same_thread() {
    assert_eq!(
        run_lua_one(
            r#"local t = coroutine.create(function()
  local _, main = coroutine.running()
  return main and 1 or 2
end)
local _, a = coroutine.resume(t)
local t2 = coroutine.create(function() return coroutine.running() ~= nil end)
local _, b = coroutine.resume(t2)
print(a == 2 and b == true)"#
        ),
        "true"
    );
}

#[test]
fn test_running_with_arguments() {
    assert_eq!(
        run_lua_one(
            r#"local t = coroutine.create(function(x)
  local _, main = coroutine.running()
  return (main and 0 or x)
end)
local ok, v = coroutine.resume(t, 9)
print(ok and v == 9)"#
        ),
        "true"
    );
}

#[test]
fn test_running_after_yield_then_resume() {
    assert_eq!(
        run_lua_one(
            r#"local t = coroutine.create(function()
  local _, main = coroutine.running()
  local v = coroutine.yield(main)
  return v
end)
local ok, first = coroutine.resume(t)
local ok2, second = coroutine.resume(t, true)
print(ok and ok2 and type(first) == "boolean" and second == true)"#
        ),
        "true"
    );
}

#[test]
fn test_running_multiple_coroutines() {
    assert_eq!(
        run_lua_one(
            r#"local c1 = coroutine.create(function() return coroutine.running() ~= nil end)
local c2 = coroutine.create(function() return coroutine.running() ~= nil end)
local _, a = coroutine.resume(c1)
local _, b = coroutine.resume(c2)
print(a == true and b == true)"#
        ),
        "true"
    );
}

#[test]
fn test_running_function_name() {
    assert_eq!(
        run_lua_one(
            r#"local t = coroutine.create(function() coroutine.running() end)
local _, first = coroutine.resume(t)
print(first == nil)"#
        ),
        "true"
    );
}

#[test]
fn test_running_after_completion() {
    assert_eq!(
        run_lua_one(
            r#"local done = false
local t = coroutine.create(function()
  local _, main = coroutine.running()
  done = not main
end)
coroutine.resume(t)
print(done == true)"#
        ),
        "true"
    );
}

#[test]
fn test_running_nested_coroutine_create() {
    assert_eq!(
        run_lua_one(
            r#"local t = coroutine.create(function()
  local inner = coroutine.create(function()
    local _, main = coroutine.running()
    return main
  end)
  local ok, v = coroutine.resume(inner)
  return ok and v == false
end)
local ok, v = coroutine.resume(t)
print(ok and v == true)"#
        ),
        "true"
    );
}

#[test]
fn test_running_when_no_coroutine() {
    assert_eq!(
        run_lua_one(
            r#"local _, main = coroutine.running()
print(main)"#
        ),
        "true"
    );
}
