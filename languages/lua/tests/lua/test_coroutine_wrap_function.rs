use super::helpers::run_lua_one;

#[test]
fn test_wrap_creates_callable() {
    assert_eq!(
        run_lua_one(
            r#"local f = coroutine.wrap(function() return 1 end)
print(type(f))"#
        ),
        "function"
    );
}

#[test]
fn test_wrap_invocation_returns_value() {
    assert_eq!(
        run_lua_one(
            r#"local f = coroutine.wrap(function() return 2 end)
print(f())"#
        ),
        "2"
    );
}

#[test]
fn test_wrap_invocation_addition() {
    assert_eq!(
        run_lua_one(
            r#"local f = coroutine.wrap(function() return 1 + 2 end)
print(f())"#
        ),
        "3"
    );
}

#[test]
fn test_wrap_yield_value() {
    assert_eq!(
        run_lua_one(
            r#"local f = coroutine.wrap(function() coroutine.yield(7); return 9 end)
print(f())"#
        ),
        "7"
    );
}

#[test]
fn test_wrap_second_call_after_yield() {
    assert_eq!(
        run_lua_one(
            r#"local f = coroutine.wrap(function() local x = coroutine.yield(1); return x + 1 end)
f()
print(f(4))"#
        ),
        "5"
    );
}

#[test]
fn test_wrap_table_value() {
    assert_eq!(
        run_lua_one(
            r#"local f = coroutine.wrap(function() return {a = 2} end)
local v = f()
print(type(v))"#
        ),
        "table"
    );
}

#[test]
fn test_wrap_error_propagates() {
    assert_eq!(
        run_lua_one(
            r#"local f = coroutine.wrap(function() error("no") end)
local ok, err = pcall(f)
print(ok == false)"#
        ),
        "true"
    );
}

#[test]
fn test_wrap_argument_passthrough() {
    assert_eq!(
        run_lua_one(
            r#"local f = coroutine.wrap(function(x) return x * 2 end)
print(f(6))"#
        ),
        "12"
    );
}

#[test]
fn test_wrap_closure_capture() {
    assert_eq!(
        run_lua_one(
            r#"local n = 10
local f = coroutine.wrap(function() return n end)
print(f())"#
        ),
        "10"
    );
}

#[test]
fn test_wrap_nested_wrap() {
    assert_eq!(
        run_lua_one(
            r#"local f = coroutine.wrap(function() return (coroutine.wrap(function() return 3 end))() end)
print(f())"#
        ),
        "3"
    );
}

#[test]
fn test_wrap_type_check() {
    assert_eq!(
        run_lua_one(
            r#"local f = coroutine.wrap(function() return 1 end)
print(type(f))"#
        ),
        "function"
    );
}

#[test]
fn test_wrap_false_return() {
    assert_eq!(
        run_lua_one(
            r#"local f = coroutine.wrap(function() return false end)
local v = f()
print(v == false)"#
        ),
        "true"
    );
}

#[test]
fn test_wrap_string_return() {
    assert_eq!(
        run_lua_one(
            r#"local f = coroutine.wrap(function() return "ok" end)
print(f())"#
        ),
        "ok"
    );
}

#[test]
fn test_wrap_nil_return() {
    assert_eq!(
        run_lua_one(
            r#"local f = coroutine.wrap(function() return nil end)
local v = f()
print(v == nil)"#
        ),
        "true"
    );
}

#[test]
fn test_wrap_math_return() {
    assert_eq!(
        run_lua_one(
            r#"local f = coroutine.wrap(function() return math.sqrt(16) end)
print(f())"#
        ),
        "4"
    );
}

#[test]
fn test_wrap_boolean_guard() {
    assert_eq!(
        run_lua_one(
            r#"local f = coroutine.wrap(function() return not false end)
print(f() == true)"#
        ),
        "true"
    );
}

#[test]
fn test_wrap_conditional_return() {
    assert_eq!(
        run_lua_one(
            r#"local f = coroutine.wrap(function() if true then return 1 else return 2 end end)
print(f())"#
        ),
        "1"
    );
}

#[test]
fn test_wrap_yield_and_return() {
    assert_eq!(
        run_lua_one(
            r#"local f = coroutine.wrap(function() coroutine.yield(9); return 10 end)
print(f())"#
        ),
        "9"
    );
}

#[test]
fn test_wrap_second_resume_value() {
    assert_eq!(
        run_lua_one(
            r#"local f = coroutine.wrap(function()
  local x = coroutine.yield(1)
  return x + 2
end)
f()
print(f(4))"#
        ),
        "6"
    );
}

#[test]
fn test_wrap_many_additions() {
    assert_eq!(
        run_lua_one(
            r#"local f = coroutine.wrap(function(a,b,c) return a+b+c end)
print(f(1,2,3))"#
        ),
        "6"
    );
}
