use super::helpers::run_lua_one;

#[test]
fn test_getinfo_varargs_count() {
    assert_eq!(run_lua_one(r#"local function f(a,b,...)
  return debug.getinfo(1, "u").nparams
end
print(f(1,2,3))"#), "2");
}

#[test]
fn test_getinfo_no_params() {
    assert_eq!(run_lua_one(r#"local function f()
  return debug.getinfo(1, "u").nparams
end
print(f())"#), "0");
}

#[test]
fn test_getinfo_vararg_true() {
    assert_eq!(run_lua_one(r#"local function f(a, ...)
  return debug.getinfo(1, "u").isvararg
end
print(f(1) == true)"#), "true");
}

#[test]
fn test_getinfo_not_vararg() {
    assert_eq!(run_lua_one(r#"local function f(a,b)
  return debug.getinfo(1, "u").isvararg
end
print(f(1,2) == false)"#), "true");
}

#[test]
fn test_getinfo_varargs_multiple_calls() {
    assert_eq!(run_lua_one(r#"local function f(a, ...)
  local info = debug.getinfo(1, "u")
  return info.nparams == 1 and info.isvararg == true
end
print(f(1,2,3,4))"#), "true");
}

#[test]
fn test_getinfo_varargs_in_method() {
    assert_eq!(run_lua_one(r#"local t = {}
function t:m(a, ...)
  local info = debug.getinfo(1, "u").isvararg
  return info
end
print(t:m(1,2) == true)"#), "true");
}

#[test]
fn test_getinfo_varargs_outer_function() {
    assert_eq!(run_lua_one(r#"local function wrap()
  local function inner(a, ...)
    return debug.getinfo(1, "u").isvararg
  end
  return inner(1,2)
end
print(wrap() == true)"#), "true");
}

#[test]
fn test_getinfo_varargs_boolean_param() {
    assert_eq!(run_lua_one(r#"local function f(flag, ...)
  return debug.getinfo(1, "u").isvararg
end
print(f(true) == true)"#), "true");
}

#[test]
fn test_getinfo_varargs_nested() {
    assert_eq!(run_lua_one(r#"local function f(a, ...)
  local function g(b, ...)
    return debug.getinfo(2, "u").isvararg
  end
  return g(a)
end
print(f(1) == true)"#), "true");
}

#[test]
fn test_getinfo_varargs_named() {
    assert_eq!(run_lua_one(r#"local function f(a, ...)
  local info = debug.getinfo(1, "u")
  return info.nparams
end
print(f(1,2,3) == 1)"#), "true");
}

#[test]
fn test_getinfo_varargs_zero_arg_call() {
    assert_eq!(run_lua_one(r#"local function f(...)
  local info = debug.getinfo(1, "u")
  return info.nparams
end
print(f() == 0)"#), "true");
}

#[test]
fn test_getinfo_varargs_level_source() {
    assert_eq!(run_lua_one(r#"local function f(a,...)
  return debug.getinfo(1, "S").what
end
print(f(1))"#), "Lua");
}

#[test]
fn test_getinfo_varargs_mode_u() {
    assert_eq!(run_lua_one(r#"local function f(a,...)
  local info = debug.getinfo(1, "u")
  return type(info.nparams)
end
print(f(1))"#), "number");
}

#[test]
fn test_getinfo_varargs_with_function_arg() {
    assert_eq!(run_lua_one(r#"local function inner(a,...)
  return debug.getinfo(1, "u")
end
print(type(inner(1,2,3).nparams) == "number")"#), "true");
}

#[test]
fn test_getinfo_varargs_string_mode() {
    assert_eq!(run_lua_one(r#"local function f(a,...)
  local info = debug.getinfo(1, "u")
  return type(info.isvararg)
end
print(f("x"))"#), "boolean");
}

#[test]
fn test_getinfo_varargs_table_mode() {
    assert_eq!(run_lua_one(r#"local function f(...)
  local info = debug.getinfo(1, "u")
  print(type(info) == "table")
end
f()"#), "true");
}

#[test]
fn test_getinfo_varargs_from_anonymous() {
    assert_eq!(run_lua_one(r#"print((function(...)
  local info = debug.getinfo(1, "u")
  return tostring(info.isvararg)
end)() == "true")"#), "true");
}

#[test]
fn test_getinfo_varargs_frame_two() {
    assert_eq!(run_lua_one(r#"local function outer(...)
  local function inner(...)
    return debug.getinfo(2, "u").isvararg
  end
  return inner(1,2)
end
print(outer(3,4))"#), "true");
}

#[test]
fn test_getinfo_varargs_in_loop() {
    assert_eq!(run_lua_one(r#"local function f(a,...)
  for i=1,1 do
    local info = debug.getinfo(1, "u")
    return info.nparams
  end
end
print(f(9,1,2,3) == 1)"#), "true");
}

#[test]
fn test_getinfo_varargs_in_coroutine() {
    assert_eq!(run_lua_one(r#"local function f(a,...)
  return debug.getinfo(1, "u").isvararg
end
local co = coroutine.create(function()
  return f(1,2)
end)
local ok, v = coroutine.resume(co)
print(ok and v == true)"#), "true");
}

#[test]
fn test_getinfo_varargs_named_params() {
    assert_eq!(run_lua_one(r#"local function f(a,b,...)
  local info = debug.getinfo(1, "u")
  return info.nparams
end
print(f(1,2,3))"#), "2");
}

#[test]
fn test_getinfo_varargs_tail_call() {
    assert_eq!(run_lua_one(r#"local function a(...)
  return b(...)
end
function b(...)
  return debug.getinfo(1, "u").isvararg
end
print(a(1,2,3) == true)"#), "true");
}
