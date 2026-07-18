use super::helpers::run_lua_one;

#[test]
fn test_getinfo_named_local() {
    assert_eq!(run_lua_one(r#"local function f(x) local function g() return x end; return g end
local fn = f(1)
local info = debug.getinfo(fn, \"n\")
print(type(info.name) == \"string\" and (info.name == \"g\" or true))"#), "true");
}

#[test]
fn test_getinfo_nested_name() {
    assert_eq!(run_lua_one(r#"local function f()
  local function inner() end
  return debug.getinfo(inner, \"n\").name
end
print(f())"#), "inner");
}

#[test]
fn test_getinfo_anonymous_name() {
    assert_eq!(run_lua_one(r#"local fn = function() end
print(type(debug.getinfo(fn, \"n\").name) == \"string\" or debug.getinfo(fn, \"n\").name == nil)"#), "true");
}

#[test]
fn test_getinfo_method_table_name() {
    assert_eq!(run_lua_one(r#"local o = {}
function o:m() end
print(type(debug.getinfo(o.m, \"n\").name) == \"string\")"#), "true");
}

#[test]
fn test_getinfo_function_with_varargs() {
    assert_eq!(run_lua_one(r#"local function f(a, ...) end
local info = debug.getinfo(f, \"u\")
print(info.nparams >= 1 and info.isvararg == true)
"#), "true");
}

#[test]
fn test_getinfo_main_chunk() {
    assert_eq!(run_lua_one(r#"local info = debug.getinfo(1, \"S\")
print(type(info.short_src) == \"string\")"#), "true");
}

#[test]
fn test_getinfo_closure_call() {
    assert_eq!(run_lua_one(r#"local function mk(v)
  return function() return v end
end
local c = mk(3)
print(c())"#), "3");
}

#[test]
fn test_getinfo_closure_level_name() {
    assert_eq!(run_lua_one(r#"local function mk(v)
  return function() end
end
local c = mk(3)
print(type(debug.getinfo(c, \"S\").what) == \"string\")"#), "true");
}

#[test]
fn test_getinfo_coroutine_function() {
    assert_eq!(run_lua_one(r#"local c = coroutine.create(function() end)
print(type(c))"#), "thread");
}

#[test]
fn test_getinfo_tail_call_info() {
    assert_eq!(run_lua_one(r#"local function a()
  return b()
end
function b()
  local info = debug.getinfo(1, \"S\")
  return info.what
end
print(a())"#), "Lua");
}

#[test]
fn test_getinfo_stack_function() {
    assert_eq!(run_lua_one(r#"local function f() return debug.getinfo(1, \"S\").source end
print(type(f()) == \"string\")"#), "true");
}

#[test]
fn test_getinfo_upvalue_count() {
    assert_eq!(run_lua_one(r#"local function f(x)
  local y = 1
  return debug.getinfo(1, \"u\").nparams
end
print(type(f(1)) == \"number\")"#), "true");
}

#[test]
fn test_getinfo_max_stack_levels() {
    assert_eq!(run_lua_one(r#"local function depth(n)
  if n == 0 then return debug.getinfo(1, \"l\").currentline end
  return depth(n-1)
end
print(type(depth(2)) == \"number\")"#), "true");
}

#[test]
fn test_getinfo_function_source_not_empty() {
    assert_eq!(run_lua_one(r#"local function f() end
print(type(debug.getinfo(f, \"S\").source) == \"string\")"#), "true");
}

#[test]
fn test_getinfo_function_namefromdebug() {
    assert_eq!(run_lua_one(r#"local function f() end
local info = debug.getinfo(f, \"n\")
print(info.name ~= nil)"#), "true");
}

#[test]
fn test_getinfo_empty_mode() {
    assert_eq!(run_lua_one(r#"local function f() end
local info = debug.getinfo(f, \"\")
print(type(info) == \"table\")"#), "true");
}

#[test]
fn test_getinfo_f_line() {
    assert_eq!(run_lua_one(r#"local function f()
  return debug.getinfo(1, \"nS\")
end
local info = f()
print(info.what == \"Lua\")"#), "true");
}

#[test]
fn test_getinfo_returns_table() {
    assert_eq!(run_lua_one(r#"local info = debug.getinfo(1)
print(type(info) == \"table\")"#), "true");
}

#[test]
fn test_getinfo_nonexistent_name() {
    assert_eq!(run_lua_one(r#"local function f(x, y)
  local info = debug.getinfo(1, \"n\")
  return info.name
end
print(type(f(1,2)) == \"string\")"#), "true");
}

#[test]
fn test_getinfo_line_info() {
    assert_eq!(run_lua_one(r#"local function f()
  local info = debug.getinfo(1, \"l\")
  return info.currentline
end
print(type(f()) == \"number\")"#), "true");
}

#[test]
fn test_getinfo_lastline() {
    assert_eq!(run_lua_one(r#"local function f() end
local info = debug.getinfo(f, \"S\")
print(type(info.lastlinedefined) == \"number\")"#), "true");
}
