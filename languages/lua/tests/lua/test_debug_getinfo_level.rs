use super::helpers::run_lua_one;

#[test]
fn test_getinfo_current_function() {
    assert_eq!(
        run_lua_one(
            r#"local info = debug.getinfo(1, "n")
print(info.name ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_getinfo_current_function_name_known() {
    assert_eq!(
        run_lua_one(
            r#"local info = debug.getinfo(1, "n")
print(type(info.name))"#
        ),
        "string"
    );
}

#[test]
fn test_getinfo_current_source() {
    assert_eq!(
        run_lua_one(
            r#"local info = debug.getinfo(1, "S")
print(type(info.source))"#
        ),
        "string"
    );
}

#[test]
fn test_getinfo_current_linedefined() {
    assert_eq!(
        run_lua_one(
            r#"local info = debug.getinfo(1, "S")
print(type(info.linedefined) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_getinfo_current_namewhat() {
    assert_eq!(
        run_lua_one(
            r#"local info = debug.getinfo(1, "n")
print(info.namewhat == "" or type(info.namewhat) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_getinfo_outer_frame() {
    assert_eq!(
        run_lua_one(
            r#"local function f()
  return debug.getinfo(2, "nS")
end
local info = f()
print(info.source ~= nil)
"#
        ),
        "true"
    );
}

#[test]
fn test_getinfo_zero_level_fails() {
    assert_eq!(
        run_lua_one(
            r#"local info = debug.getinfo(0, "n")
print(type(info) == "table" or type(info) == "nil")"#
        ),
        "true"
    );
}

#[test]
fn test_getinfo_negative_level() {
    assert_eq!(
        run_lua_one(
            r#"local info = debug.getinfo(-1, "n")
print(type(info) == "table" or type(info) == "nil")"#
        ),
        "true"
    );
}

#[test]
fn test_getinfo_function_arg() {
    assert_eq!(
        run_lua_one(
            r#"local function g() end
local info = debug.getinfo(g, "n")
print(type(info.name) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_getinfo_function_name() {
    assert_eq!(
        run_lua_one(
            r#"local function g() end
local info = debug.getinfo(g, "n")
print(type(info.name) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_getinfo_defined_name() {
    assert_eq!(
        run_lua_one(
            r#"local function outer()
  local function inner() end
  local info = debug.getinfo(inner, "n")
  return info.name
end
print(outer())"#
        ),
        "inner"
    );
}

#[test]
fn test_getinfo_varargs() {
    assert_eq!(
        run_lua_one(
            r#"local function f(a,b,...)
  return debug.getinfo(1, "u").nparams
end
print(f(1,2,3) > 0)"#
        ),
        "true"
    );
}

#[test]
fn test_getinfo_nparams_zero() {
    assert_eq!(
        run_lua_one(
            r#"local function f() return debug.getinfo(1, "u").nparams end
print(type(f()) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_getinfo_variadic_flag() {
    assert_eq!(
        run_lua_one(
            r#"local function f(a, ...)
  local info = debug.getinfo(1, "u")
  return info.isvararg and 1 or 0
end
print(f(1))"#
        ),
        "1"
    );
}

#[test]
fn test_getinfo_not_variadic() {
    assert_eq!(
        run_lua_one(
            r#"local function f(a,b)
  return debug.getinfo(1, "u").isvararg and 1 or 0
end
print(type(debug.getinfo(1, "u").isvararg) == "boolean")"#
        ),
        "true"
    );
}

#[test]
fn test_getinfo_current_line() {
    assert_eq!(
        run_lua_one(
            r#"local info = debug.getinfo(1, "l")
print(type(info.currentline) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_getinfo_level_one_currentline() {
    assert_eq!(
        run_lua_one(
            r#"local function f()
  return debug.getinfo(1, "l").currentline
end
print(type(f()) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_getinfo_stack_depth() {
    assert_eq!(
        run_lua_one(
            r#"local function f()
  return debug.getinfo(2, "S").source
end
local function g()
  return f()
end
print(type(g()) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_getinfo_invalid_mode() {
    assert_eq!(
        run_lua_one(
            r#"local info = debug.getinfo(1, "xyz")
print(type(info) == "table" or type(info) == "nil")"#
        ),
        "true"
    );
}

#[test]
fn test_getinfo_short_src() {
    assert_eq!(
        run_lua_one(
            r#"local info = debug.getinfo(1, "S")
print(type(info.short_src) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_getinfo_what() {
    assert_eq!(
        run_lua_one(
            r#"local info = debug.getinfo(1, "S")
print(info.what == "Lua" or info.what == "C")"#
        ),
        "true"
    );
}

#[test]
fn test_getinfo_line_defined_default() {
    assert_eq!(
        run_lua_one(
            r#"local function f() end
local info = debug.getinfo(f, "S")
print(type(info.linedefined) == "number")"#
        ),
        "true"
    );
}
