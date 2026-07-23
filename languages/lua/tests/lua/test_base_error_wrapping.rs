use super::helpers::run_lua_one;

#[test]
fn test_error_object_preserved() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() error("x") end)
print(type(err) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_error_numeric_payload() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() error(42) end)
print(type(err) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_error_boolean_payload() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() error(true) end)
print(type(err) == "boolean")"#
        ),
        "true"
    );
}

#[test]
fn test_error_table_payload() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() error({code = 1}) end)
print(type(err) == "table")"#
        ),
        "true"
    );
}

#[test]
fn test_error_function_payload() {
    assert_eq!(
        run_lua_one(
            r#"local function marker() end
local ok, err = pcall(function() error(marker) end)
print(type(err) == "function")"#
        ),
        "true"
    );
}

#[test]
fn test_error_userdata_payload() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() error(coroutine.create(function() end)) end)
print(type(err) == "thread")"#
        ),
        "true"
    );
}

#[test]
fn test_error_with_level_and_message() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() error("wrapped", 2) end)
print(string.find(err, "wrapped") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_error_nested_wrap() {
    assert_eq!(
        run_lua_one(
            r#"local function fail() error("inner") end
local function boom() fail() end
local ok, err = pcall(boom)
print(type(err) == "string" and string.find(err, "inner") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_error_then_assert() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() assert(false, "boom") end)
print(type(err) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_error_pcall_boundary() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() error("x", 99) end)
print(ok == false)"#
        ),
        "true"
    );
}

#[test]
fn test_error_multiple_payloads() {
    assert_eq!(
        run_lua_one(
            r#"local function f() return 1 end
local ok, err = pcall(function() local first = f(); error(first) end)
print(ok == false and type(err) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_error_payload_from_vararg() {
    assert_eq!(
        run_lua_one(
            r#"local function f() return 2 end
local ok, err = pcall(function()
  local b = f()
  error(b)
end)
print(type(err) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_error_message_with_format() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() error(string.format("v=%d", 3)) end)
print(string.find(err, "v=3") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_error_inside_if() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() if false then error("no") else error("yes") end end)
print(string.find(err, "yes") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_error_inside_repeat() {
    assert_eq!(
        run_lua_one(
            r#"local i = 0
local ok, err = pcall(function()
  i = i + 1
  if i == 1 then error("repeat") end
end)
print(string.find(err, "repeat") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_error_catch_and_recover() {
    assert_eq!(
        run_lua_one(
            r#"local ok = pcall(function() error("recover") end)
print(ok == false)"#
        ),
        "true"
    );
}

#[test]
fn test_error_chained_calls() {
    assert_eq!(
        run_lua_one(
            r#"local function a() return b() end
local function b() return c() end
local function c() error("chain") end
local ok, err = pcall(a)
print(string.find(err, "chain") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_error_with_table_field() {
    assert_eq!(
        run_lua_one(
            r#"local payload = {tag = "e"}
local ok, err = pcall(function() error(payload) end)
print(type(err) == "table")"#
        ),
        "true"
    );
}

#[test]
fn test_error_level_override() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() error("lvl", 3) end)
print(string.find(err, "lvl") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_error_then_pcall_false() {
    assert_eq!(
        run_lua_one(
            r#"local ok, _ = pcall(function() error("x") end)
print(ok)"#
        ),
        "false"
    );
}

#[test]
fn test_error_message_default_when_nil() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() error(nil) end)
print(type(err) == "nil")"#
        ),
        "true"
    );
}

#[test]
fn test_error_message_type_boolean() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() error(false) end)
print(err == false)"#
        ),
        "true"
    );
}
