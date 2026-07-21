use super::helpers::run_lua_one;

#[test]
fn test_xpcall_passes_function() {
    assert_eq!(run_lua_one(r#"local function ok() return 1 end
local function handler(err) return "handled" end
local ok, v = xpcall(ok, handler)
print(ok and v == 1)"#), "true");
}

#[test]
fn test_xpcall_handler_catches_error() {
    assert_eq!(run_lua_one(r#"local function bad() error("bad") end
local function handler(err) return "handled" end
local ok, v = xpcall(bad, handler)
print(ok == false and v == "handled")"#), "true");
}

#[test]
fn test_xpcall_handler_receives_error_payload() {
    assert_eq!(run_lua_one(r#"local function bad() error("boom") end
local function handler(err) return string.find(err, "boom") ~= nil end
local ok, v = xpcall(bad, handler)
print(ok == false and v == true)"#), "true");
}

#[test]
fn test_xpcall_handler_receives_table() {
    assert_eq!(run_lua_one(r#"local function bad() error({m = 2}) end
local function handler(err) return err.m end
local ok, v = xpcall(bad, handler)
print(ok == false and v == 2)"#), "true");
}

#[test]
fn test_xpcall_handler_with_args() {
    assert_eq!(run_lua_one(r#"local function good(x) return x * 2 end
local function handler(err) return 0 end
local ok, v = xpcall(good, handler, 4)
print(ok and v == 8)"#), "true");
}

#[test]
fn test_xpcall_error_with_zero() {
    assert_eq!(run_lua_one(r#"local function bad() error(0) end
local function handler(err) return err + 1 end
local ok, v = xpcall(bad, handler)
print(ok == false and v == 1)"#), "true");
}

#[test]
fn test_xpcall_error_with_false() {
    assert_eq!(run_lua_one(r#"local function bad() error(false) end
local function handler(err) return tostring(err) end
local ok, v = xpcall(bad, handler)
print(ok == false and v == "false")"#), "true");
}

#[test]
fn test_xpcall_error_with_true() {
    assert_eq!(run_lua_one(r#"local function bad() error(true) end
local function handler(err) return err and 1 or 0 end
local ok, v = xpcall(bad, handler)
print(ok == false and v == 1)"#), "true");
}

#[test]
fn test_xpcall_nested_ok() {
    assert_eq!(run_lua_one(r#"local function good() return 10 end
local function handler(err) return 0 end
local function outer() return xpcall(good, handler) end
local a, b = pcall(outer)
print(a == true and b == true and b == true)"#), "true");
}

#[test]
fn test_xpcall_nested_error() {
    assert_eq!(run_lua_one(r#"local function bad() error("bad") end
local function handler(err) return "handled" end
local function outer() return xpcall(bad, handler) end
local ok, status, val = pcall(outer)
print(ok and status == false and val == "handled")"#), "true");
}

#[test]
fn test_xpcall_multi_result() {
    assert_eq!(run_lua_one(r#"local function good() return 1, 2 end
local function handler(err) return 0 end
local ok, a = xpcall(good, handler)
print(ok and a == 1)"#), "true");
}

#[test]
fn test_xpcall_handler_boolean_result() {
    assert_eq!(run_lua_one(r#"local function bad() error("x") end
local function handler(err) return true end
local ok, v = xpcall(bad, handler)
print(ok == false and v == true)"#), "true");
}

#[test]
fn test_xpcall_handler_nil_result() {
    assert_eq!(run_lua_one(r#"local function bad() error("x") end
local function handler(err) return nil end
local ok, v = xpcall(bad, handler)
print(ok == false and v == nil)"#), "true");
}

#[test]
fn test_xpcall_with_function_returning_error() {
    assert_eq!(run_lua_one(r#"local function bad() error("x") end
local function handler(err) return error(err) end
local ok, v = xpcall(bad, handler)
print(ok == false)"#), "true");
}

#[test]
fn test_xpcall_error_message_length() {
    assert_eq!(run_lua_one(r#"local function bad() error("abcde") end
local function handler(err) return string.len(err) end
local ok, v = xpcall(bad, handler)
print(ok == false and v == 5)"#), "true");
}

#[test]
fn test_xpcall_assert_payload() {
    assert_eq!(run_lua_one(r#"local function bad() assert(false, "bad") end
local function handler(err) return string.find(err, "bad") ~= nil end
local ok, v = xpcall(bad, handler)
print(ok == false and v == true)"#), "true");
}

#[test]
fn test_xpcall_pcall_wrapper() {
    assert_eq!(run_lua_one(r#"local function good(x) return x * 2 end
local function handler(err) return 0 end
local ok, v = xpcall(good, handler, 7)
print(ok and v == 14)"#), "true");
}

#[test]
fn test_xpcall_handler_from_var() {
    assert_eq!(run_lua_one(r#"local function bad() error("x") end
local h = function(err) return err .. "!" end
local ok, v = xpcall(bad, h)
print(ok == false and string.find(v, "!") ~= nil)"#), "true");
}

#[test]
fn test_xpcall_handler_reads_level() {
    assert_eq!(run_lua_one(r#"local function bad() error("lvl") end
local function handler(err)
  return string.find(err, "lvl") ~= nil and 1 or 0
end
local ok, v = xpcall(bad, handler)
print(ok == false and v == 1)"#), "true");
}

#[test]
fn test_xpcall_error_with_nonstring_handler() {
    assert_eq!(run_lua_one(r#"local function bad() error("x") end
local function handler(err) return 99 end
local ok, v = xpcall(bad, handler)
print(ok == false and v == 99)"#), "true");
}

#[test]
fn test_xpcall_error_function_argument() {
    assert_eq!(run_lua_one(r#"local function bad() error(function() return 1 end) end
local function handler(err) return type(err) == "function" end
local ok, v = xpcall(bad, handler)
print(ok == false and v == true)"#), "true");
}

#[test]
fn test_xpcall_handles_nested_error() {
    assert_eq!(run_lua_one(r#"local function bad()
  local f = function() error("deep") end
  f()
end
local function handler(err)
  return err
end
local ok, v = xpcall(bad, handler)
print(ok == false and string.find(v, "deep") ~= nil)"#), "true");
}
