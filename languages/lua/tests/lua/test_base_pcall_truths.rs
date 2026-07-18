use super::helpers::run_lua_one;

#[test]
fn test_pcall_true_returns_true() {
    assert_eq!(run_lua_one(r#"local ok, v = pcall(function() return true end)
print(ok and v)"#), "true");
}

#[test]
fn test_pcall_number_returns_number() {
    assert_eq!(run_lua_one(r#"local ok, v = pcall(function() return 10 end)
print(ok and v == 10)"#), "true");
}

#[test]
fn test_pcall_string_returns_string() {
    assert_eq!(run_lua_one(r#"local ok, v = pcall(function() return \"x\" end)
print(ok and v == \"x\")"#), "true");
}

#[test]
fn test_pcall_multiple_return_second() {
    assert_eq!(run_lua_one(r#"local ok, a = pcall(function() return 1, 2 end)
print(ok and a)"#), "1");
}

#[test]
fn test_pcall_bool_arg() {
    assert_eq!(run_lua_one(r#"local ok, v = pcall(function(x) return x end, true)
print(ok and v == true)"#), "true");
}

#[test]
fn test_pcall_string_arg() {
    assert_eq!(run_lua_one(r#"local ok, v = pcall(function(x) return x .. \"y\" end, \"x\")
print(ok and v == \"xy\")"#), "true");
}

#[test]
fn test_pcall_table_arg() {
    assert_eq!(run_lua_one(r#"local t = {value = 1}
local ok, v = pcall(function(x) return x.value end, t)
print(ok and v == 1)"#), "true");
}

#[test]
fn test_pcall_nested_call() {
    assert_eq!(run_lua_one(r#"local function inner() return 3 end
local ok, v = pcall(function() return inner() end)
print(ok and v == 3)"#), "true");
}

#[test]
fn test_pcall_with_local_function() {
    assert_eq!(run_lua_one(r#"local function make(v) return v end
local ok, v = pcall(make, 9)
print(ok and v == 9)"#), "true");
}

#[test]
fn test_pcall_math() {
    assert_eq!(run_lua_one(r#"local ok, v = pcall(function() return 8 + 2 end)
print(ok and v == 10)"#), "true");
}

#[test]
fn test_pcall_abs() {
    assert_eq!(run_lua_one(r#"local ok, v = pcall(function() return math.abs(-9) end)
print(ok and v == 9)"#), "true");
}

#[test]
fn test_pcall_nil_return_false() {
    assert_eq!(run_lua_one(r#"local ok, v = pcall(function() return nil end)
print(ok and v == nil)"#), "true");
}

#[test]
fn test_pcall_function_error_propagates_false() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() error(\"boom\") end)
print(ok == false)"#), "true");
}

#[test]
fn test_pcall_error_type() {
    assert_eq!(run_lua_one(r#"local ok, err = pcall(function() error(42) end)
print(ok == false and type(err) == \"number\")"#), "true");
}

#[test]
fn test_pcall_with_two_args() {
    assert_eq!(run_lua_one(r#"local ok, v = pcall(function(a,b) return a + b end, 2, 3)
print(ok and v == 5)"#), "true");
}

#[test]
fn test_pcall_nested_pcall() {
    assert_eq!(run_lua_one(r#"local ok, inner = pcall(function()
  local a = pcall(function() return 2 end)
  return a
end)
print(ok)"#), "true");
}

#[test]
fn test_pcall_error_from_nested() {
    assert_eq!(run_lua_one(r#"local ok, v = pcall(function() pcall(function() error(\"x\") end) end)
print(ok and v == false)"#), "true");
}

#[test]
fn test_pcall_variadic_wrapper() {
    assert_eq!(run_lua_one(r#"local function f(...)
  local a,b,c = ...
  return a+b+c
end
local ok, v = pcall(f, 1,2,3)
print(ok and v == 6)"#), "true");
}

#[test]
fn test_pcall_boolean_not() {
    assert_eq!(run_lua_one(r#"local ok, v = pcall(function() return not false end)
print(ok and v == true)"#), "true");
}

#[test]
fn test_pcall_comparison() {
    assert_eq!(run_lua_one(r#"local ok, v = pcall(function() return 4 > 3 end)
print(ok and v == true)"#), "true");
}

#[test]
fn test_pcall_type_query() {
    assert_eq!(run_lua_one(r#"local ok, v = pcall(function() return type(\"x\") end)
print(ok and v == \"string\")"#), "true");
}

#[test]
fn test_pcall_string_len() {
    assert_eq!(run_lua_one(r#"local ok, v = pcall(function() return #\"abc\" end)
print(ok and v == 3)"#), "true");
}
