use super::helpers::run_lua_one;

#[test]
fn test_pcall_single_return() {
    assert_eq!(
        run_lua_one(
            r#"local ok, v = pcall(function() return 5 end)
print(ok and v)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_return_three_values() {
    assert_eq!(
        run_lua_one(
            r#"local ok, a = pcall(function() return 1,2,3 end)
print(ok and a)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_return_nil() {
    assert_eq!(
        run_lua_one(
            r#"local ok, v = pcall(function() return nil end)
print(ok and (v == nil))"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_return_false() {
    assert_eq!(
        run_lua_one(
            r#"local ok, v = pcall(function() return false end)
print(ok and (v == false))"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_return_true() {
    assert_eq!(
        run_lua_one(
            r#"local ok, v = pcall(function() return true end)
print(ok and v)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_return_empty_string() {
    assert_eq!(
        run_lua_one(
            r#"local ok, v = pcall(function() return "" end)
print(ok and v == "")"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_return_table() {
    assert_eq!(
        run_lua_one(
            r#"local ok, v = pcall(function() return {} end)
print(ok and type(v) == "table")"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_return_function() {
    assert_eq!(
        run_lua_one(
            r#"local ok, v = pcall(function() return function() end end)
print(ok and type(v) == "function")"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_return_coroutine() {
    assert_eq!(
        run_lua_one(
            r#"local t = coroutine.create(function() end)
local ok, v = pcall(function() return t end)
print(ok and type(v) == "thread")"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_return_nested_function() {
    assert_eq!(
        run_lua_one(
            r#"local ok, v = pcall(function()
  local function f() return 8 end
  return f()
end)
print(ok and v == 8)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_return_math_abs() {
    assert_eq!(
        run_lua_one(
            r#"local ok, v = pcall(function() return math.abs(-2) end)
print(ok and v == 2)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_return_string_rep() {
    assert_eq!(
        run_lua_one(
            r#"local ok, v = pcall(function() return string.rep("a", 3) end)
print(ok and v == "aaa")"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_return_table_len() {
    assert_eq!(
        run_lua_one(
            r#"local ok, v = pcall(function() local t={1,2,3}; return #t end)
print(ok and v == 3)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_return_typed_operation() {
    assert_eq!(
        run_lua_one(
            r#"local ok, v = pcall(function(x,y) return x + y end, 5, 6)
print(ok and v == 11)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_return_concat() {
    assert_eq!(
        run_lua_one(
            r#"local ok, v = pcall(function() return "x" .. "y" end)
print(ok and v == "xy")"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_return_compare() {
    assert_eq!(
        run_lua_one(
            r#"local ok, v = pcall(function() return 10 > 3 end)
print(ok and v == true)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_return_index() {
    assert_eq!(
        run_lua_one(
            r#"local ok, v = pcall(function() local t={a=1}; return t.a end)
print(ok and v == 1)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_return_conditional() {
    assert_eq!(
        run_lua_one(
            r#"local ok, v = pcall(function() local x=2; if x>1 then return x end return 0 end)
print(ok and v == 2)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_return_variadic() {
    assert_eq!(
        run_lua_one(
            r#"local function f(...)
  local a,b = ...
  return a * b
end
local ok, v = pcall(f, 3, 4)
print(ok and v == 12)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_return_if_true() {
    assert_eq!(
        run_lua_one(
            r#"local ok, v = pcall(function() if true then return 1 else return 2 end end)
print(ok and v == 1)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_return_if_false() {
    assert_eq!(
        run_lua_one(
            r#"local ok, v = pcall(function() if false then return 1 else return 2 end end)
print(ok and v == 2)"#
        ),
        "true"
    );
}
