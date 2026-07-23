use super::helpers::run_lua_one;

#[test]
fn test_pcall_runtime_error() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() error("x") end)
print(ok == false)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_assert_error() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() assert(false, "bad") end)
print(ok == false)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_division_error() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() return 1/0 end)
print(ok)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_index_error() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() local a = nil; return a.x end)
print(ok == false)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_call_non_function() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() local n = 1; n() end)
print(ok == false)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_bad_argument_count() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function(a) return a + 1 end)
print(ok == false)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_bad_math() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() return "a" + 1 end)
print(ok == false)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_nested_error() {
    assert_eq!(
        run_lua_one(
            r#"local function inner() error("inner") end
local ok, err = pcall(function() inner() end)
print(ok == false and string.find(err, "inner") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_recursive_error() {
    assert_eq!(
        run_lua_one(
            r#"local function go(n) if n == 0 then error("end") else return go(n-1) end end
local ok, err = pcall(function() go(2) end)
print(ok == false and string.find(err, "end") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_error_to_string() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() error("z") end)
print(type(err) == "string")"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_error_payload_table() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() error({k=1}) end)
print(type(err) == "table")"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_error_payload_number() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() error(12) end)
print(type(err) == "number")"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_error_payload_bool() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() error(true) end)
print(type(err) == "boolean")"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_error_payload_function() {
    assert_eq!(
        run_lua_one(
            r#"local f = function() end
local ok, err = pcall(function() error(f) end)
print(type(err) == "function")"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_error_level_message() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() error("l", 2) end)
print(string.find(err, "l") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_error_after_return() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() if true then return 1 end error("x") end)
print(ok and err == 1)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_error_in_function_call() {
    assert_eq!(
        run_lua_one(
            r#"local ok, fn = pcall(function() return function() return 1 end end)
print(ok and type(fn) == "function")
"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_error_when_passing_arg_type() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function(x) return x.bad end, 10)
print(ok == false)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_error_if_false_guard() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() if false then error("x") else return false end end)
print(ok and err == false)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_error_string_match() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() if string.find(nil, "a") then end end)
print(ok == false)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_error_table_index() {
    assert_eq!(
        run_lua_one(
            r#"local t = {a = 1}
local ok, err = pcall(function() return t[1][1] end)
print(ok == false)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_error_concat_nil() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() return nil .. "x" end)
print(ok == false)"#
        ),
        "true"
    );
}

#[test]
fn test_pcall_error_boolean_concat() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() return true .. "x" end)
print(ok == false)"#
        ),
        "true"
    );
}
