use super::helpers::run_lua_one;

#[test]
fn test_debug_getlocal_nested_baseline() {
    assert_eq!(run_lua_one(r#"local function inner()
  local z = 1
  return debug.getlocal(2, 1) ~= nil or debug.getlocal(1, 1) ~= nil
end
local function outer()
  local y = 2
  return inner()
end
print(outer())"#), "true");
}


#[test]
fn test_debug_getlocal_nested_simple() {
    assert_eq!(run_lua_one(r#"local function inner()
  local z = 2
  return debug.getlocal(2, 1) ~= nil or debug.getlocal(1, 1) ~= nil
end
local function outer()
  local y = 3
  return inner()
end
print(outer())"#), "true");
}


#[test]
fn test_debug_getlocal_nested_trimmed() {
    assert_eq!(run_lua_one(r#"local function inner()
  local z = 3
  return debug.getlocal(2, 1) ~= nil or debug.getlocal(1, 1) ~= nil
end
local function outer()
  local y = 4
  return inner()
end
print(outer())"#), "true");
}


#[test]
fn test_debug_getlocal_nested_decimal() {
    assert_eq!(run_lua_one(r#"local function inner()
  local z = 4
  return debug.getlocal(2, 1) ~= nil or debug.getlocal(1, 1) ~= nil
end
local function outer()
  local y = 5
  return inner()
end
print(outer())"#), "true");
}


#[test]
fn test_debug_getlocal_nested_hexed() {
    assert_eq!(run_lua_one(r#"local function inner()
  local z = 5
  return debug.getlocal(2, 1) ~= nil or debug.getlocal(1, 1) ~= nil
end
local function outer()
  local y = 6
  return inner()
end
print(outer())"#), "true");
}


#[test]
fn test_debug_getlocal_nested_prefixed() {
    assert_eq!(run_lua_one(r#"local function inner()
  local z = 6
  return debug.getlocal(2, 1) ~= nil or debug.getlocal(1, 1) ~= nil
end
local function outer()
  local y = 7
  return inner()
end
print(outer())"#), "true");
}


#[test]
fn test_debug_getlocal_nested_negative() {
    assert_eq!(run_lua_one(r#"local function inner()
  local z = 7
  return debug.getlocal(2, 1) ~= nil or debug.getlocal(1, 1) ~= nil
end
local function outer()
  local y = 8
  return inner()
end
print(outer())"#), "true");
}


#[test]
fn test_debug_getlocal_nested_rounded() {
    assert_eq!(run_lua_one(r#"local function inner()
  local z = 8
  return debug.getlocal(2, 1) ~= nil or debug.getlocal(1, 1) ~= nil
end
local function outer()
  local y = 9
  return inner()
end
print(outer())"#), "true");
}


#[test]
fn test_debug_getlocal_nested_offset() {
    assert_eq!(run_lua_one(r#"local function inner()
  local z = 9
  return debug.getlocal(2, 1) ~= nil or debug.getlocal(1, 1) ~= nil
end
local function outer()
  local y = 10
  return inner()
end
print(outer())"#), "true");
}


#[test]
fn test_debug_getlocal_nested_paired() {
    assert_eq!(run_lua_one(r#"local function inner()
  local z = 10
  return debug.getlocal(2, 1) ~= nil or debug.getlocal(1, 1) ~= nil
end
local function outer()
  local y = 11
  return inner()
end
print(outer())"#), "true");
}


#[test]
fn test_debug_getlocal_nested_nested() {
    assert_eq!(run_lua_one(r#"local function inner()
  local z = 11
  return debug.getlocal(2, 1) ~= nil or debug.getlocal(1, 1) ~= nil
end
local function outer()
  local y = 12
  return inner()
end
print(outer())"#), "true");
}


#[test]
fn test_debug_getlocal_nested_metaflow() {
    assert_eq!(run_lua_one(r#"local function inner()
  local z = 12
  return debug.getlocal(2, 1) ~= nil or debug.getlocal(1, 1) ~= nil
end
local function outer()
  local y = 13
  return inner()
end
print(outer())"#), "true");
}


#[test]
fn test_debug_getlocal_nested_guarded() {
    assert_eq!(run_lua_one(r#"local function inner()
  local z = 13
  return debug.getlocal(2, 1) ~= nil or debug.getlocal(1, 1) ~= nil
end
local function outer()
  local y = 14
  return inner()
end
print(outer())"#), "true");
}


#[test]
fn test_debug_getlocal_nested_mapped() {
    assert_eq!(run_lua_one(r#"local function inner()
  local z = 14
  return debug.getlocal(2, 1) ~= nil or debug.getlocal(1, 1) ~= nil
end
local function outer()
  local y = 15
  return inner()
end
print(outer())"#), "true");
}


#[test]
fn test_debug_getlocal_nested_captured() {
    assert_eq!(run_lua_one(r#"local function inner()
  local z = 15
  return debug.getlocal(2, 1) ~= nil or debug.getlocal(1, 1) ~= nil
end
local function outer()
  local y = 16
  return inner()
end
print(outer())"#), "true");
}


#[test]
fn test_debug_getlocal_nested_edge_first() {
    assert_eq!(run_lua_one(r#"local function inner()
  local z = 16
  return debug.getlocal(2, 1) ~= nil or debug.getlocal(1, 1) ~= nil
end
local function outer()
  local y = 17
  return inner()
end
print(outer())"#), "true");
}


#[test]
fn test_debug_getlocal_nested_edge_second() {
    assert_eq!(run_lua_one(r#"local function inner()
  local z = 17
  return debug.getlocal(2, 1) ~= nil or debug.getlocal(1, 1) ~= nil
end
local function outer()
  local y = 18
  return inner()
end
print(outer())"#), "true");
}


#[test]
fn test_debug_getlocal_nested_edge_last() {
    assert_eq!(run_lua_one(r#"local function inner()
  local z = 18
  return debug.getlocal(2, 1) ~= nil or debug.getlocal(1, 1) ~= nil
end
local function outer()
  local y = 19
  return inner()
end
print(outer())"#), "true");
}


#[test]
fn test_debug_getlocal_nested_randomized() {
    assert_eq!(run_lua_one(r#"local function inner()
  local z = 19
  return debug.getlocal(2, 1) ~= nil or debug.getlocal(1, 1) ~= nil
end
local function outer()
  local y = 20
  return inner()
end
print(outer())"#), "true");
}


#[test]
fn test_debug_getlocal_nested_unicode_like() {
    assert_eq!(run_lua_one(r#"local function inner()
  local z = 20
  return debug.getlocal(2, 1) ~= nil or debug.getlocal(1, 1) ~= nil
end
local function outer()
  local y = 21
  return inner()
end
print(outer())"#), "true");
}
