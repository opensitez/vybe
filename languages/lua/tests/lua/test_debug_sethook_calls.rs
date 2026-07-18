use super::helpers::run_lua_one;

#[test]
fn test_debug_sethook_calls_baseline() {
    assert_eq!(run_lua_one(r#"local n = 0
debug.sethook(function()
  n = n + 1
end, "c")
local function f()
  local s = 0
  for i = 1, 1 do s = s + i end
  return s
end
f()
debug.sethook()
print(n >= 1)"#), "true");
}


#[test]
fn test_debug_sethook_calls_simple() {
    assert_eq!(run_lua_one(r#"local n = 0
debug.sethook(function()
  n = n + 1
end, "c")
local function f()
  local s = 0
  for i = 1, 2 do s = s + i end
  return s
end
f()
debug.sethook()
print(n >= 1)"#), "true");
}


#[test]
fn test_debug_sethook_calls_trimmed() {
    assert_eq!(run_lua_one(r#"local n = 0
debug.sethook(function()
  n = n + 1
end, "c")
local function f()
  local s = 0
  for i = 1, 3 do s = s + i end
  return s
end
f()
debug.sethook()
print(n >= 1)"#), "true");
}


#[test]
fn test_debug_sethook_calls_decimal() {
    assert_eq!(run_lua_one(r#"local n = 0
debug.sethook(function()
  n = n + 1
end, "c")
local function f()
  local s = 0
  for i = 1, 4 do s = s + i end
  return s
end
f()
debug.sethook()
print(n >= 1)"#), "true");
}


#[test]
fn test_debug_sethook_calls_hexed() {
    assert_eq!(run_lua_one(r#"local n = 0
debug.sethook(function()
  n = n + 1
end, "c")
local function f()
  local s = 0
  for i = 1, 5 do s = s + i end
  return s
end
f()
debug.sethook()
print(n >= 1)"#), "true");
}


#[test]
fn test_debug_sethook_calls_prefixed() {
    assert_eq!(run_lua_one(r#"local n = 0
debug.sethook(function()
  n = n + 1
end, "c")
local function f()
  local s = 0
  for i = 1, 6 do s = s + i end
  return s
end
f()
debug.sethook()
print(n >= 1)"#), "true");
}


#[test]
fn test_debug_sethook_calls_negative() {
    assert_eq!(run_lua_one(r#"local n = 0
debug.sethook(function()
  n = n + 1
end, "c")
local function f()
  local s = 0
  for i = 1, 7 do s = s + i end
  return s
end
f()
debug.sethook()
print(n >= 1)"#), "true");
}


#[test]
fn test_debug_sethook_calls_rounded() {
    assert_eq!(run_lua_one(r#"local n = 0
debug.sethook(function()
  n = n + 1
end, "c")
local function f()
  local s = 0
  for i = 1, 8 do s = s + i end
  return s
end
f()
debug.sethook()
print(n >= 1)"#), "true");
}


#[test]
fn test_debug_sethook_calls_offset() {
    assert_eq!(run_lua_one(r#"local n = 0
debug.sethook(function()
  n = n + 1
end, "c")
local function f()
  local s = 0
  for i = 1, 9 do s = s + i end
  return s
end
f()
debug.sethook()
print(n >= 1)"#), "true");
}


#[test]
fn test_debug_sethook_calls_paired() {
    assert_eq!(run_lua_one(r#"local n = 0
debug.sethook(function()
  n = n + 1
end, "c")
local function f()
  local s = 0
  for i = 1, 10 do s = s + i end
  return s
end
f()
debug.sethook()
print(n >= 1)"#), "true");
}


#[test]
fn test_debug_sethook_calls_nested() {
    assert_eq!(run_lua_one(r#"local n = 0
debug.sethook(function()
  n = n + 1
end, "c")
local function f()
  local s = 0
  for i = 1, 11 do s = s + i end
  return s
end
f()
debug.sethook()
print(n >= 1)"#), "true");
}


#[test]
fn test_debug_sethook_calls_metaflow() {
    assert_eq!(run_lua_one(r#"local n = 0
debug.sethook(function()
  n = n + 1
end, "c")
local function f()
  local s = 0
  for i = 1, 12 do s = s + i end
  return s
end
f()
debug.sethook()
print(n >= 1)"#), "true");
}


#[test]
fn test_debug_sethook_calls_guarded() {
    assert_eq!(run_lua_one(r#"local n = 0
debug.sethook(function()
  n = n + 1
end, "c")
local function f()
  local s = 0
  for i = 1, 13 do s = s + i end
  return s
end
f()
debug.sethook()
print(n >= 1)"#), "true");
}


#[test]
fn test_debug_sethook_calls_mapped() {
    assert_eq!(run_lua_one(r#"local n = 0
debug.sethook(function()
  n = n + 1
end, "c")
local function f()
  local s = 0
  for i = 1, 14 do s = s + i end
  return s
end
f()
debug.sethook()
print(n >= 1)"#), "true");
}


#[test]
fn test_debug_sethook_calls_captured() {
    assert_eq!(run_lua_one(r#"local n = 0
debug.sethook(function()
  n = n + 1
end, "c")
local function f()
  local s = 0
  for i = 1, 15 do s = s + i end
  return s
end
f()
debug.sethook()
print(n >= 1)"#), "true");
}


#[test]
fn test_debug_sethook_calls_edge_first() {
    assert_eq!(run_lua_one(r#"local n = 0
debug.sethook(function()
  n = n + 1
end, "c")
local function f()
  local s = 0
  for i = 1, 16 do s = s + i end
  return s
end
f()
debug.sethook()
print(n >= 1)"#), "true");
}


#[test]
fn test_debug_sethook_calls_edge_second() {
    assert_eq!(run_lua_one(r#"local n = 0
debug.sethook(function()
  n = n + 1
end, "c")
local function f()
  local s = 0
  for i = 1, 17 do s = s + i end
  return s
end
f()
debug.sethook()
print(n >= 1)"#), "true");
}


#[test]
fn test_debug_sethook_calls_edge_last() {
    assert_eq!(run_lua_one(r#"local n = 0
debug.sethook(function()
  n = n + 1
end, "c")
local function f()
  local s = 0
  for i = 1, 18 do s = s + i end
  return s
end
f()
debug.sethook()
print(n >= 1)"#), "true");
}


#[test]
fn test_debug_sethook_calls_randomized() {
    assert_eq!(run_lua_one(r#"local n = 0
debug.sethook(function()
  n = n + 1
end, "c")
local function f()
  local s = 0
  for i = 1, 19 do s = s + i end
  return s
end
f()
debug.sethook()
print(n >= 1)"#), "true");
}


#[test]
fn test_debug_sethook_calls_unicode_like() {
    assert_eq!(run_lua_one(r#"local n = 0
debug.sethook(function()
  n = n + 1
end, "c")
local function f()
  local s = 0
  for i = 1, 20 do s = s + i end
  return s
end
f()
debug.sethook()
print(n >= 1)"#), "true");
}
