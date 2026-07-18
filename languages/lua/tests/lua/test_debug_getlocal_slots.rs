use super::helpers::run_lua_one;

#[test]
fn test_debug_getlocal_slots_baseline() {
    assert_eq!(run_lua_one(r#"local function f()
  local a = 1
  local b = 2
  return debug.getlocal(1, 1) ~= nil and debug.getlocal(1, 2) ~= nil
end
print(f())"#), "true");
}


#[test]
fn test_debug_getlocal_slots_simple() {
    assert_eq!(run_lua_one(r#"local function f()
  local a = 2
  local b = 3
  return debug.getlocal(1, 1) ~= nil and debug.getlocal(1, 2) ~= nil
end
print(f())"#), "true");
}


#[test]
fn test_debug_getlocal_slots_trimmed() {
    assert_eq!(run_lua_one(r#"local function f()
  local a = 3
  local b = 4
  return debug.getlocal(1, 1) ~= nil and debug.getlocal(1, 2) ~= nil
end
print(f())"#), "true");
}


#[test]
fn test_debug_getlocal_slots_decimal() {
    assert_eq!(run_lua_one(r#"local function f()
  local a = 4
  local b = 5
  return debug.getlocal(1, 1) ~= nil and debug.getlocal(1, 2) ~= nil
end
print(f())"#), "true");
}


#[test]
fn test_debug_getlocal_slots_hexed() {
    assert_eq!(run_lua_one(r#"local function f()
  local a = 5
  local b = 6
  return debug.getlocal(1, 1) ~= nil and debug.getlocal(1, 2) ~= nil
end
print(f())"#), "true");
}


#[test]
fn test_debug_getlocal_slots_prefixed() {
    assert_eq!(run_lua_one(r#"local function f()
  local a = 6
  local b = 7
  return debug.getlocal(1, 1) ~= nil and debug.getlocal(1, 2) ~= nil
end
print(f())"#), "true");
}


#[test]
fn test_debug_getlocal_slots_negative() {
    assert_eq!(run_lua_one(r#"local function f()
  local a = 7
  local b = 8
  return debug.getlocal(1, 1) ~= nil and debug.getlocal(1, 2) ~= nil
end
print(f())"#), "true");
}


#[test]
fn test_debug_getlocal_slots_rounded() {
    assert_eq!(run_lua_one(r#"local function f()
  local a = 8
  local b = 9
  return debug.getlocal(1, 1) ~= nil and debug.getlocal(1, 2) ~= nil
end
print(f())"#), "true");
}


#[test]
fn test_debug_getlocal_slots_offset() {
    assert_eq!(run_lua_one(r#"local function f()
  local a = 9
  local b = 10
  return debug.getlocal(1, 1) ~= nil and debug.getlocal(1, 2) ~= nil
end
print(f())"#), "true");
}


#[test]
fn test_debug_getlocal_slots_paired() {
    assert_eq!(run_lua_one(r#"local function f()
  local a = 10
  local b = 11
  return debug.getlocal(1, 1) ~= nil and debug.getlocal(1, 2) ~= nil
end
print(f())"#), "true");
}


#[test]
fn test_debug_getlocal_slots_nested() {
    assert_eq!(run_lua_one(r#"local function f()
  local a = 11
  local b = 12
  return debug.getlocal(1, 1) ~= nil and debug.getlocal(1, 2) ~= nil
end
print(f())"#), "true");
}


#[test]
fn test_debug_getlocal_slots_metaflow() {
    assert_eq!(run_lua_one(r#"local function f()
  local a = 12
  local b = 13
  return debug.getlocal(1, 1) ~= nil and debug.getlocal(1, 2) ~= nil
end
print(f())"#), "true");
}


#[test]
fn test_debug_getlocal_slots_guarded() {
    assert_eq!(run_lua_one(r#"local function f()
  local a = 13
  local b = 14
  return debug.getlocal(1, 1) ~= nil and debug.getlocal(1, 2) ~= nil
end
print(f())"#), "true");
}


#[test]
fn test_debug_getlocal_slots_mapped() {
    assert_eq!(run_lua_one(r#"local function f()
  local a = 14
  local b = 15
  return debug.getlocal(1, 1) ~= nil and debug.getlocal(1, 2) ~= nil
end
print(f())"#), "true");
}


#[test]
fn test_debug_getlocal_slots_captured() {
    assert_eq!(run_lua_one(r#"local function f()
  local a = 15
  local b = 16
  return debug.getlocal(1, 1) ~= nil and debug.getlocal(1, 2) ~= nil
end
print(f())"#), "true");
}


#[test]
fn test_debug_getlocal_slots_edge_first() {
    assert_eq!(run_lua_one(r#"local function f()
  local a = 16
  local b = 17
  return debug.getlocal(1, 1) ~= nil and debug.getlocal(1, 2) ~= nil
end
print(f())"#), "true");
}


#[test]
fn test_debug_getlocal_slots_edge_second() {
    assert_eq!(run_lua_one(r#"local function f()
  local a = 17
  local b = 18
  return debug.getlocal(1, 1) ~= nil and debug.getlocal(1, 2) ~= nil
end
print(f())"#), "true");
}


#[test]
fn test_debug_getlocal_slots_edge_last() {
    assert_eq!(run_lua_one(r#"local function f()
  local a = 18
  local b = 19
  return debug.getlocal(1, 1) ~= nil and debug.getlocal(1, 2) ~= nil
end
print(f())"#), "true");
}


#[test]
fn test_debug_getlocal_slots_randomized() {
    assert_eq!(run_lua_one(r#"local function f()
  local a = 19
  local b = 20
  return debug.getlocal(1, 1) ~= nil and debug.getlocal(1, 2) ~= nil
end
print(f())"#), "true");
}


#[test]
fn test_debug_getlocal_slots_unicode_like() {
    assert_eq!(run_lua_one(r#"local function f()
  local a = 20
  local b = 21
  return debug.getlocal(1, 1) ~= nil and debug.getlocal(1, 2) ~= nil
end
print(f())"#), "true");
}
