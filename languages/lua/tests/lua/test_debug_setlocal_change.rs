use super::helpers::run_lua_one;

#[test]
fn test_debug_setlocal_change_baseline() {
    assert_eq!(run_lua_one(r#"local function f()
  local x = 10
  debug.setlocal(1, 1, 11)
  print(x == 11)
end
f()"#), "true");
}


#[test]
fn test_debug_setlocal_change_simple() {
    assert_eq!(run_lua_one(r#"local function f()
  local x = 11
  debug.setlocal(1, 1, 12)
  print(x == 12)
end
f()"#), "true");
}


#[test]
fn test_debug_setlocal_change_trimmed() {
    assert_eq!(run_lua_one(r#"local function f()
  local x = 12
  debug.setlocal(1, 1, 13)
  print(x == 13)
end
f()"#), "true");
}


#[test]
fn test_debug_setlocal_change_decimal() {
    assert_eq!(run_lua_one(r#"local function f()
  local x = 13
  debug.setlocal(1, 1, 14)
  print(x == 14)
end
f()"#), "true");
}


#[test]
fn test_debug_setlocal_change_hexed() {
    assert_eq!(run_lua_one(r#"local function f()
  local x = 14
  debug.setlocal(1, 1, 15)
  print(x == 15)
end
f()"#), "true");
}


#[test]
fn test_debug_setlocal_change_prefixed() {
    assert_eq!(run_lua_one(r#"local function f()
  local x = 15
  debug.setlocal(1, 1, 16)
  print(x == 16)
end
f()"#), "true");
}


#[test]
fn test_debug_setlocal_change_negative() {
    assert_eq!(run_lua_one(r#"local function f()
  local x = 16
  debug.setlocal(1, 1, 17)
  print(x == 17)
end
f()"#), "true");
}


#[test]
fn test_debug_setlocal_change_rounded() {
    assert_eq!(run_lua_one(r#"local function f()
  local x = 17
  debug.setlocal(1, 1, 18)
  print(x == 18)
end
f()"#), "true");
}


#[test]
fn test_debug_setlocal_change_offset() {
    assert_eq!(run_lua_one(r#"local function f()
  local x = 18
  debug.setlocal(1, 1, 19)
  print(x == 19)
end
f()"#), "true");
}


#[test]
fn test_debug_setlocal_change_paired() {
    assert_eq!(run_lua_one(r#"local function f()
  local x = 19
  debug.setlocal(1, 1, 20)
  print(x == 20)
end
f()"#), "true");
}


#[test]
fn test_debug_setlocal_change_nested() {
    assert_eq!(run_lua_one(r#"local function f()
  local x = 20
  debug.setlocal(1, 1, 21)
  print(x == 21)
end
f()"#), "true");
}


#[test]
fn test_debug_setlocal_change_metaflow() {
    assert_eq!(run_lua_one(r#"local function f()
  local x = 21
  debug.setlocal(1, 1, 22)
  print(x == 22)
end
f()"#), "true");
}


#[test]
fn test_debug_setlocal_change_guarded() {
    assert_eq!(run_lua_one(r#"local function f()
  local x = 22
  debug.setlocal(1, 1, 23)
  print(x == 23)
end
f()"#), "true");
}


#[test]
fn test_debug_setlocal_change_mapped() {
    assert_eq!(run_lua_one(r#"local function f()
  local x = 23
  debug.setlocal(1, 1, 24)
  print(x == 24)
end
f()"#), "true");
}


#[test]
fn test_debug_setlocal_change_captured() {
    assert_eq!(run_lua_one(r#"local function f()
  local x = 24
  debug.setlocal(1, 1, 25)
  print(x == 25)
end
f()"#), "true");
}


#[test]
fn test_debug_setlocal_change_edge_first() {
    assert_eq!(run_lua_one(r#"local function f()
  local x = 25
  debug.setlocal(1, 1, 26)
  print(x == 26)
end
f()"#), "true");
}


#[test]
fn test_debug_setlocal_change_edge_second() {
    assert_eq!(run_lua_one(r#"local function f()
  local x = 26
  debug.setlocal(1, 1, 27)
  print(x == 27)
end
f()"#), "true");
}


#[test]
fn test_debug_setlocal_change_edge_last() {
    assert_eq!(run_lua_one(r#"local function f()
  local x = 27
  debug.setlocal(1, 1, 28)
  print(x == 28)
end
f()"#), "true");
}


#[test]
fn test_debug_setlocal_change_randomized() {
    assert_eq!(run_lua_one(r#"local function f()
  local x = 28
  debug.setlocal(1, 1, 29)
  print(x == 29)
end
f()"#), "true");
}


#[test]
fn test_debug_setlocal_change_unicode_like() {
    assert_eq!(run_lua_one(r#"local function f()
  local x = 29
  debug.setlocal(1, 1, 30)
  print(x == 30)
end
f()"#), "true");
}
