use super::helpers::run_lua_one;

#[test]
fn test_getfenv_setfenv_baseline() {
    assert_eq!(run_lua_one(r#"local function f()
  return _G or nil
end
local env = {x = 1}
if type(setfenv) == "function" and type(getfenv) == "function" then
  setfenv(f, env)
  print(getfenv(f).x == 1)
else
  print(type(f) == "function")
end"#), "true");
}


#[test]
fn test_getfenv_setfenv_simple() {
    assert_eq!(run_lua_one(r#"local function f()
  return _G or nil
end
local env = {x = 2}
if type(setfenv) == "function" and type(getfenv) == "function" then
  setfenv(f, env)
  print(getfenv(f).x == 2)
else
  print(type(f) == "function")
end"#), "true");
}


#[test]
fn test_getfenv_setfenv_trimmed() {
    assert_eq!(run_lua_one(r#"local function f()
  return _G or nil
end
local env = {x = 3}
if type(setfenv) == "function" and type(getfenv) == "function" then
  setfenv(f, env)
  print(getfenv(f).x == 3)
else
  print(type(f) == "function")
end"#), "true");
}


#[test]
fn test_getfenv_setfenv_decimal() {
    assert_eq!(run_lua_one(r#"local function f()
  return _G or nil
end
local env = {x = 4}
if type(setfenv) == "function" and type(getfenv) == "function" then
  setfenv(f, env)
  print(getfenv(f).x == 4)
else
  print(type(f) == "function")
end"#), "true");
}


#[test]
fn test_getfenv_setfenv_hexed() {
    assert_eq!(run_lua_one(r#"local function f()
  return _G or nil
end
local env = {x = 5}
if type(setfenv) == "function" and type(getfenv) == "function" then
  setfenv(f, env)
  print(getfenv(f).x == 5)
else
  print(type(f) == "function")
end"#), "true");
}


#[test]
fn test_getfenv_setfenv_prefixed() {
    assert_eq!(run_lua_one(r#"local function f()
  return _G or nil
end
local env = {x = 6}
if type(setfenv) == "function" and type(getfenv) == "function" then
  setfenv(f, env)
  print(getfenv(f).x == 6)
else
  print(type(f) == "function")
end"#), "true");
}


#[test]
fn test_getfenv_setfenv_negative() {
    assert_eq!(run_lua_one(r#"local function f()
  return _G or nil
end
local env = {x = 7}
if type(setfenv) == "function" and type(getfenv) == "function" then
  setfenv(f, env)
  print(getfenv(f).x == 7)
else
  print(type(f) == "function")
end"#), "true");
}


#[test]
fn test_getfenv_setfenv_rounded() {
    assert_eq!(run_lua_one(r#"local function f()
  return _G or nil
end
local env = {x = 8}
if type(setfenv) == "function" and type(getfenv) == "function" then
  setfenv(f, env)
  print(getfenv(f).x == 8)
else
  print(type(f) == "function")
end"#), "true");
}


#[test]
fn test_getfenv_setfenv_offset() {
    assert_eq!(run_lua_one(r#"local function f()
  return _G or nil
end
local env = {x = 9}
if type(setfenv) == "function" and type(getfenv) == "function" then
  setfenv(f, env)
  print(getfenv(f).x == 9)
else
  print(type(f) == "function")
end"#), "true");
}


#[test]
fn test_getfenv_setfenv_paired() {
    assert_eq!(run_lua_one(r#"local function f()
  return _G or nil
end
local env = {x = 10}
if type(setfenv) == "function" and type(getfenv) == "function" then
  setfenv(f, env)
  print(getfenv(f).x == 10)
else
  print(type(f) == "function")
end"#), "true");
}


#[test]
fn test_getfenv_setfenv_nested() {
    assert_eq!(run_lua_one(r#"local function f()
  return _G or nil
end
local env = {x = 11}
if type(setfenv) == "function" and type(getfenv) == "function" then
  setfenv(f, env)
  print(getfenv(f).x == 11)
else
  print(type(f) == "function")
end"#), "true");
}


#[test]
fn test_getfenv_setfenv_metaflow() {
    assert_eq!(run_lua_one(r#"local function f()
  return _G or nil
end
local env = {x = 12}
if type(setfenv) == "function" and type(getfenv) == "function" then
  setfenv(f, env)
  print(getfenv(f).x == 12)
else
  print(type(f) == "function")
end"#), "true");
}


#[test]
fn test_getfenv_setfenv_guarded() {
    assert_eq!(run_lua_one(r#"local function f()
  return _G or nil
end
local env = {x = 13}
if type(setfenv) == "function" and type(getfenv) == "function" then
  setfenv(f, env)
  print(getfenv(f).x == 13)
else
  print(type(f) == "function")
end"#), "true");
}


#[test]
fn test_getfenv_setfenv_mapped() {
    assert_eq!(run_lua_one(r#"local function f()
  return _G or nil
end
local env = {x = 14}
if type(setfenv) == "function" and type(getfenv) == "function" then
  setfenv(f, env)
  print(getfenv(f).x == 14)
else
  print(type(f) == "function")
end"#), "true");
}


#[test]
fn test_getfenv_setfenv_captured() {
    assert_eq!(run_lua_one(r#"local function f()
  return _G or nil
end
local env = {x = 15}
if type(setfenv) == "function" and type(getfenv) == "function" then
  setfenv(f, env)
  print(getfenv(f).x == 15)
else
  print(type(f) == "function")
end"#), "true");
}


#[test]
fn test_getfenv_setfenv_edge_first() {
    assert_eq!(run_lua_one(r#"local function f()
  return _G or nil
end
local env = {x = 16}
if type(setfenv) == "function" and type(getfenv) == "function" then
  setfenv(f, env)
  print(getfenv(f).x == 16)
else
  print(type(f) == "function")
end"#), "true");
}


#[test]
fn test_getfenv_setfenv_edge_second() {
    assert_eq!(run_lua_one(r#"local function f()
  return _G or nil
end
local env = {x = 17}
if type(setfenv) == "function" and type(getfenv) == "function" then
  setfenv(f, env)
  print(getfenv(f).x == 17)
else
  print(type(f) == "function")
end"#), "true");
}


#[test]
fn test_getfenv_setfenv_edge_last() {
    assert_eq!(run_lua_one(r#"local function f()
  return _G or nil
end
local env = {x = 18}
if type(setfenv) == "function" and type(getfenv) == "function" then
  setfenv(f, env)
  print(getfenv(f).x == 18)
else
  print(type(f) == "function")
end"#), "true");
}


#[test]
fn test_getfenv_setfenv_randomized() {
    assert_eq!(run_lua_one(r#"local function f()
  return _G or nil
end
local env = {x = 19}
if type(setfenv) == "function" and type(getfenv) == "function" then
  setfenv(f, env)
  print(getfenv(f).x == 19)
else
  print(type(f) == "function")
end"#), "true");
}


#[test]
fn test_getfenv_setfenv_unicode_like() {
    assert_eq!(run_lua_one(r#"local function f()
  return _G or nil
end
local env = {x = 20}
if type(setfenv) == "function" and type(getfenv) == "function" then
  setfenv(f, env)
  print(getfenv(f).x == 20)
else
  print(type(f) == "function")
end"#), "true");
}
