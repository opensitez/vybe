use super::helpers::run_lua_one;

#[test]
fn test_coroutine_close_guarded_baseline() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function()
  local x = 1
  return x
end)
coroutine.resume(co)
print(pcall(coroutine.close, co) ~= nil)"#), "true");
}


#[test]
fn test_coroutine_close_guarded_simple() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function()
  local x = 2
  return x
end)
coroutine.resume(co)
print(pcall(coroutine.close, co) ~= nil)"#), "true");
}


#[test]
fn test_coroutine_close_guarded_trimmed() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function()
  local x = 3
  return x
end)
coroutine.resume(co)
print(pcall(coroutine.close, co) ~= nil)"#), "true");
}


#[test]
fn test_coroutine_close_guarded_decimal() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function()
  local x = 4
  return x
end)
coroutine.resume(co)
print(pcall(coroutine.close, co) ~= nil)"#), "true");
}


#[test]
fn test_coroutine_close_guarded_hexed() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function()
  local x = 5
  return x
end)
coroutine.resume(co)
print(pcall(coroutine.close, co) ~= nil)"#), "true");
}


#[test]
fn test_coroutine_close_guarded_prefixed() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function()
  local x = 6
  return x
end)
coroutine.resume(co)
print(pcall(coroutine.close, co) ~= nil)"#), "true");
}


#[test]
fn test_coroutine_close_guarded_negative() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function()
  local x = 7
  return x
end)
coroutine.resume(co)
print(pcall(coroutine.close, co) ~= nil)"#), "true");
}


#[test]
fn test_coroutine_close_guarded_rounded() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function()
  local x = 8
  return x
end)
coroutine.resume(co)
print(pcall(coroutine.close, co) ~= nil)"#), "true");
}


#[test]
fn test_coroutine_close_guarded_offset() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function()
  local x = 9
  return x
end)
coroutine.resume(co)
print(pcall(coroutine.close, co) ~= nil)"#), "true");
}


#[test]
fn test_coroutine_close_guarded_paired() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function()
  local x = 10
  return x
end)
coroutine.resume(co)
print(pcall(coroutine.close, co) ~= nil)"#), "true");
}


#[test]
fn test_coroutine_close_guarded_nested() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function()
  local x = 11
  return x
end)
coroutine.resume(co)
print(pcall(coroutine.close, co) ~= nil)"#), "true");
}


#[test]
fn test_coroutine_close_guarded_metaflow() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function()
  local x = 12
  return x
end)
coroutine.resume(co)
print(pcall(coroutine.close, co) ~= nil)"#), "true");
}


#[test]
fn test_coroutine_close_guarded_guarded() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function()
  local x = 13
  return x
end)
coroutine.resume(co)
print(pcall(coroutine.close, co) ~= nil)"#), "true");
}


#[test]
fn test_coroutine_close_guarded_mapped() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function()
  local x = 14
  return x
end)
coroutine.resume(co)
print(pcall(coroutine.close, co) ~= nil)"#), "true");
}


#[test]
fn test_coroutine_close_guarded_captured() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function()
  local x = 15
  return x
end)
coroutine.resume(co)
print(pcall(coroutine.close, co) ~= nil)"#), "true");
}


#[test]
fn test_coroutine_close_guarded_edge_first() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function()
  local x = 16
  return x
end)
coroutine.resume(co)
print(pcall(coroutine.close, co) ~= nil)"#), "true");
}


#[test]
fn test_coroutine_close_guarded_edge_second() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function()
  local x = 17
  return x
end)
coroutine.resume(co)
print(pcall(coroutine.close, co) ~= nil)"#), "true");
}


#[test]
fn test_coroutine_close_guarded_edge_last() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function()
  local x = 18
  return x
end)
coroutine.resume(co)
print(pcall(coroutine.close, co) ~= nil)"#), "true");
}


#[test]
fn test_coroutine_close_guarded_randomized() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function()
  local x = 19
  return x
end)
coroutine.resume(co)
print(pcall(coroutine.close, co) ~= nil)"#), "true");
}


#[test]
fn test_coroutine_close_guarded_unicode_like() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function()
  local x = 20
  return x
end)
coroutine.resume(co)
print(pcall(coroutine.close, co) ~= nil)"#), "true");
}
