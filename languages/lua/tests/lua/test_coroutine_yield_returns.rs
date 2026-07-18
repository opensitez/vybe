use super::helpers::run_lua_one;

#[test]
fn test_coroutine_yield_returns_baseline() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function(x)
  coroutine.yield(x)
  return x + 1, x + 2
end)
local ok1, first = coroutine.resume(co, 1)
local ok2, a, b = coroutine.resume(co)
print(ok1 and first == 1 and ok2 and a == 2 and b == 3)"#), "true");
}


#[test]
fn test_coroutine_yield_returns_simple() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function(x)
  coroutine.yield(x)
  return x + 1, x + 2
end)
local ok1, first = coroutine.resume(co, 2)
local ok2, a, b = coroutine.resume(co)
print(ok1 and first == 2 and ok2 and a == 3 and b == 4)"#), "true");
}


#[test]
fn test_coroutine_yield_returns_trimmed() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function(x)
  coroutine.yield(x)
  return x + 1, x + 2
end)
local ok1, first = coroutine.resume(co, 3)
local ok2, a, b = coroutine.resume(co)
print(ok1 and first == 3 and ok2 and a == 4 and b == 5)"#), "true");
}


#[test]
fn test_coroutine_yield_returns_decimal() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function(x)
  coroutine.yield(x)
  return x + 1, x + 2
end)
local ok1, first = coroutine.resume(co, 4)
local ok2, a, b = coroutine.resume(co)
print(ok1 and first == 4 and ok2 and a == 5 and b == 6)"#), "true");
}


#[test]
fn test_coroutine_yield_returns_hexed() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function(x)
  coroutine.yield(x)
  return x + 1, x + 2
end)
local ok1, first = coroutine.resume(co, 5)
local ok2, a, b = coroutine.resume(co)
print(ok1 and first == 5 and ok2 and a == 6 and b == 7)"#), "true");
}


#[test]
fn test_coroutine_yield_returns_prefixed() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function(x)
  coroutine.yield(x)
  return x + 1, x + 2
end)
local ok1, first = coroutine.resume(co, 6)
local ok2, a, b = coroutine.resume(co)
print(ok1 and first == 6 and ok2 and a == 7 and b == 8)"#), "true");
}


#[test]
fn test_coroutine_yield_returns_negative() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function(x)
  coroutine.yield(x)
  return x + 1, x + 2
end)
local ok1, first = coroutine.resume(co, 7)
local ok2, a, b = coroutine.resume(co)
print(ok1 and first == 7 and ok2 and a == 8 and b == 9)"#), "true");
}


#[test]
fn test_coroutine_yield_returns_rounded() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function(x)
  coroutine.yield(x)
  return x + 1, x + 2
end)
local ok1, first = coroutine.resume(co, 8)
local ok2, a, b = coroutine.resume(co)
print(ok1 and first == 8 and ok2 and a == 9 and b == 10)"#), "true");
}


#[test]
fn test_coroutine_yield_returns_offset() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function(x)
  coroutine.yield(x)
  return x + 1, x + 2
end)
local ok1, first = coroutine.resume(co, 9)
local ok2, a, b = coroutine.resume(co)
print(ok1 and first == 9 and ok2 and a == 10 and b == 11)"#), "true");
}


#[test]
fn test_coroutine_yield_returns_paired() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function(x)
  coroutine.yield(x)
  return x + 1, x + 2
end)
local ok1, first = coroutine.resume(co, 10)
local ok2, a, b = coroutine.resume(co)
print(ok1 and first == 10 and ok2 and a == 11 and b == 12)"#), "true");
}


#[test]
fn test_coroutine_yield_returns_nested() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function(x)
  coroutine.yield(x)
  return x + 1, x + 2
end)
local ok1, first = coroutine.resume(co, 11)
local ok2, a, b = coroutine.resume(co)
print(ok1 and first == 11 and ok2 and a == 12 and b == 13)"#), "true");
}


#[test]
fn test_coroutine_yield_returns_metaflow() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function(x)
  coroutine.yield(x)
  return x + 1, x + 2
end)
local ok1, first = coroutine.resume(co, 12)
local ok2, a, b = coroutine.resume(co)
print(ok1 and first == 12 and ok2 and a == 13 and b == 14)"#), "true");
}


#[test]
fn test_coroutine_yield_returns_guarded() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function(x)
  coroutine.yield(x)
  return x + 1, x + 2
end)
local ok1, first = coroutine.resume(co, 13)
local ok2, a, b = coroutine.resume(co)
print(ok1 and first == 13 and ok2 and a == 14 and b == 15)"#), "true");
}


#[test]
fn test_coroutine_yield_returns_mapped() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function(x)
  coroutine.yield(x)
  return x + 1, x + 2
end)
local ok1, first = coroutine.resume(co, 14)
local ok2, a, b = coroutine.resume(co)
print(ok1 and first == 14 and ok2 and a == 15 and b == 16)"#), "true");
}


#[test]
fn test_coroutine_yield_returns_captured() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function(x)
  coroutine.yield(x)
  return x + 1, x + 2
end)
local ok1, first = coroutine.resume(co, 15)
local ok2, a, b = coroutine.resume(co)
print(ok1 and first == 15 and ok2 and a == 16 and b == 17)"#), "true");
}


#[test]
fn test_coroutine_yield_returns_edge_first() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function(x)
  coroutine.yield(x)
  return x + 1, x + 2
end)
local ok1, first = coroutine.resume(co, 16)
local ok2, a, b = coroutine.resume(co)
print(ok1 and first == 16 and ok2 and a == 17 and b == 18)"#), "true");
}


#[test]
fn test_coroutine_yield_returns_edge_second() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function(x)
  coroutine.yield(x)
  return x + 1, x + 2
end)
local ok1, first = coroutine.resume(co, 17)
local ok2, a, b = coroutine.resume(co)
print(ok1 and first == 17 and ok2 and a == 18 and b == 19)"#), "true");
}


#[test]
fn test_coroutine_yield_returns_edge_last() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function(x)
  coroutine.yield(x)
  return x + 1, x + 2
end)
local ok1, first = coroutine.resume(co, 18)
local ok2, a, b = coroutine.resume(co)
print(ok1 and first == 18 and ok2 and a == 19 and b == 20)"#), "true");
}


#[test]
fn test_coroutine_yield_returns_randomized() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function(x)
  coroutine.yield(x)
  return x + 1, x + 2
end)
local ok1, first = coroutine.resume(co, 19)
local ok2, a, b = coroutine.resume(co)
print(ok1 and first == 19 and ok2 and a == 20 and b == 21)"#), "true");
}


#[test]
fn test_coroutine_yield_returns_unicode_like() {
    assert_eq!(run_lua_one(r#"local co = coroutine.create(function(x)
  coroutine.yield(x)
  return x + 1, x + 2
end)
local ok1, first = coroutine.resume(co, 20)
local ok2, a, b = coroutine.resume(co)
print(ok1 and first == 20 and ok2 and a == 21 and b == 22)"#), "true");
}
