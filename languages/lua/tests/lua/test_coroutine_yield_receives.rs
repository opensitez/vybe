use super::helpers::run_lua_one;

#[test]
fn test_coroutine_yield_receives_baseline() {
    assert_eq!(
        run_lua_one(
            r#"local c = coroutine.create(function(x)
  local y = coroutine.yield(x)
  return y == x * 2
end)
local _, first = coroutine.resume(c, 1)
local _, second = coroutine.resume(c, 2)
print(first == 1 and second == true)"#
        ),
        "true"
    );
}

#[test]
fn test_coroutine_yield_receives_simple() {
    assert_eq!(
        run_lua_one(
            r#"local c = coroutine.create(function(x)
  local y = coroutine.yield(x)
  return y == x * 2
end)
local _, first = coroutine.resume(c, 2)
local _, second = coroutine.resume(c, 4)
print(first == 2 and second == true)"#
        ),
        "true"
    );
}

#[test]
fn test_coroutine_yield_receives_trimmed() {
    assert_eq!(
        run_lua_one(
            r#"local c = coroutine.create(function(x)
  local y = coroutine.yield(x)
  return y == x * 2
end)
local _, first = coroutine.resume(c, 3)
local _, second = coroutine.resume(c, 6)
print(first == 3 and second == true)"#
        ),
        "true"
    );
}

#[test]
fn test_coroutine_yield_receives_decimal() {
    assert_eq!(
        run_lua_one(
            r#"local c = coroutine.create(function(x)
  local y = coroutine.yield(x)
  return y == x * 2
end)
local _, first = coroutine.resume(c, 4)
local _, second = coroutine.resume(c, 8)
print(first == 4 and second == true)"#
        ),
        "true"
    );
}

#[test]
fn test_coroutine_yield_receives_hexed() {
    assert_eq!(
        run_lua_one(
            r#"local c = coroutine.create(function(x)
  local y = coroutine.yield(x)
  return y == x * 2
end)
local _, first = coroutine.resume(c, 5)
local _, second = coroutine.resume(c, 10)
print(first == 5 and second == true)"#
        ),
        "true"
    );
}

#[test]
fn test_coroutine_yield_receives_prefixed() {
    assert_eq!(
        run_lua_one(
            r#"local c = coroutine.create(function(x)
  local y = coroutine.yield(x)
  return y == x * 2
end)
local _, first = coroutine.resume(c, 6)
local _, second = coroutine.resume(c, 12)
print(first == 6 and second == true)"#
        ),
        "true"
    );
}

#[test]
fn test_coroutine_yield_receives_negative() {
    assert_eq!(
        run_lua_one(
            r#"local c = coroutine.create(function(x)
  local y = coroutine.yield(x)
  return y == x * 2
end)
local _, first = coroutine.resume(c, 7)
local _, second = coroutine.resume(c, 14)
print(first == 7 and second == true)"#
        ),
        "true"
    );
}

#[test]
fn test_coroutine_yield_receives_rounded() {
    assert_eq!(
        run_lua_one(
            r#"local c = coroutine.create(function(x)
  local y = coroutine.yield(x)
  return y == x * 2
end)
local _, first = coroutine.resume(c, 8)
local _, second = coroutine.resume(c, 16)
print(first == 8 and second == true)"#
        ),
        "true"
    );
}

#[test]
fn test_coroutine_yield_receives_offset() {
    assert_eq!(
        run_lua_one(
            r#"local c = coroutine.create(function(x)
  local y = coroutine.yield(x)
  return y == x * 2
end)
local _, first = coroutine.resume(c, 9)
local _, second = coroutine.resume(c, 18)
print(first == 9 and second == true)"#
        ),
        "true"
    );
}

#[test]
fn test_coroutine_yield_receives_paired() {
    assert_eq!(
        run_lua_one(
            r#"local c = coroutine.create(function(x)
  local y = coroutine.yield(x)
  return y == x * 2
end)
local _, first = coroutine.resume(c, 10)
local _, second = coroutine.resume(c, 20)
print(first == 10 and second == true)"#
        ),
        "true"
    );
}

#[test]
fn test_coroutine_yield_receives_nested() {
    assert_eq!(
        run_lua_one(
            r#"local c = coroutine.create(function(x)
  local y = coroutine.yield(x)
  return y == x * 2
end)
local _, first = coroutine.resume(c, 11)
local _, second = coroutine.resume(c, 22)
print(first == 11 and second == true)"#
        ),
        "true"
    );
}

#[test]
fn test_coroutine_yield_receives_metaflow() {
    assert_eq!(
        run_lua_one(
            r#"local c = coroutine.create(function(x)
  local y = coroutine.yield(x)
  return y == x * 2
end)
local _, first = coroutine.resume(c, 12)
local _, second = coroutine.resume(c, 24)
print(first == 12 and second == true)"#
        ),
        "true"
    );
}

#[test]
fn test_coroutine_yield_receives_guarded() {
    assert_eq!(
        run_lua_one(
            r#"local c = coroutine.create(function(x)
  local y = coroutine.yield(x)
  return y == x * 2
end)
local _, first = coroutine.resume(c, 13)
local _, second = coroutine.resume(c, 26)
print(first == 13 and second == true)"#
        ),
        "true"
    );
}

#[test]
fn test_coroutine_yield_receives_mapped() {
    assert_eq!(
        run_lua_one(
            r#"local c = coroutine.create(function(x)
  local y = coroutine.yield(x)
  return y == x * 2
end)
local _, first = coroutine.resume(c, 14)
local _, second = coroutine.resume(c, 28)
print(first == 14 and second == true)"#
        ),
        "true"
    );
}

#[test]
fn test_coroutine_yield_receives_captured() {
    assert_eq!(
        run_lua_one(
            r#"local c = coroutine.create(function(x)
  local y = coroutine.yield(x)
  return y == x * 2
end)
local _, first = coroutine.resume(c, 15)
local _, second = coroutine.resume(c, 30)
print(first == 15 and second == true)"#
        ),
        "true"
    );
}

#[test]
fn test_coroutine_yield_receives_edge_first() {
    assert_eq!(
        run_lua_one(
            r#"local c = coroutine.create(function(x)
  local y = coroutine.yield(x)
  return y == x * 2
end)
local _, first = coroutine.resume(c, 16)
local _, second = coroutine.resume(c, 32)
print(first == 16 and second == true)"#
        ),
        "true"
    );
}

#[test]
fn test_coroutine_yield_receives_edge_second() {
    assert_eq!(
        run_lua_one(
            r#"local c = coroutine.create(function(x)
  local y = coroutine.yield(x)
  return y == x * 2
end)
local _, first = coroutine.resume(c, 17)
local _, second = coroutine.resume(c, 34)
print(first == 17 and second == true)"#
        ),
        "true"
    );
}

#[test]
fn test_coroutine_yield_receives_edge_last() {
    assert_eq!(
        run_lua_one(
            r#"local c = coroutine.create(function(x)
  local y = coroutine.yield(x)
  return y == x * 2
end)
local _, first = coroutine.resume(c, 18)
local _, second = coroutine.resume(c, 36)
print(first == 18 and second == true)"#
        ),
        "true"
    );
}

#[test]
fn test_coroutine_yield_receives_randomized() {
    assert_eq!(
        run_lua_one(
            r#"local c = coroutine.create(function(x)
  local y = coroutine.yield(x)
  return y == x * 2
end)
local _, first = coroutine.resume(c, 19)
local _, second = coroutine.resume(c, 38)
print(first == 19 and second == true)"#
        ),
        "true"
    );
}

#[test]
fn test_coroutine_yield_receives_unicode_like() {
    assert_eq!(
        run_lua_one(
            r#"local c = coroutine.create(function(x)
  local y = coroutine.yield(x)
  return y == x * 2
end)
local _, first = coroutine.resume(c, 20)
local _, second = coroutine.resume(c, 40)
print(first == 20 and second == true)"#
        ),
        "true"
    );
}
