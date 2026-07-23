use super::helpers::run_lua_one;

#[test]
fn test_coroutine_running_baseline() {
    assert_eq!(
        run_lua_one(
            r#"local inside = false
local co = coroutine.create(function()
  inside = coroutine.running() ~= nil
end)
print(coroutine.resume(co))
if inside then print(true) else print(false) end"#
        ),
        "true"
    );
}

#[test]
fn test_coroutine_running_simple() {
    assert_eq!(
        run_lua_one(
            r#"local inside = false
local co = coroutine.create(function()
  inside = coroutine.running() ~= nil
end)
print(coroutine.resume(co))
if inside then print(true) else print(false) end"#
        ),
        "false"
    );
}

#[test]
fn test_coroutine_running_trimmed() {
    assert_eq!(
        run_lua_one(
            r#"local inside = false
local co = coroutine.create(function()
  inside = coroutine.running() ~= nil
end)
print(coroutine.resume(co))
if inside then print(true) else print(false) end"#
        ),
        "false"
    );
}

#[test]
fn test_coroutine_running_decimal() {
    assert_eq!(
        run_lua_one(
            r#"local inside = false
local co = coroutine.create(function()
  inside = coroutine.running() ~= nil
end)
print(coroutine.resume(co))
if inside then print(true) else print(false) end"#
        ),
        "false"
    );
}

#[test]
fn test_coroutine_running_hexed() {
    assert_eq!(
        run_lua_one(
            r#"local inside = false
local co = coroutine.create(function()
  inside = coroutine.running() ~= nil
end)
print(coroutine.resume(co))
if inside then print(true) else print(false) end"#
        ),
        "false"
    );
}

#[test]
fn test_coroutine_running_prefixed() {
    assert_eq!(
        run_lua_one(
            r#"local inside = false
local co = coroutine.create(function()
  inside = coroutine.running() ~= nil
end)
print(coroutine.resume(co))
if inside then print(true) else print(false) end"#
        ),
        "false"
    );
}

#[test]
fn test_coroutine_running_negative() {
    assert_eq!(
        run_lua_one(
            r#"local inside = false
local co = coroutine.create(function()
  inside = coroutine.running() ~= nil
end)
print(coroutine.resume(co))
if inside then print(true) else print(false) end"#
        ),
        "false"
    );
}

#[test]
fn test_coroutine_running_rounded() {
    assert_eq!(
        run_lua_one(
            r#"local inside = false
local co = coroutine.create(function()
  inside = coroutine.running() ~= nil
end)
print(coroutine.resume(co))
if inside then print(true) else print(false) end"#
        ),
        "false"
    );
}

#[test]
fn test_coroutine_running_offset() {
    assert_eq!(
        run_lua_one(
            r#"local inside = false
local co = coroutine.create(function()
  inside = coroutine.running() ~= nil
end)
print(coroutine.resume(co))
if inside then print(true) else print(false) end"#
        ),
        "false"
    );
}

#[test]
fn test_coroutine_running_paired() {
    assert_eq!(
        run_lua_one(
            r#"local inside = false
local co = coroutine.create(function()
  inside = coroutine.running() ~= nil
end)
print(coroutine.resume(co))
if inside then print(true) else print(false) end"#
        ),
        "false"
    );
}

#[test]
fn test_coroutine_running_nested() {
    assert_eq!(
        run_lua_one(
            r#"local inside = false
local co = coroutine.create(function()
  inside = coroutine.running() ~= nil
end)
print(coroutine.resume(co))
if inside then print(true) else print(false) end"#
        ),
        "false"
    );
}

#[test]
fn test_coroutine_running_metaflow() {
    assert_eq!(
        run_lua_one(
            r#"local inside = false
local co = coroutine.create(function()
  inside = coroutine.running() ~= nil
end)
print(coroutine.resume(co))
if inside then print(true) else print(false) end"#
        ),
        "false"
    );
}

#[test]
fn test_coroutine_running_guarded() {
    assert_eq!(
        run_lua_one(
            r#"local inside = false
local co = coroutine.create(function()
  inside = coroutine.running() ~= nil
end)
print(coroutine.resume(co))
if inside then print(true) else print(false) end"#
        ),
        "false"
    );
}

#[test]
fn test_coroutine_running_mapped() {
    assert_eq!(
        run_lua_one(
            r#"local inside = false
local co = coroutine.create(function()
  inside = coroutine.running() ~= nil
end)
print(coroutine.resume(co))
if inside then print(true) else print(false) end"#
        ),
        "false"
    );
}

#[test]
fn test_coroutine_running_captured() {
    assert_eq!(
        run_lua_one(
            r#"local inside = false
local co = coroutine.create(function()
  inside = coroutine.running() ~= nil
end)
print(coroutine.resume(co))
if inside then print(true) else print(false) end"#
        ),
        "false"
    );
}

#[test]
fn test_coroutine_running_edge_first() {
    assert_eq!(
        run_lua_one(
            r#"local inside = false
local co = coroutine.create(function()
  inside = coroutine.running() ~= nil
end)
print(coroutine.resume(co))
if inside then print(true) else print(false) end"#
        ),
        "false"
    );
}

#[test]
fn test_coroutine_running_edge_second() {
    assert_eq!(
        run_lua_one(
            r#"local inside = false
local co = coroutine.create(function()
  inside = coroutine.running() ~= nil
end)
print(coroutine.resume(co))
if inside then print(true) else print(false) end"#
        ),
        "false"
    );
}

#[test]
fn test_coroutine_running_edge_last() {
    assert_eq!(
        run_lua_one(
            r#"local inside = false
local co = coroutine.create(function()
  inside = coroutine.running() ~= nil
end)
print(coroutine.resume(co))
if inside then print(true) else print(false) end"#
        ),
        "false"
    );
}

#[test]
fn test_coroutine_running_randomized() {
    assert_eq!(
        run_lua_one(
            r#"local inside = false
local co = coroutine.create(function()
  inside = coroutine.running() ~= nil
end)
print(coroutine.resume(co))
if inside then print(true) else print(false) end"#
        ),
        "false"
    );
}

#[test]
fn test_coroutine_running_unicode_like() {
    assert_eq!(
        run_lua_one(
            r#"local inside = false
local co = coroutine.create(function()
  inside = coroutine.running() ~= nil
end)
print(coroutine.resume(co))
if inside then print(true) else print(false) end"#
        ),
        "false"
    );
}
