use super::helpers::run_lua_one;

#[test]
fn test_getmetatable_presence_baseline() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
local mt = {__tostring = function() return "ok" end}
setmetatable(t, mt)
print(getmetatable(t) == mt)"#
        ),
        "true"
    );
}

#[test]
fn test_getmetatable_presence_simple() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
print(getmetatable(t) == nil)"#
        ),
        "true"
    );
}

#[test]
fn test_getmetatable_presence_trimmed() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
local mt = {__tostring = function() return "ok" end}
setmetatable(t, mt)
print(getmetatable(t) == mt)"#
        ),
        "true"
    );
}

#[test]
fn test_getmetatable_presence_decimal() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
print(getmetatable(t) == nil)"#
        ),
        "true"
    );
}

#[test]
fn test_getmetatable_presence_hexed() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
local mt = {__tostring = function() return "ok" end}
setmetatable(t, mt)
print(getmetatable(t) == mt)"#
        ),
        "true"
    );
}

#[test]
fn test_getmetatable_presence_prefixed() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
print(getmetatable(t) == nil)"#
        ),
        "true"
    );
}

#[test]
fn test_getmetatable_presence_negative() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
local mt = {__tostring = function() return "ok" end}
setmetatable(t, mt)
print(getmetatable(t) == mt)"#
        ),
        "true"
    );
}

#[test]
fn test_getmetatable_presence_rounded() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
print(getmetatable(t) == nil)"#
        ),
        "true"
    );
}

#[test]
fn test_getmetatable_presence_offset() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
local mt = {__tostring = function() return "ok" end}
setmetatable(t, mt)
print(getmetatable(t) == mt)"#
        ),
        "true"
    );
}

#[test]
fn test_getmetatable_presence_paired() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
print(getmetatable(t) == nil)"#
        ),
        "true"
    );
}

#[test]
fn test_getmetatable_presence_nested() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
local mt = {__tostring = function() return "ok" end}
setmetatable(t, mt)
print(getmetatable(t) == mt)"#
        ),
        "true"
    );
}

#[test]
fn test_getmetatable_presence_metaflow() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
print(getmetatable(t) == nil)"#
        ),
        "true"
    );
}

#[test]
fn test_getmetatable_presence_guarded() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
local mt = {__tostring = function() return "ok" end}
setmetatable(t, mt)
print(getmetatable(t) == mt)"#
        ),
        "true"
    );
}

#[test]
fn test_getmetatable_presence_mapped() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
print(getmetatable(t) == nil)"#
        ),
        "true"
    );
}

#[test]
fn test_getmetatable_presence_captured() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
local mt = {__tostring = function() return "ok" end}
setmetatable(t, mt)
print(getmetatable(t) == mt)"#
        ),
        "true"
    );
}

#[test]
fn test_getmetatable_presence_edge_first() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
print(getmetatable(t) == nil)"#
        ),
        "true"
    );
}

#[test]
fn test_getmetatable_presence_edge_second() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
local mt = {__tostring = function() return "ok" end}
setmetatable(t, mt)
print(getmetatable(t) == mt)"#
        ),
        "true"
    );
}

#[test]
fn test_getmetatable_presence_edge_last() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
print(getmetatable(t) == nil)"#
        ),
        "true"
    );
}

#[test]
fn test_getmetatable_presence_randomized() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
local mt = {__tostring = function() return "ok" end}
setmetatable(t, mt)
print(getmetatable(t) == mt)"#
        ),
        "true"
    );
}

#[test]
fn test_getmetatable_presence_unicode_like() {
    assert_eq!(
        run_lua_one(
            r#"local t = {}
print(getmetatable(t) == nil)"#
        ),
        "true"
    );
}
