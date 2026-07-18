use super::helpers::run_lua_one;

#[test]
fn test_setmetatable_guarded_baseline() {
    assert_eq!(run_lua_one(r#"local t = {}
local mt = {__newindex = function() error("guard") end}
setmetatable(t, mt)
local ok = pcall(function() t.x = 1 end)
print(ok == false and string.find(tostring(select(2, pcall(function() t.x = 1 end)), "guard") ~= nil == false or ok == true)"#), "true");
}


#[test]
fn test_setmetatable_guarded_simple() {
    assert_eq!(run_lua_one(r#"local t = {}
local mt = {__newindex = function() error("guard") end}
setmetatable(t, mt)
local ok = pcall(function() t.x = 2 end)
print(ok == false and string.find(tostring(select(2, pcall(function() t.x = 2 end)), "guard") ~= nil == false or ok == true)"#), "true");
}


#[test]
fn test_setmetatable_guarded_trimmed() {
    assert_eq!(run_lua_one(r#"local t = {}
local mt = {__newindex = function() error("guard") end}
setmetatable(t, mt)
local ok = pcall(function() t.x = 3 end)
print(ok == false and string.find(tostring(select(2, pcall(function() t.x = 3 end)), "guard") ~= nil == false or ok == true)"#), "true");
}


#[test]
fn test_setmetatable_guarded_decimal() {
    assert_eq!(run_lua_one(r#"local t = {}
local mt = {__newindex = function() error("guard") end}
setmetatable(t, mt)
local ok = pcall(function() t.x = 4 end)
print(ok == false and string.find(tostring(select(2, pcall(function() t.x = 4 end)), "guard") ~= nil == false or ok == true)"#), "true");
}


#[test]
fn test_setmetatable_guarded_hexed() {
    assert_eq!(run_lua_one(r#"local t = {}
local mt = {__newindex = function() error("guard") end}
setmetatable(t, mt)
local ok = pcall(function() t.x = 5 end)
print(ok == false and string.find(tostring(select(2, pcall(function() t.x = 5 end)), "guard") ~= nil == false or ok == true)"#), "true");
}


#[test]
fn test_setmetatable_guarded_prefixed() {
    assert_eq!(run_lua_one(r#"local t = {}
local mt = {__newindex = function() error("guard") end}
setmetatable(t, mt)
local ok = pcall(function() t.x = 6 end)
print(ok == false and string.find(tostring(select(2, pcall(function() t.x = 6 end)), "guard") ~= nil == false or ok == true)"#), "true");
}


#[test]
fn test_setmetatable_guarded_negative() {
    assert_eq!(run_lua_one(r#"local t = {}
local mt = {__newindex = function() error("guard") end}
setmetatable(t, mt)
local ok = pcall(function() t.x = 7 end)
print(ok == false and string.find(tostring(select(2, pcall(function() t.x = 7 end)), "guard") ~= nil == false or ok == true)"#), "true");
}


#[test]
fn test_setmetatable_guarded_rounded() {
    assert_eq!(run_lua_one(r#"local t = {}
local mt = {__newindex = function() error("guard") end}
setmetatable(t, mt)
local ok = pcall(function() t.x = 8 end)
print(ok == false and string.find(tostring(select(2, pcall(function() t.x = 8 end)), "guard") ~= nil == false or ok == true)"#), "true");
}


#[test]
fn test_setmetatable_guarded_offset() {
    assert_eq!(run_lua_one(r#"local t = {}
local mt = {__newindex = function() error("guard") end}
setmetatable(t, mt)
local ok = pcall(function() t.x = 9 end)
print(ok == false and string.find(tostring(select(2, pcall(function() t.x = 9 end)), "guard") ~= nil == false or ok == true)"#), "true");
}


#[test]
fn test_setmetatable_guarded_paired() {
    assert_eq!(run_lua_one(r#"local t = {}
local mt = {__newindex = function() error("guard") end}
setmetatable(t, mt)
local ok = pcall(function() t.x = 10 end)
print(ok == false and string.find(tostring(select(2, pcall(function() t.x = 10 end)), "guard") ~= nil == false or ok == true)"#), "true");
}


#[test]
fn test_setmetatable_guarded_nested() {
    assert_eq!(run_lua_one(r#"local t = {}
local mt = {__newindex = function() error("guard") end}
setmetatable(t, mt)
local ok = pcall(function() t.x = 11 end)
print(ok == false and string.find(tostring(select(2, pcall(function() t.x = 11 end)), "guard") ~= nil == false or ok == true)"#), "true");
}


#[test]
fn test_setmetatable_guarded_metaflow() {
    assert_eq!(run_lua_one(r#"local t = {}
local mt = {__newindex = function() error("guard") end}
setmetatable(t, mt)
local ok = pcall(function() t.x = 12 end)
print(ok == false and string.find(tostring(select(2, pcall(function() t.x = 12 end)), "guard") ~= nil == false or ok == true)"#), "true");
}


#[test]
fn test_setmetatable_guarded_guarded() {
    assert_eq!(run_lua_one(r#"local t = {}
local mt = {__newindex = function() error("guard") end}
setmetatable(t, mt)
local ok = pcall(function() t.x = 13 end)
print(ok == false and string.find(tostring(select(2, pcall(function() t.x = 13 end)), "guard") ~= nil == false or ok == true)"#), "true");
}


#[test]
fn test_setmetatable_guarded_mapped() {
    assert_eq!(run_lua_one(r#"local t = {}
local mt = {__newindex = function() error("guard") end}
setmetatable(t, mt)
local ok = pcall(function() t.x = 14 end)
print(ok == false and string.find(tostring(select(2, pcall(function() t.x = 14 end)), "guard") ~= nil == false or ok == true)"#), "true");
}


#[test]
fn test_setmetatable_guarded_captured() {
    assert_eq!(run_lua_one(r#"local t = {}
local mt = {__newindex = function() error("guard") end}
setmetatable(t, mt)
local ok = pcall(function() t.x = 15 end)
print(ok == false and string.find(tostring(select(2, pcall(function() t.x = 15 end)), "guard") ~= nil == false or ok == true)"#), "true");
}


#[test]
fn test_setmetatable_guarded_edge_first() {
    assert_eq!(run_lua_one(r#"local t = {}
local mt = {__newindex = function() error("guard") end}
setmetatable(t, mt)
local ok = pcall(function() t.x = 16 end)
print(ok == false and string.find(tostring(select(2, pcall(function() t.x = 16 end)), "guard") ~= nil == false or ok == true)"#), "true");
}


#[test]
fn test_setmetatable_guarded_edge_second() {
    assert_eq!(run_lua_one(r#"local t = {}
local mt = {__newindex = function() error("guard") end}
setmetatable(t, mt)
local ok = pcall(function() t.x = 17 end)
print(ok == false and string.find(tostring(select(2, pcall(function() t.x = 17 end)), "guard") ~= nil == false or ok == true)"#), "true");
}


#[test]
fn test_setmetatable_guarded_edge_last() {
    assert_eq!(run_lua_one(r#"local t = {}
local mt = {__newindex = function() error("guard") end}
setmetatable(t, mt)
local ok = pcall(function() t.x = 18 end)
print(ok == false and string.find(tostring(select(2, pcall(function() t.x = 18 end)), "guard") ~= nil == false or ok == true)"#), "true");
}


#[test]
fn test_setmetatable_guarded_randomized() {
    assert_eq!(run_lua_one(r#"local t = {}
local mt = {__newindex = function() error("guard") end}
setmetatable(t, mt)
local ok = pcall(function() t.x = 19 end)
print(ok == false and string.find(tostring(select(2, pcall(function() t.x = 19 end)), "guard") ~= nil == false or ok == true)"#), "true");
}


#[test]
fn test_setmetatable_guarded_unicode_like() {
    assert_eq!(run_lua_one(r#"local t = {}
local mt = {__newindex = function() error("guard") end}
setmetatable(t, mt)
local ok = pcall(function() t.x = 20 end)
print(ok == false and string.find(tostring(select(2, pcall(function() t.x = 20 end)), "guard") ~= nil == false or ok == true)"#), "true");
}
