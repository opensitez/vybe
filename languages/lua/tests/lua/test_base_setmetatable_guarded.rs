use super::helpers::run_lua_one;

#[test]
fn newindex_guard_rejects_missing_key_assignment() {
    assert_eq!(run_lua_one(r#"local t = {}
setmetatable(t, {__newindex = function() error("guard") end})
local ok = pcall(function() t.x = 1 end)
print(ok)"#), "false");
}

#[test]
fn newindex_guard_error_message_is_visible_to_pcall() {
    assert_eq!(run_lua_one(r#"local t = {}
setmetatable(t, {__newindex = function() error("guard") end})
local ok, err = pcall(function() t.x = 1 end)
print(tostring(ok) .. " " .. tostring(string.find(tostring(err), "guard") ~= nil))"#), "false true");
}

#[test]
fn rawset_bypasses_newindex_guard() {
    assert_eq!(run_lua_one(r#"local t = {}
setmetatable(t, {__newindex = function() error("guard") end})
rawset(t, "x", 1)
print(t.x)"#), "1");
}

#[test]
fn existing_key_assignment_bypasses_newindex_guard() {
    assert_eq!(run_lua_one(r#"local t = {x = 1}
setmetatable(t, {__newindex = function() error("guard") end})
t.x = 2
print(t.x)"#), "2");
}

#[test]
fn nil_assignment_to_missing_key_invokes_newindex_guard() {
    assert_eq!(run_lua_one(r#"local t = {}
setmetatable(t, {__newindex = function(tbl, key, value)
  if value == nil then rawset(tbl, key, "was_nil") end
end})
t.x = nil
print(t.x)"#), "was_nil");
}
