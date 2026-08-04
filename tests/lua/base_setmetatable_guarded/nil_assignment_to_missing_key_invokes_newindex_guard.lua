-- vybe-test: lua/base_setmetatable_guarded/nil_assignment_to_missing_key_invokes_newindex_guard
-- origin: languages/lua/tests/lua/test_base_setmetatable_guarded.rs

local __w1 = "was_nil"
local __i = 0

local t = {}
setmetatable(t, {__newindex = function(tbl, key, value)
  if value == nil then rawset(tbl, key, "was_nil") end
end})
t.x = nil
do local __t = tostring(t.x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
