-- vybe-test: lua/base_setmetatable_guarded/existing_key_assignment_bypasses_newindex_guard
-- origin: languages/lua/tests/lua/test_base_setmetatable_guarded.rs

local __w1 = "2"
local __i = 0

local t = {x = 1}
setmetatable(t, {__newindex = function() error("guard") end})
t.x = 2
do local __t = tostring(t.x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
