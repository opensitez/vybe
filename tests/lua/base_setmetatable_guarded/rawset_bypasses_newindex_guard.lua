-- vybe-test: lua/base_setmetatable_guarded/rawset_bypasses_newindex_guard
-- origin: languages/lua/tests/lua/test_base_setmetatable_guarded.rs

local __w1 = "1"
local __i = 0

local t = {}
setmetatable(t, {__newindex = function() error("guard") end})
rawset(t, "x", 1)
do local __t = tostring(t.x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
