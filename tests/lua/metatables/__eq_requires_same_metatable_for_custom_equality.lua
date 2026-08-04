-- vybe-test: lua/metatables/__eq_requires_same_metatable_for_custom_equality
-- origin: languages/lua/tests/lua/test_metatables.rs

local __w1 = "true"
local __i = 0

local mt = {__eq = function(a, b) return a.id == b.id end}
do local __t = tostring(setmetatable({id=1}, mt) == setmetatable({id=1}, mt)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
