-- vybe-test: lua/metatables_fallback_inheritance/inherit_override
-- origin: languages/lua/tests/lua/test_metatables_fallback_inheritance.rs

local __w1 = "20"
local __i = 0

local parent = {val = 10}
local child = setmetatable({val = 20}, {__index = parent})
do local __t = tostring(child.val); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
