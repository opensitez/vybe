-- vybe-test: lua/metatables_math/test_math_different_metatables_right_priority
-- origin: languages/lua/tests/lua/test_metatables_math.rs

local __w1 = "2"
local __i = 0

local mt1={}; local mt2={__add=function() return 2 end}; local t1=setmetatable({}, mt1); local t2=setmetatable({}, mt2); do local __t = tostring(t1+t2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
