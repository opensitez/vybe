-- vybe-test: lua/metatables_math/test_math_sub_tables
-- origin: languages/lua/tests/lua/test_metatables_math.rs

local __w1 = "30"
local __i = 0

local mt={__sub=function(a,b) return a.v-b.v end}; local t1=setmetatable({v=50}, mt); local t2=setmetatable({v=20}, mt); do local __t = tostring(t1-t2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
