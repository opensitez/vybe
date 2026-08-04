-- vybe-test: lua/math_misc/test_math_floor_positive
-- origin: languages/lua/tests/lua/test_math_misc.rs

local __w1 = "10"
local __i = 0

do local __t = tostring(math.floor(10.75)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
