-- vybe-test: lua/math_misc/test_math_max_basic
-- origin: languages/lua/tests/lua/test_math_misc.rs

local __w1 = "20"
local __i = 0

do local __t = tostring(math.max(10, 5, 20)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
