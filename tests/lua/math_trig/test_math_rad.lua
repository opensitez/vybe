-- vybe-test: lua/math_trig/test_math_rad
-- origin: languages/lua/tests/lua/test_math_trig.rs

local __w1 = "314"
local __i = 0

do local __t = tostring(math.floor(math.rad(180) * 100)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
