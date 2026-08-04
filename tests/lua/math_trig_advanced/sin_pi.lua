-- vybe-test: lua/math_trig_advanced/sin_pi
-- origin: languages/lua/tests/lua/test_math_trig_advanced.rs

local __w1 = "true"
local __i = 0

do local __t = tostring(math.abs(math.sin(math.pi)) < 1e-14); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
