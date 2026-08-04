-- vybe-test: lua/math_misc/test_math_modf_positive
-- origin: languages/lua/tests/lua/test_math_misc.rs

local __w1 = "10 0.25"
local __i = 0

local i, f = math.modf(10.25); do local __t = tostring(i..' '..f); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
