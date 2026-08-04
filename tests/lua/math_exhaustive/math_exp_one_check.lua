-- vybe-test: lua/math_exhaustive/math_exp_one_check
-- origin: languages/lua/tests/lua/test_math_exhaustive.rs

local __w1 = "true"
local __i = 0

do local __t = tostring(math.exp(1) > 2.71 and math.exp(1) < 2.72); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
