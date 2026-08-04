-- vybe-test: lua/math_advanced_exhaustive/math_adv_exh_acos_neg_one
-- origin: languages/lua/tests/lua/test_math_advanced_exhaustive.rs

local __w1 = "true"
local __i = 0

do local __t = tostring(math.abs(math.acos(-1) - math.pi) < 1e-10); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
