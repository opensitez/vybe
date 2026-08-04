-- vybe-test: lua/math_exhaustive/math_ceil_neg_flt_1
-- origin: languages/lua/tests/lua/test_math_exhaustive.rs

local __w1 = "-10"
local __i = 0

do local __t = tostring(math.ceil(-10.1)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
