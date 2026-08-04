-- vybe-test: lua/nan_and_inf/infinity_minus_infinity_is_nan
-- origin: languages/lua/tests/lua/test_nan_and_inf.rs

local __w1 = "true"
local __i = 0

local x = 1/0 - 1/0
do local __t = tostring(x ~= x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
