-- vybe-test: lua/numeric_pow_boundaries/test_numeric_pow_boundaries_odd_power
-- origin: languages/lua/tests/lua/test_numeric_pow_boundaries.rs

local __w1 = "-8.0"
local __i = 0

do local __t = tostring((-2) ^ 3); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
