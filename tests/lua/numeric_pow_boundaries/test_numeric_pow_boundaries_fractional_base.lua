-- vybe-test: lua/numeric_pow_boundaries/test_numeric_pow_boundaries_fractional_base
-- origin: languages/lua/tests/lua/test_numeric_pow_boundaries.rs

local __w1 = "20.25"
local __i = 0

do local __t = tostring(4.5 ^ 2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
