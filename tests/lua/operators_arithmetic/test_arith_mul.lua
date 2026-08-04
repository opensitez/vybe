-- vybe-test: lua/operators_arithmetic/test_arith_mul
-- origin: languages/lua/tests/lua/test_operators_arithmetic.rs

local __w1 = "30"
local __i = 0

do local __t = tostring(10 * 3); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
