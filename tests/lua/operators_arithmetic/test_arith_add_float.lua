-- vybe-test: lua/operators_arithmetic/test_arith_add_float
-- origin: languages/lua/tests/lua/test_operators_arithmetic.rs

local __w1 = "12.5"
local __i = 0

do local __t = tostring(10.5 + 2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
